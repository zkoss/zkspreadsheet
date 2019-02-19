# A script to upload ZSS Maven bundle files to POTIX file server.
# If bundle file is freshly release (determined by version string), copy to EE-eval folder
# If bundle file is official release, copy to EE folder
# Bundle Location:
# * Maven CE: \\fileserver\potix\rd\[project]\releases\[version]\maven\
# * Maven EE-eval: \\fileserver\potix\rd\[project]\releases\[version]\maven\EE-eval\ 
# * Maven EE: \\fileserver\potix\rd\zss\releases\$Version\maven\proprietary\EE\ 
# Bundle file naming: [project]-[version]-bundle.* 

# could be refined to a MavenUploader which knows their version, project list, bundle files, 
# target path respectively and just create 2 instances with different parameters and ask them to upload.  

import xml.etree.ElementTree as ET
import os
import subprocess
import logging
import shutil
import configparser

# find zss, zpoi version from pom.xml
def findProjectVersion(pomFilePath):
    tree = ET.parse(pomFilePath)
    root = tree.getroot()
    for child in root:
        if child.tag.find('version') > -1:
            return child.text

ZSS_MAVEN_PATH = '/zss/maven/' 
ZPOI_MAVEN_PATH = '/zpoi/maven/'
REMOTE_RELEASE_PATH = "//guest@10.1.3.252/potix/rd/" #fileserver
# mount at /tmp can avoid permission denied
MOUNTED_RELEASE_PATH = "/tmp/zss-release/"
DESTINATION_PATH_JENKINS = "/media/potix/rd/" # the path to file server at Jenkins
destination_path = DESTINATION_PATH_JENKINS

# mount the ZK release file vault on the file server
# no need to mount the folder on jenkins
def mountRemoteFolder():
    # run on Jenkins
    if os.path.exists(DESTINATION_PATH_JENKINS):
       return

    if (os.path.ismount(MOUNTED_RELEASE_PATH)):
        return
    # subprocess.check_call(["umount", MOUNTED_RELEASE_PATH])

    if (os.path.exists(MOUNTED_RELEASE_PATH)):
        os.rmdir(MOUNTED_RELEASE_PATH)
    os.mkdir(MOUNTED_RELEASE_PATH)
    subprocess.check_call(["mount_smbfs", "-N", REMOTE_RELEASE_PATH, MOUNTED_RELEASE_PATH])
    global  destination_path
    destination_path = MOUNTED_RELEASE_PATH
    logger.info("destination: " + destination_path)


ZSS_PROJECT_LIST = ['zss','zssmodel', 'zssex', 'zssjsf','zssjsp','zsspdf', 'zsshtml']
ZPOI_PROJECT_LIST = ['zpoi', 'zpoiex']

# create folders in file server
# all bundles in ZSS_PROJECT_LIST => zss/releases/[version]/maven/
# all bundles in ZPOI_PROJECT_LIST => zpoi/releases/[version]/maven/
def createDestinationFolder():
    createFolderIfNotExist(getBundleFileTargetFolder('zss')) # for ZSS_PROJECT_LIST
    createFolderIfNotExist(getBundleFileTargetFolder('zpoi')) # for ZPOI_PROJECT_LIST


def createFolderIfNotExist(path):
    if os.path.exists(path):
        logger.info(path + ' existed')
    else:
        os.makedirs(path)
        logger.info('created folder: '+path)


# get target folder for EE-Eval and EE
def getBundleFileTargetFolder(projectName):
    if (projectName in ZSS_PROJECT_LIST):
        project_folder = 'zss'
    else:
        project_folder = 'zpoi'
    if (isEval()):
        return destination_path + project_folder +'/releases/'+getProjectVersion(projectName)+'/maven/EE-Eval'
    else:
        return destination_path + project_folder + '/releases/' + getProjectVersion(projectName) + '/maven/proprietary/EE'


def getProjectVersion(projectName):
    if projectName in ZSS_PROJECT_LIST:
        return zss_version
    elif projectName in ZPOI_PROJECT_LIST:
        return zpoi_version


# copy to Potix file server / release folder
def copyMavenBundle():
    for projectName in ZSS_PROJECT_LIST + ZPOI_PROJECT_LIST:
        bundle_file_name = projectName+"-"+getProjectVersion(projectName)+"-bundle.jar"
        sourceBundleFile = getLocalBundleFolder(projectName) + bundle_file_name
        destinationFolder = getBundleFileTargetFolder(projectName)
        if os.path.exists(sourceBundleFile):
            shutil.copyfile(getLocalBundleFolder(projectName)+bundle_file_name, destinationFolder+"/"+bundle_file_name)
            logger.info("copied "+sourceBundleFile + " to " + destinationFolder)


# get local bundle file path
def getLocalBundleFolder(projectName):
    PROJECT_PATH = {
        'zpoi'      : 'zsspoi',
        'zss'       : 'zkspreadsheet',
        'zssmodel'  : 'zkspreadsheet',
        'zssex'     : 'zsscml',
        'zpoiex'    : 'zsscml',
        'zssjsf'    : 'zsscml',
        'zssjsp'    : 'zsscml',
        'zsspdf'    : 'zsscml',
        'zsshtml'   : 'zsscml',
    }
    return os.path.join(PROJECT_PATH[projectName], projectName, 'target/')


def isEval():
    return "Eval" in zss_version

# create a version properties file as parameters for jenkins to run the next job
def createVersionProperties():
    Config = configparser.ConfigParser()
    properties_file = open("version.properties",'w')
    Config.add_section('version')
    Config.set('version','zss_version',zss_version)
    Config.set('version','zpoi_version',zpoi_version)
    if (isEval()):
        Config.set('version','maven', 'ee-eval')
    else:
        Config.set('version','maven', 'ee')
    Config.write(properties_file)
    properties_file.close()


logger = logging.getLogger('uploadMaven')
logging.basicConfig(level='INFO')

zss_version = findProjectVersion('zkspreadsheet/zss/pom.xml')
zpoi_version = findProjectVersion('zsspoi/zpoi/pom.xml')

mountRemoteFolder()
createDestinationFolder()
copyMavenBundle()
createVersionProperties()