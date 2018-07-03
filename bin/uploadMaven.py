# A script to upload ZSS Maven bundle files to POTIX release file server
# Bundle Location for Maven CE: \\fileserver\potix\rd\[project]\releases\[version]\maven\ 
# Bundles Location for Maven EE-eval: \\fileserver\potix\rd\[project]\releases\[version]\maven\EE-eval\ 
# Bundle file format: [project]-[version]-bundle.* 
# could be refined to a MavenUploader which knows their version, project list, bundle files, 
# target path respectively and just create 2 instances with different parameters and ask them to upload.  

import xml.etree.ElementTree as ET
import os
import subprocess
import logging
import shutil
import ConfigParser

# find zss, zpoi version from pom.xml
def findProjectVersion(pomFilePath):
    tree = ET.parse(pomFilePath)
    root = tree.getroot()
    for child in root:
        if child.tag.find('version') > -1:
            return child.text

ZSS_MAVEN_PATH = '/zss/maven/' 
ZPOI_MAVEN_PATH = '/zpoi/maven/'
LOCAL_RELEASE_PATH = "/tmp"
DESTINATION_PATH_JENKINS = "/media/potix/rd/"
REMOTE_RELEASE_PATH = "//guest@10.1.3.252/potix/rd" #fileserver
MOUNTED_RELEASE_PATH = LOCAL_RELEASE_PATH + "/potix-rd/"
destination_path = DESTINATION_PATH_JENKINS

# mount the ZK release file vault on the file server after removing previous one
# no need to mount the folder on jenkins
def mountRemoteFolder():
    if (os.path.ismount(MOUNTED_RELEASE_PATH)):
        return
    # subprocess.check_call(["umount", MOUNTED_RELEASE_PATH])

    if (os.path.exists(MOUNTED_RELEASE_PATH)):
        os.rmdir(MOUNTED_RELEASE_PATH)
    os.mkdir(MOUNTED_RELEASE_PATH)
    subprocess.check_call(["mount_smbfs", "-N", REMOTE_RELEASE_PATH, MOUNTED_RELEASE_PATH])
    destination_path = MOUNTED_RELEASE_PATH


ZSS_PROJECT_LIST = ['zss','zssmodel', 'zssex', 'zssjsf','zssjsp','zsspdf']
ZPOI_PROJECT_LIST = ['zpoi', 'zpoiex']

# create folders in file server
def createArtifactFolder():
    for project in ZSS_PROJECT_LIST + ZPOI_PROJECT_LIST:
        createFolderIfNotExist(getBundleFileTargetFolder(project))
    global isFreshlyVersion
    if (not isFreshlyVersion):
        for project in ZSS_PROJECT_LIST + ZPOI_PROJECT_LIST:
            createFolderIfNotExist(getProprietaryFileTargetFolder(project))


def createFolderIfNotExist(path):
    if os.path.exists(path):
        logger.info(path + ' existed')
    else:
        os.makedirs(path)
        logger.info('created folder: '+path)


def getBundleFileTargetFolder(projectName):
    return destination_path+projectName+'/releases/'+getProjectVersion(projectName)+'/maven/EE-Eval'


def getProprietaryFileTargetFolder(projectName):
    return destination_path+projectName+'/releases/'+getProjectVersion(projectName)+'/maven/proprietary/EE'


def getProjectVersion(projectName):
    if projectName in ZSS_PROJECT_LIST:
        return zss_version
    elif projectName in ZPOI_PROJECT_LIST:
        return zpoi_version


# copy to potix fileserver server \ release folder
def copyMavenBundle():
    for projectName in ZSS_PROJECT_LIST + ZPOI_PROJECT_LIST:
        bundleFileName = projectName+"-"+getProjectVersion(projectName)+"-bundle.jar"
        sourceBundleFile = getLocalBundleFolder(projectName) + bundleFileName
        destinationFolder = getBundleFileTargetFolder(projectName)
        if os.path.exists(sourceBundleFile):
            shutil.copyfile(getLocalBundleFolder(projectName)+bundleFileName, destinationFolder+"/"+bundleFileName)
            logger.info("copied "+sourceBundleFile + " to " + destinationFolder)

    #copy proprietary bundle files into zss project maven folder
    global isFreshlyVersion
    if (not isFreshlyVersion):
        for projectName in ZSS_PROJECT_LIST+ZPOI_PROJECT_LIST:
            bundleFileName = projectName+"-"+getProjectVersion(projectName)+"-bundle.jar"
            shutil.copyfile(getProprietaryFilePath("projectName")+bundleFileName, getProprietaryFileTargetFolder("zss")+"/"+bundleFileName)
            logger.info("copied "+bundleFileName )


def getProprietaryFilePath(projectName):
    if projectName in ZSS_PROJECT_LIST:
        return LOCAL_RELEASE_PATH+ZSS_MAVEN_PATH+'proprietary/'
    elif projectName in ZPOI_PROJECT_LIST:
        return LOCAL_RELEASE_PATH+ZPOI_MAVEN_PATH+'proprietary/'


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
    }
    return os.path.join(PROJECT_PATH[projectName], projectName, 'target/')


# create a version properties file as parameters for jenkins to run the next job
def createVersionProperties():
    Config = ConfigParser.ConfigParser()
    properties_file = open("version.properties",'w')
    Config.add_section('version')
    Config.set('version','zss_version',zss_version)
    Config.set('version','zpoi_version',zpoi_version)
    Config.set('version','maven', 'ee-eval')
    Config.write(properties_file)
    properties_file.close()


logger = logging.getLogger('uploadMaven')
logging.basicConfig(level='INFO')

isFreshlyVersion = True
zss_version = findProjectVersion('zkspreadsheet/zss/pom.xml')
zpoi_version = findProjectVersion('zsspoi/zpoi/pom.xml')


if not os.path.exists(DESTINATION_PATH_JENKINS):
    mountRemoteFolder()
logger.info("destination: " + destination_path)
createArtifactFolder()
copyMavenBundle()
createVersionProperties()