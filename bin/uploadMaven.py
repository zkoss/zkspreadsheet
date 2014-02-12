# A script to upload ZSS Maven bundle files to POTIX release file server
# Written in Python 3
# Usage:
#	Run it with python 3: python3 uploadMaven.py
# could be refined to a MavenUploader which knows their version, project list, bundle files, 
# target path respectively and just create 2 instances with different parameters and ask them to upload.  

import subprocess
import os
import re
import shutil

ZSS_VERSION = None 
ZPOI_VERSION = None

isFreshlyVersion = True


ZSS_MAVEN_PATH = '/zss/maven/' 
ZPOI_MAVEN_PATH = '/zpoi/maven/'


LOCAL_RELEASE_PATH = "/tmp"
REMOTE_RELEASE_PATH = "//10.1.3.252/potix/rd"
MOUNTED_RELEASE_PATH = LOCAL_RELEASE_PATH+"/potix-rd/"

ZSS_PROJECT_LIST = ['zss', 'zssex', 'zsshtml', 'zssjsf','zssjsp','zsspdf']
ZPOI_PROJECT_LIST = ['zpoi', 'zpoiex']




# mount the ZK release file vault on the file server after removing previous one
def mountRemoteFolder():
	if (os.path.ismount(MOUNTED_RELEASE_PATH)):
		subprocess.check_call(["umount", MOUNTED_RELEASE_PATH])
	if (os.path.exists(MOUNTED_RELEASE_PATH)):
		os.rmdir(MOUNTED_RELEASE_PATH)
	os.mkdir(MOUNTED_RELEASE_PATH);
	subprocess.check_call(["mount", REMOTE_RELEASE_PATH, MOUNTED_RELEASE_PATH])

# initialize zss, zpoi version from released file names
def initProjectVersion():
	#officialVersionRe = 
	global isFreshlyVersion
	if os.path.exists(LOCAL_RELEASE_PATH+ZSS_MAVEN_PATH+'proprietary'):
		isFreshlyVersion = False
		print("Official version")
	else:
		isFreshlyVersion = True
		print("Freshly version")
	zssBundleFileName = os.listdir(getBundleFilePath(ZSS_PROJECT_LIST[0]))[0]
	zpoiBundleFileName = os.listdir(getBundleFilePath(ZPOI_PROJECT_LIST[0]))[0]
	global ZPOI_VERSION 
	global ZSS_VERSION 
	if (isFreshlyVersion):
		freshlyVersionRegExpr = re.compile("(\d\.)+FL\.(\d)+")
		ZPOI_VERSION = freshlyVersionRegExpr.search(zpoiBundleFileName).group(0)
		print("zpoi version: "+ZPOI_VERSION)
		ZSS_VERSION = freshlyVersionRegExpr.search(zssBundleFileName).group(0)
		print("zss version: "+ZSS_VERSION)
	else:
		officialVersionRegExpr = re.compile("\d\.(\d)+\.(\d)+")
		ZPOI_VERSION = officialVersionRegExpr.search(zpoiBundleFileName).group(0)
		print("zpoi version: "+ZPOI_VERSION)
		ZSS_VERSION = officialVersionRegExpr.search(zssBundleFileName).group(0)
		print("zss version: "+ZSS_VERSION)
		

# upload to potix release server
def uploadMavenBundle():
	for projectName in ZSS_PROJECT_LIST+ZPOI_PROJECT_LIST:
		bundleFileName = projectName+"-"+getProjectVersion(projectName)+"-bundle.jar"
		shutil.copyfile(getBundleFilePath(projectName)+bundleFileName, getBundleFileTargetFolder(projectName)+"/"+bundleFileName)
		print("copied "+bundleFileName )
	#copy proprietary bundle files 
	global isFreshlyVersion
	if (not isFreshlyVersion):
		for projectName in ZSS_PROJECT_LIST+ZPOI_PROJECT_LIST:
			bundleFileName = projectName+"-"+getProjectVersion(projectName)+"-bundle.jar"
			shutil.copyfile(getProprietaryFilePath(projectName)+bundleFileName, getProprietaryFileTargetFolder(projectName)+"/"+bundleFileName)
			print("copied "+bundleFileName )
		
def createArtifactFolder():
	for project in ZSS_PROJECT_LIST+ZPOI_PROJECT_LIST:
		createFolderIfNotExist(getBundleFileTargetFolder(project))
	global isFreshlyVersion
	if (not isFreshlyVersion):
		for project in ZSS_PROJECT_LIST+ZPOI_PROJECT_LIST:
			createFolderIfNotExist(getProprietaryFileTargetFolder(project))
			
def createFolderIfNotExist(path):
	if (not os.path.exists(path)):
		os.makedirs(path)
		print('created folder: '+path)

def getBundleFileTargetFolder(projectName):
	return MOUNTED_RELEASE_PATH+projectName+'/releases/'+getProjectVersion(projectName)+'/maven'

def getProprietaryFileTargetFolder(projectName):
	return MOUNTED_RELEASE_PATH+projectName+'/releases/'+getProjectVersion(projectName)+'/maven/proprietary'

def getProjectVersion(projectName):
	if projectName in ZSS_PROJECT_LIST:
		return ZSS_VERSION
	elif projectName in ZPOI_PROJECT_LIST:
		return ZPOI_VERSION

def getBundleFilePath(projectName):
	global isFreshlyVersion
	if (isFreshlyVersion):
		bundleFilePath = 'FL/'
	else:
		bundleFilePath = 'eval/'
	if projectName in ZSS_PROJECT_LIST:
		return LOCAL_RELEASE_PATH+ZSS_MAVEN_PATH+bundleFilePath
	elif projectName in ZPOI_PROJECT_LIST:
		return LOCAL_RELEASE_PATH+ZPOI_MAVEN_PATH+bundleFilePath

def getProprietaryFilePath(projectName):
	if projectName in ZSS_PROJECT_LIST:
		return LOCAL_RELEASE_PATH+ZSS_MAVEN_PATH+'proprietary/'
	elif projectName in ZPOI_PROJECT_LIST:
		return LOCAL_RELEASE_PATH+ZPOI_MAVEN_PATH+'proprietary/'

initProjectVersion()
mountRemoteFolder()
createArtifactFolder()
uploadMavenBundle()