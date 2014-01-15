# A script to upload ZSS Maven bundle files to POTIX release file server
# Written in Python 3
# Usage:
# 	1. copy this file under /tmp, the same folder with those ready-to-release files
#	2. Run it with python 3: python3 uploadMaven.py

import subprocess
import os
import re
import shutil

ZSS_VERSION = None 
ZPOI_VERSION = None

ZSS_BUNDLE_FILE_PATH = 'zss/maven/FL/' #FIXME not for proprietary
ZPOI_BUNDLE_FILE_PATH = 'zpoi/maven/FL/'

REMOTE_RELEASE_PATH = "//10.1.3.252/potix/rd"
MOUNTED_RELEASE_PATH = "/tmp/potix-rd/"
zssList = ['zss', 'zssex', 'zsshtml', 'zssjsf','zssjsp','zsspdf']
zpoiList = ['zpoi', 'zpoiex']




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
	freshlyVersionRe = re.compile("(\d\.)+FL\.(\d)+")
	#officialVersionRe =
	
	print("Get freshly version")
	#find zpoi version
	global ZPOI_VERSION 
	ZPOI_VERSION = freshlyVersionRe.search(os.listdir("zpoi/maven/FL")[0]).group(0)
	print("zpoi version: "+ZPOI_VERSION)
	#find zss version
	global ZSS_VERSION 
	ZSS_VERSION = freshlyVersionRe.search(os.listdir("zss/maven/FL")[0]).group(0)
	print("zss version: "+ZSS_VERSION)


# upload to potix release server
def uploadMavenBundle():
	for projectName in zssList+zpoiList:
		bundleFileName = projectName+"-"+getProjectVersion(projectName)+"-bundle.jar"
		shutil.copyfile(getBundleFilePath(projectName)+bundleFileName, getBundleFileTargetFolder(projectName)+"/"+bundleFileName)
		print("copied "+bundleFileName )
		
		
def createArtifactFolder():
	for project in zssList+zpoiList:
		createFolderIfNotExist(getBundleFileTargetFolder(project))
			
def createFolderIfNotExist(path):
	if (not os.path.exists(path)):
		os.makedirs(path)
		print('created folder: '+path)

def getBundleFileTargetFolder(projectName):
	return MOUNTED_RELEASE_PATH+projectName+'/releases/'+getProjectVersion(projectName)+'/maven'

def getProjectVersion(projectName):
	if projectName in zssList:
		return ZSS_VERSION
	elif projectName in zpoiList:
		return ZPOI_VERSION

def getBundleFilePath(projectName):
	if projectName in zssList:
		return ZSS_BUNDLE_FILE_PATH
	elif projectName in zpoiList:
		return ZPOI_BUNDLE_FILE_PATH


initProjectVersion()
mountRemoteFolder()
createArtifactFolder()
uploadMavenBundle()