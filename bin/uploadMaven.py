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


ZSS_BUNDLE_FILE_PATH = '/zss/maven/FL/' #FIXME not for proprietary
ZPOI_BUNDLE_FILE_PATH = '/zpoi/maven/FL/'

LOCAL_RELEASE_PATH = "/tmp"
REMOTE_RELEASE_PATH = "//10.1.3.252/potix/rd"
MOUNTED_RELEASE_PATH = LOCAL_RELEASE_PATH+"/potix-rd/"

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
	freshlyVersionRegExpr = re.compile("(\d\.)+FL\.(\d)+")
	#officialVersionRe =
	
	print("Get freshly version")
	#find zpoi version
	global ZPOI_VERSION 
	ZPOI_VERSION = freshlyVersionRegExpr.search(os.listdir(getBundleFilePath(zpoiList[0]))[0]).group(0)
	print("zpoi version: "+ZPOI_VERSION)
	#find zss version
	global ZSS_VERSION 
	ZSS_VERSION = freshlyVersionRegExpr.search(os.listdir(getBundleFilePath(zssList[0]))[0]).group(0)
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
		return LOCAL_RELEASE_PATH+ZSS_BUNDLE_FILE_PATH
	elif projectName in zpoiList:
		return LOCAL_RELEASE_PATH+ZPOI_BUNDLE_FILE_PATH


initProjectVersion()
mountRemoteFolder()
createArtifactFolder()
uploadMavenBundle()