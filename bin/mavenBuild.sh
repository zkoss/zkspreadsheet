#!/bin/bash
# Purpose:
# build zk spreadsheet with maven. Because the whole process requires many steps, therefore I write it as script instead of configuring in jenkins

flBundleGoals='clean source:jar javadoc:jar repository:bundle-create -Dmaven.test.skip=true'
# a relative path based on the current path
zpoiPom='zsspoi/zpoi/pom.xml'
zssmodelPom='zkspreadsheet/zssmodel/'
zssPom='zkspreadsheet/zss/'
zpoiexPom='zsscml/zpoiex/'
zssexPom='zsscml/zssex/'
zssjsfPom='zsscml/zssjsf/'
zssjspPom='zsscml/zssjsp/'
zsspdfPom='zsscml/zsspdf/'
zsshtmlPom='zsscml/zsshtml/'

zpoiFlVersion='unset'


function buildBundleInstall(){
    buildBundle $1
    # for dependent artifact to resolve correctly
    mvn -f $1 install -Dmaven.test.skip=true
}

# build a maven bundle file
function buildBundle(){
    mvn -f $1 versions:set -DremoveSnapshot #remove '-SNAPSHOT' from project version
    mvn -f $1 -P build-fl validate # set freshly version
    # http://maven.apache.org/plugins/maven-repository-plugin/usage.html
    mvn -B -f $1 ${flBundleGoals}
}

function setZpoiVersion(){
    sed -i.bak "s/zpoi.version>.*<\/zpoi.version>/zpoi.version>${zpoiFlVersion}<\/zpoi.version>/" $1pom.xml
}

buildBundleInstall ${zpoiPom}
# get zpoi fl version
# get project version with a maven plugin sometimes requires a download and failed to obtain the version
# zpoiFlVersion=$(mvn -f ${zpoiPom} help:evaluate -Dexpression=project.version | egrep -v '\[|Downloading:' | tr -d '\n')
zpoiFlVersion=$(python zkspreadsheet/bin/read_pom_version.py ${zpoiPom})
echo zpoi freshly version: ${zpoiFlVersion}
# versions:use-latest-releases, version:update-property both failed set zpoi version in Jenkins
setZpoiVersion ${zssmodelPom}
buildBundleInstall ${zssmodelPom}
buildBundleInstall ${zssPom}
buildBundleInstall ${zpoiexPom}
setZpoiVersion ${zssexPom}
buildBundleInstall ${zssexPom}
buildBundle ${zsspdfPom}
buildBundle ${zssjsfPom}
buildBundle ${zssjspPom}
buildBundle ${zsshtmlPom}