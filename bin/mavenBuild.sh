#!/bin/bash
# Purpose:
# build zk spreadsheet with maven. Because the whole process requires many steps, therefore I write it as script instead of configuring in jenkins

# Usage:
# run the script on top of 3 folders checked out from 3 repositories (zsspoi, zkspreadsheet, zsscml)

# Command parameters:
# no parameter: freshly
# official: official EE
# eval: EE evaluation

bundleGoals='clean source:jar javadoc:jar repository:bundle-create -Dmaven.test.skip=true'
evalBundleGoals='clean repository:bundle-create -Dmaven.test.skip=true'
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

zpoiVersion='unset'


function buildBundleInstall(){
   buildBundle $1 $2
    # install an artifact for resolving dependencies correctly
    mvn -f $1 install -Dmaven.test.skip=true
}

# build a maven bundle file
function buildBundle(){
    mvn -B -f $1 versions:set -DremoveSnapshot #remove '-SNAPSHOT' from project version
    if [[ $edition = "official" ]]
    then
        mvn -B -f $1 -P $edition ${bundleGoals}
    else
        mvn -f $1 -P $edition validate # set freshly version
        # http://maven.apache.org/plugins/maven-repository-plugin/usage.html
        mvn -B -f $1 ${evalBundleGoals}
    fi
}


function setZpoiVersion(){
    sed -i.bak "s/zpoi.version>.*<\/zpoi.version>/zpoi.version>${zpoiVersion}<\/zpoi.version>/" $1pom.xml
}

edition=$1
if [[ "official" = $edition ]] || [[ "eval" = $edition ]]
then
    echo "=== Build $edition ==="
else
    echo "=== Build freshly ==="
    edition="freshly"
fi

buildBundleInstall ${zpoiPom}

# get project version with a maven plugin sometimes requires a download and failed to obtain the version
# zpoiFlVersion=$(mvn -f ${zpoiPom} help:evaluate -Dexpression=project.version | egrep -v '\[|Downloading:' | tr -d '\n')
zpoiVersion=$(python zkspreadsheet/bin/read_pom_version.py ${zpoiPom})
echo zpoi version: ${zpoiVersion}


# versions:use-latest-releases, version:update-property both failed to set zpoi version in Jenkins
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
python3 zkspreadsheet/bin/uploadMaven.py $edition