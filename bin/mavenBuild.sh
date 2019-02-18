#!/bin/bash
# Purpose:
# build zk spreadsheet with maven. Because the whole process requires many steps, therefore I write it as script instead of configuring in jenkins

bundleGoals='clean source:jar javadoc:jar repository:bundle-create -Dmaven.test.skip=true'
# a relative path based on the current path
zpoiPom='zsspoi/zpoi/pom.xml'
zssParent='zkspreadsheet/zssparent/'
zssmodelPom='zkspreadsheet/zssmodel/'
zssPom='zkspreadsheet/zss/'

zsscmlParent='zsscml/zsscmlparent/'
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
    mvn -f $1 versions:set -DremoveSnapshot #remove '-SNAPSHOT' from project version
    if [[ $2 = "official" ]]
    then
        mvn -B -f $1 -P official ${bundleGoals}
    else
        mvn -f $1 -P build-fl validate # set freshly version
        # http://maven.apache.org/plugins/maven-repository-plugin/usage.html
        mvn -B -f $1 ${bundleGoals}
    fi
}


function setZpoiVersion(){
    sed -i.bak "s/zpoi.version>.*<\/zpoi.version>/zpoi.version>${zpoiVersion}<\/zpoi.version>/" $1pom.xml
}


if [[ "official" = $1 ]]
then
    echo "=== Build official ==="
else
    echo "=== Build freshly ==="
fi

buildBundleInstall ${zpoiPom} $1

# get zpoi fl version
# get project version with a maven plugin sometimes requires a download and failed to obtain the version
# zpoiFlVersion=$(mvn -f ${zpoiPom} help:evaluate -Dexpression=project.version | egrep -v '\[|Downloading:' | tr -d '\n')
zpoiVersion=$(python zkspreadsheet/bin/read_pom_version.py ${zpoiPom})
echo zpoi version: ${zpoiVersion}

mvn -f ${zssParent} install # install parent pom so that child project can resolve during building
# versions:use-latest-releases, version:update-property both failed set zpoi version in Jenkins
setZpoiVersion ${zssmodelPom}
buildBundleInstall ${zssmodelPom} $1
buildBundleInstall ${zssPom} $1


mvn -f ${zsscmlParent} install # install parent pom so that child project can resolve during building
buildBundleInstall ${zpoiexPom} $1
setZpoiVersion ${zssexPom}
buildBundleInstall ${zssexPom} $1
buildBundle ${zsspdfPom} $1
buildBundle ${zssjsfPom} $1
buildBundle ${zssjspPom} $1
buildBundle ${zsshtmlPom} $1