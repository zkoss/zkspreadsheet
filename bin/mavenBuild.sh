#!/bin/bash
# Purpose:
# build zk spreadsheet with maven. Because the whole process requires many steps, therefore I write it as script instead of configuring in jenkins

flBundleGoals='clean source:jar javadoc:jar repository:bundle-create -Dmaven.test.skip=true'
# a relative path based on the current path
zpoiPom='zsspoi/zpoi/'
zssmodelPom='zkspreadsheet/zssmodel/'
zssPom='zkspreadsheet/zss/'
zpoiexPom='zsscml/zpoiex/'
zssexPom='zsscml/zssex/'
zssjsfPom='zsscml/zssjsf/'
zssjspPom='zsscml/zssjsp/'
zsspdfPom='zsscml/zsspdf/'

zpoiFlVersion='unset'

function buildZpoi(){
    buildBundle ${zpoiPom}
    mvn -B -f ${zpoiPom} install -Dmaven.test.skip=true
}

function buildZssmodel(){
    # get zpoi fl version
    zpoiFlVersion=$(mvn -f ${zpoiPom} help:evaluate -Dexpression=project.version | egrep -v '\[|Downloading:' | tr -d '\n')
    echo zpoi freshly version: ${zpoiFlVersion}
    # versions:use-latest-releases, version:update-property both failed in Jenkins
    sed -i.bak "s/zpoi.version>.*<\/zpoi.version>/zpoi.version>${zpoiFlVersion}<\/zpoi.version>/" ${zssmodelPom}pom.xml

    buildBundle ${zssmodelPom}
    # for zss to resolve
    mvn -f ${zssmodelPom} install -Dmaven.test.skip=true
}

function buildZss(){
    buildBundle ${zssPom}
    mvn -f ${zssPom} install -Dmaven.test.skip=true
}

function buildZpoiex(){
    buildBundle ${zpoiexPom}
    mvn -B -f ${zpoiexPom} install -Dmaven.test.skip=true
}

function buildZssex(){
    sed -i.bak "s/zpoi.version>.*<\/zpoi.version>/zpoi.version>${zpoiFlVersion}<\/zpoi.version>/" ${zssexPom}pom.xml

    buildBundle ${zssexPom}
    # for zssjsp to resolve
    mvn -f ${zssexPom} install -Dmaven.test.skip=true
}

# build a maven bundle file
function buildBundle(){
    mvn -f $1 versions:set -DremoveSnapshot #remove '-SNAPSHOT' from project version
    mvn -f $1 -P build-fl validate # set freshly version
    # http://maven.apache.org/plugins/maven-repository-plugin/usage.html
    mvn -B -f $1 ${flBundleGoals}
}

buildZpoi
buildZssmodel
buildZss
buildZpoiex
buildZssex
buildBundle ${zsspdfPom}
buildBundle ${zssjsfPom}
buildBundle ${zssjspPom}