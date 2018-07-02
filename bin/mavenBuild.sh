#!/bin/bash
# Purpose:
# build zk spreadsheet with maven. Because the whole process requires many steps, therefore I write it as script instead of configuring in jenkins

flBundleGoals='clean source:jar javadoc:jar repository:bundle-create -Dmaven.test.skip=true'
removeSnapshot='versions:set -DremoveSnapshot'
# a relative path based on the current path
zpoiPom='zsspoi/zpoi/'
zssmodelPom='zkspreadsheet/zssmodel/'
zssPom='zkspreadsheet/zss/'

function buildZpoi(){
    # remove '-SNAPSHOT' from project version
    mvn -f ${zpoiPom} versions:set -DremoveSnapshot
    mvn -f ${zpoiPom} -P build-fl validate
    # http://maven.apache.org/plugins/maven-repository-plugin/usage.html
    mvn -B -f ${zpoiPom} ${flBundleGoals}
    mvn -B -f ${zpoiPom} install -Dmaven.test.skip=true
}

function buildZssmodel(){
    # get zpoi fl version
    zpoiFlVersion=$(mvn -f ${zpoiPom} help:evaluate -Dexpression=project.version | egrep -v '\[|Downloading:' | tr -d '\n')
    echo zpoi freshly version: ${zpoiFlVersion}
    # versions:use-latest-releases, version:update-property both failed in Jenkins
    sed -i.bak "s/zpoi.version>.*<\/zpoi.version>/zpoi.version>${zpoiFlVersion}<\/zpoi.version>/" ${zssmodelPom}pom.xml

    mvn -f ${zssmodelPom} versions:set -DremoveSnapshot
    mvn -f ${zssmodelPom} -P build-fl validate
    mvn -B -f ${zssmodelPom} ${flBundleGoals}
    # for zss to resolve
    mvn -f ${zssmodelPom} install -Dmaven.test.skip=true
}

function buildZss(){
    #remove snapshot
    mvn -f ${zssPom} versions:set -DremoveSnapshot
    # set FL version
    mvn -f ${zssPom} -P build-fl validate
    mvn -B -f ${zssPom} clean source:jar javadoc:jar repository:bundle-create -Dmaven.test.skip=true
    mvn -f ${zssPom} install -Dmaven.test.skip=true
}

buildZpoi
buildZssmodel
buildZss