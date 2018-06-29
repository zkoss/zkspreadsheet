#!/bin/bash
# Purpose:
# build zk spreadsheet with maven. Because the whole process requires many steps, therefore I write it as script instead of configuring in jenkins

flBundleGoals='-P build-fl clean source:jar javadoc:jar repository:bundle-create -Dmaven.test.skip=true'
removeSnapshot='versions:set -DremoveSnapshot'

function buildZpoi(){
    # a relative path based on the current path
    zpoiPom='zsspoi/zpoi/'
    # remove '-SNAPSHOT' from project version
    mvn -f ${zpoiPom} versions:set -DremoveSnapshot
    # http://maven.apache.org/plugins/maven-repository-plugin/usage.html
    mvn -B -f ${zpoiPom} ${flBundleGoals}
    mvn -B -f ${zpoiPom} install -Dmaven.test.skip=true
}

function buildZssmodel(){
    zssmodelPom='zkspreadsheet/zssmodel/'
    mvn -f ${zssmodelPom} versions:update-property -Dproperty=zpoi.version -DnewVersion=${zpoiFlVersion}
    mvn -f ${zssmodelPom} versions:set -DremoveSnapshot
    mvn -B -f ${zssmodelPom} ${flBundleGoals}
}

function buildZss(){
    zssPom='zkspreadsheet/zss/'
    mvn -f ${zssPom} versions:set -DremoveSnapshot
    mvn -B -f ${zssPom} ${flBundleGoals}
}

buildZpoi
# get zpoi fl version
zpoiFlVersion=$(mvn -f ${zpoiPom} help:evaluate -Dexpression=project.version | egrep -v '\[|Downloading:' | tr -d '\n')
echo zpoi fl version: ${zpoiFlVersion}
buildZssmodel
buildZss