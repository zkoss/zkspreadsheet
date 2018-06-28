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
    # remove SNAPSHOT
    mvn -f ${zssmodelPom} versions:update-property -Dproperty=zpoi.version
    # update to FL version
    mvn -f ${zssmodelPom} versions:update-property -Dproperty=zpoi.version
    mvn -f ${zssmodelPom} versions:set -DremoveSnapshot
    mvn -B -f ${zssmodelPom} ${flBundleGoals}
}

function buildZss(){
    zssPom='zkspreadsheet/zss/'
    # remove SNAPSHOT
    mvn -f ${zssmodelPom} versions:update-property -Dproperty=zpoi.version    
    # update zpoi version to the latest one
    mvn -f ${zssPom} versions:set -DremoveSnapshot
    mvn -f ${zssPom} versions:update-property -Dproperty=zpoi.version
    mvn -B -f ${zssPom} ${flBundleGoals}
}

buildZpoi
buildZssmodel
buildZss