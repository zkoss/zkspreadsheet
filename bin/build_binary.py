#!/usr/bin/env python
# build binary release file

import subprocess
import read_pom_version

MVN = 'mvn'
PHASE = 'process-resources'
PROFILE_EE = '-P ee'
PROFILE_EVAL = '-P eval'
PROFILE_OSE = '-P ose'
POM = 'pom.xml'
VERSION = '3.9.5'

def create_ee_zip():
    read_pom_version.modifyVersion(POM, VERSION)
    subprocess.check_call([MVN, PROFILE_EE, PHASE])


def create_ose_zip():
    subprocess.check_call([MVN, PROFILE_OSE, PHASE])

def create_ee_eval_zip():
    read_pom_version.modifyVersion(POM, VERSION+"-Eval")
    subprocess.check_call([MVN, PROFILE_EVAL, PHASE])


subprocess.check_call([MVN, "clean"])
create_ee_zip()
create_ose_zip()
create_ee_eval_zip()