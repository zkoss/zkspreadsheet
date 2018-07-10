# A script to read a project's version in pom.xml

import xml.etree.ElementTree as ET
import os
import logging
import sys

def findProjectVersion(pom_file):
    tree = ET.parse(pom_file)
    root = tree.getroot()
    for child in root:
        if child.tag.find('version') > -1:
            return child.text


def validate(pom_file):
    return pom_file.endswith("pom.xml")


def main():
    logger = logging.getLogger( sys.argv[0])
    logging.basicConfig(level='INFO')
    pom_file = sys.argv[1]
    if not validate(pom_file):
        logger.error('not a pom.xml')
        return 2
    version = findProjectVersion(pom_file)
    print version

if __name__== "__main__":
  main()        