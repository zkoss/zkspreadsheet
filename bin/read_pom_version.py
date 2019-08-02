#!/usr/bin/env python
# A script to read/write a project's version in pom.xml

import xml.etree.ElementTree as ElementTree
import os
import logging
import sys

logger = logging.getLogger( sys.argv[0])

def findProjectVersion(pom_file):
    tree = ElementTree.parse(pom_file)
    root = tree.getroot()
    for child in root:
        if child.tag.find('version') > -1:
            return child.text


def modifyVersion(pom_file, version):
    namespaces = {'' : 'http://maven.apache.org/POM/4.0.0'}
    for prefix, uri in namespaces.items():
        ElementTree.register_namespace(prefix, uri)
    ns = '{http://maven.apache.org/POM/4.0.0}'
    tree = ElementTree.parse(pom_file)
    root = tree.getroot()
    root.find(ns+"version").text = version
    tree.write(pom_file)
    logger.info("set version " + (root.find(ns+"version").text))


def validate(pom_file):
    return pom_file.endswith("pom.xml")


def main():
    logging.basicConfig(level='INFO')
    pom_file = sys.argv[1]
    if not validate(pom_file):
        logger.error('not a pom.xml')
        return 2
    version = findProjectVersion(pom_file)
    print(version)


if __name__== "__main__":
  main()        