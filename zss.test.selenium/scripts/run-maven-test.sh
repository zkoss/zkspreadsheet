#!/bin/bash
echo "run maven clean"
mvn clean
echo "run maven test-compile"
mvn test-compile
echo "run maven surefire:test"
echo "params=\"$1\" \"$2\" \"$3\" \"$4\" \"$5\""
mvn -X surefire:test -Dreport.dir="$1" -Dtest.dir="$2" -Dbrowser="$3" -Dtest="$4" -Dserver="$5"