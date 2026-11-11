#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
mkdir -p out
javac -encoding UTF-8 -d out *.java
java -cp out SpreadsheetApp
