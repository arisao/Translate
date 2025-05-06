#!/bin/bash

# Mavenを使ってTranslateWithFontを実行
mvn exec:java -Dexec.mainClass="TranslateWithFont" \
  -Dexec.args="resources/replace_rules.csv /Users/xxx/Documents/test"