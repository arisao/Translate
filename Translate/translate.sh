#!/bin/bash

SRC_DIR="src/main/java"
OUT_DIR="out"
MAIN_CLASS="com.example.TranslateWithFont"
LIB_DIR="lib"

# クラスパスにすべてのjarを含める
JARS=$(find "$LIB_DIR" -name "*.jar" | paste -sd ":" -)
CP="${OUT_DIR}:${JARS}"

echo "コンパイル中..."
mkdir -p "${OUT_DIR}"
javac -cp "${JARS}" -d "${OUT_DIR}" ${SRC_DIR}/com/example/TranslateWithFont.java

if [ $? -ne 0 ]; then
  echo "コンパイルエラー"
  exit 1
fi

echo "実行中..."
java -cp "${CP}" "${MAIN_CLASS}" "$1" "$2"