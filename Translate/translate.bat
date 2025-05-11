@echo off
setlocal enabledelayedexpansion

rem ▼ 引数確認（置換CSVパスとフォルダパス）
if "%~2"=="" (
    echo [使い方] translate.bat <置換CSVファイルパス、ファイル名も含む> <変換対象フォルダパス>
    exit /b 1
)

set CSV_PATH=%~1
set FOLDER_PATH=%~2

rem ▼ クラスパス作成（lib フォルダの JAR を全て含める）
set CP=.
for %%i in (lib\*.jar) do (
    set CP=!CP!;lib\%%i
)

rem ▼ コンパイル
echo 🔧 Javaファイルをコンパイル中...
javac -cp "!CP!" src\main\java\com\example\TranslateWithFont.java
if errorlevel 1 (
    echo コンパイル失敗
    exit /b 1
)

rem ▼ 実行
echo 実行中...
java -cp "!CP!;src\main\java" com.example.TranslateWithFont "!CSV_PATH!" "!FOLDER_PATH!"
