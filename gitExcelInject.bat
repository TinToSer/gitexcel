@echo off
Rem Developed by TinToSer(https://github.com/tintoser)

echo "===========Injecting to .git/config ============="

set /p gitPath=Enter Repository Path:-

echo *.xls* diff=gitExcel >> %gitPath%\.gitattributes