@echo off
Rem Developed by TinToSer(https://github.com/tintoser)

set gitExcelBatchFile=gitExcel.bat

echo "===========Setting Program Files============="
set gitExcelFolder=C:\gitExcel
set gitExcelBatchPath=%gitExcelFolder%\%gitExcelBatchFile%
mkdir %gitExcelFolder%
copy /y %gitExcelBatchFile% %gitExcelBatchPath%

echo "===========Injecting to .git/config ============="
set /p gitPath=Enter Repository Path:-

set gitConfigPath=%gitPath%\.git\config
echo [diff "gitExcel"] >> %gitConfigPath%
set compareCMD=    command = %gitExcelBatchPath:\=/%
echo %compareCMD% >> %gitConfigPath%
echo *.xls* diff=gitExcel >> %gitPath%\.gitattributes
