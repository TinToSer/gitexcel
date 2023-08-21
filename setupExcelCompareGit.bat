@echo off
Rem Developed by TinToSer(https://github.com/tintoser)

set gitExcelBatchFile=gitExcel.bat

echo "===========Setting Program Files============="
set gitExcelFolder=C:\gitExcel
set gitExcelBatchPath=%gitExcelFolder%\%gitExcelBatchFile%
mkdir %gitExcelFolder%
copy /y %gitExcelBatchFile% %gitExcelBatchPath%

echo "===========Injecting to .git/config ============="
set gitConfigPath=%homedrive%%homepath%\.gitconfig
echo [diff "gitExcel"] >> %gitConfigPath%
set compareCMD=    command = %gitExcelBatchPath:\=/%
echo %compareCMD% >> %gitConfigPath%
