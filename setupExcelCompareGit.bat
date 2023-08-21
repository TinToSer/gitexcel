@echo off
Rem Developed by TinToSer(https://github.com/tintoser)

set gitExcelBatchFile=gitExcel.cmd

echo "===========Setting Program Files============="
set gitExcelFolder=C:\gitExcel
set gitExcelBatchPath=%gitExcelFolder%\%gitExcelBatchFile%
mkdir %gitExcelFolder%
copy /y %gitExcelBatchFile% %gitExcelBatchPath%

echo "===========Injecting to global .gitconfig ============="
set gitConfigPath=%homedrive%%homepath%\.gitconfig
set diffCmd=[diff "gitExcel"]
for /f "tokens=*" %%a in (%gitConfigPath%) do (
   if "%diffCmd%"=="%%a" goto nochange
)
echo %diffCmd%>> %gitConfigPath%
set compareCMD=    command = %gitExcelBatchPath:\=/%
echo %compareCMD%>> %gitConfigPath%
echo Modification injected ...
exit
:nochange
echo .gitconfig was already modified ...
