@echo off
Rem Developed by TinToSer(https://github.com/tintoser)

set /p gitPath=Enter Repository Path:-
set gitAttrFile=%gitPath%\.gitattributes
echo "===========Injecting to .gitattributes ============="
set attrCmd=*.xls* diff=gitExcel
for /f "tokens=*" %%a in (%gitAttrFile%) do (
   if "%attrCmd%"=="%%a" goto nochange
)
echo %attrCmd%>> %gitAttrFile%
echo Modification injected ...
exit
:nochange
echo .gitattributes was already modified ...
