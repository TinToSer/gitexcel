@ECHO OFF
REM Developed by TinToSer(https://github.com/tintoser) and Muffexx (https://github.com/Muffexx)

ECHO -Starting Repo Setup
SET "attrCmd=*.xls* diff=gitExcel"

REM Read repository path
SET /P gitPath=--ENTER REPOSITORY PATH:
SET gitAttrFile=%gitPath%\.gitattributes


REM Check if the .gitattributes file exists
ECHO --Checking if .gitattributes file exists
IF NOT EXIST "%gitAttrFile%" (
    ECHO ---.gitattributes file NOT found
    ECHO --Creating and configuring.gitattributes file
    ECHO %attrCmd% > "%gitAttrFile%"
    GOTO waitAndExitSuccessful
)
ECHO ---.gitattributes file found


REM Check if the line already exists in the .gitattributes file
ECHO --Checking .gitattributes file configuration
SETLOCAL EnableDelayedExpansion
SET "attrCmdTrimmed=%attrCmd: =%"
FOR /F "delims=" %%a IN ('TYPE "%gitAttrFile%"') DO (
    SET "line=%%a"
    SET "lineTrimmed=!line: =!"
    ECHO !lineTrimmed! | FINDSTR /I /C:"%attrCmdTrimmed%" >NUL
    IF NOT ERRORLEVEL 1 (
        ECHO ---gitattributes file already configured
        ENDLOCAL
        GOTO waitAndExitSuccessful
    )
)
ENDLOCAL
ECHO ---gitattributes file NOT YET configured


REM Append the line to the .gitattributes file
ECHO --Updating .gitattributes file
ECHO %attrCmd% >> "%gitAttrFile%"
GOTO waitAndExitSuccessful


:waitAndExitSuccessful
ECHO -Finished Repo Setup
TIMEOUT /T 10
EXIT /B 0
