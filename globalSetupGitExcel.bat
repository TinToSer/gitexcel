@ECHO OFF
REM Developed by TinToSer(https://github.com/tintoser) and Muffexx (https://github.com/Muffexx)

ECHO -Starting Global Setup
SET userGitConfigPath="%homedrive%%homepath%\.gitconfig"
SET gitExcelPath="%cd%\gitExcel.cmd"


REM Search for AppVLP.exe and SPREADSHEETCOMPARE.EXE in predefined paths
ECHO --Searching AppVLP.exe and SPREADSHEETCOMPARE.EXE
SET appVlpPath=
FOR %%p IN (
    "C:\Program Files\Microsoft Office\root\Client",
    "C:\Program Files (x86)\Microsoft Office\root\Client"
) DO (
    CALL :CheckPath "%%~p\AppVLP.exe" appVlpPath
)
IF "%appVlpPath%"=="" (
    ECHO ---AppVLP.exe NOT FOUND in predefined paths
    GOTO :waitAndExitUnsuccessful
)
SET appVlpPath="%appVlpPath%"
ECHO ---Found: %appVlpPath%

SET spreadsheetComparePath=
FOR %%p IN (
    "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF",
    "C:\Program Files (x86)\Microsoft Office\Office15\DCF",
    "C:\Program Files (x86)\Microsoft Office\Office16\DCF",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\DCF"
) DO (
    CALL :CheckPath "%%~p\SPREADSHEETCOMPARE.EXE" spreadsheetComparePath
)
IF "%spreadsheetComparePath%"=="" (
    ECHO ---SPREADSHEETCOMPARE.EXE NOT FOUND in predefined paths
    GOTO :waitAndExitUnsuccessful
)
SET spreadsheetComparePath="%spreadsheetComparePath%"
ECHO ---Found: %spreadsheetComparePath%


REM Writing gitExcel section to .gitconfig
ECHO --Updating user .gitconfig
SET "gitConfigDiffText=[diff "gitExcel"]"
SET "gitConfigCommandText=    command = %gitExcelPath:\=/%"

FINDSTR /R /C:"^%gitConfigDiffText%" %userGitConfigPath% >NUL
IF ERRORLEVEL 1 (
    ECHO %gitConfigDiffText%>> %userGitConfigPath%
    ECHO ---Added: %gitConfigDiffText%
    ECHO %gitConfigCommandText%>> %userGitConfigPath%
    ECHO ---Added: %gitConfigCommandText%
) ELSE (
    ECHO ---.gitconfig already contains %gitConfigDiffText% section
)


REM Replace paths in gitExcel.cmd
ECHO --Updating %gitExcelPath%
CALL :UpdateGitExcelFile %gitExcelPath%

GOTO :waitAndExitSuccessful


REM Function returns the name if the file exists
:CheckPath
IF EXIST "%~1" (
    SET "%2=%~1")
GOTO :EOF


REM Search for "SET appVlpPath/spreadsheetComparePath" and replace the lines with the found paths
:UpdateGitExcelFile
SET filePath=%1
SET tempFilePath="%cd%\temp.cmd"

SET "_findVlpPath=SET appVlpPath="
SET "_findSscPath=SET spreadsheetComparePath="
SET "_setVlpPath=SET appVlpPath=%appVlpPath%"
SET "_setSscPath=SET spreadsheetComparePath=%spreadsheetComparePath%"

(FOR /f "delims=" %%l IN ('findstr /n "^" %filePath%') DO (
    SET "Line=%%l"
    for /F "delims=:" %%n in ("%%l") do set "LNum=%%n"
    SETLOCAL ENABLEDELAYEDEXPANSION
    SET "Line=!Line:*:=!"
    IF !Line! equ !_findVlpPath! SET "Line=!_setVlpPath!"
    IF !Line! equ !_findSscPath! SET "Line=!_setSscPath!"
    ECHO(!Line!
    ENDLOCAL
))> %tempFilePath%
MOVE /Y %tempFilePath% %filePath%
GOTO :EOF


:waitAndExitSuccessful
ECHO -Finished Global Setup
TIMEOUT /T 10
EXIT /B 0


:waitAndExitUnsuccessful
ECHO -Global Setup FAILED
TIMEOUT /T 10
EXIT /B 1
