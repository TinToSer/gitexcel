@echo off
REM Developed by TinToSer(https://github.com/tintoser) and Muffexx (https://github.com/Muffexx)
REM SpreadsheetCompare.exe is only available with Office Professional Plus 2013, Office Professional Plus 2016, Office Professional Plus 2019, or Microsoft 365 Apps for enterprise.

ECHO -Starting gitExcel
SET appVlpPath="C:\Program Files\Microsoft Office\root\Client\AppVLP.exe"
SET spreadsheetComparePath="C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE"

REM Check for TortoiseGit mode
IF "%1"=="--tortoise" GOTO :tortoiseGitMode
IF "%1"=="-T" GOTO :tortoiseGitMode

REM Standard git diff mode
SET args=%*
FOR /F "tokens=2 delims=," %%a IN ("%args%") DO SET filex=%%~a
IF "%filex%" EQU "" GOTO :shellMode
FOR /F "tokens=1 delims=," %%a IN ("%args%") DO SET file1=%%~a
SET file2=%filex%
GOTO :continue

:shellMode
ECHO --Shell mode
SET file1="%~5"
SET file2="%~2"
GOTO :continue

REM Handle input from tortoiseGit
:tortoiseGitMode
ECHO --TortoiseGit mode
SET file1="%~2"
SET file2="%~3"

:continue
ECHO --Checking if files exist
IF NOT EXIST %file1% (
    ECHO ---File1 ^(mine^) does NOT EXIST: %file1%
    EXIT /B 1
)
IF NOT EXIST %file2% (
    ECHO ---File2 ^(base^) does NOT EXIST: %file2%
    EXIT /B 1
)
ECHO ---Found files


ECHO --Creating temp file
SET tempfile="%temp%\compareExcelTemp.txt"
DIR %file1:/=\% /B /S> %tempfile%
DIR %file2:/=\% /B /S>> %tempfile%


ECHO --Starting SpreadsheetCompare--
%appVlpPath% %spreadsheetComparePath% %tempfile%
timeout /t 5

REM If it throws an error then increase the sleep time
ECHO --Waiting for SpreadsheetCompare to open ...
IF "%filex%" EQU "" (
    sleep 10
) ELSE (
    timeout /t 10
)
