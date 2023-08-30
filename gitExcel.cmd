@echo off
Rem Developed by TinToSer(https://github.com/tintoser)
Rem It works for office 2019/2021/365 further you can check

set args=%*

REM It is the safest method to incorporate filenames even with spaces when calling from thirdparty programs with comma separated arguments
FOR /F "tokens=2 delims=," %%a IN ("%args%") DO set filex=%%~a
IF "%filex%" EQU "" GOTO shelltype
FOR /F "tokens=1 delims=," %%a IN ("%args%") DO set file1=%%~a
set file2=%filex%
GOTO continue

:shelltype
set file1=%~5
set file2=%~2

:continue
set file1=%file1:/=\%
set file2=%file2:/=\%
set tmpfile="%temp%\tocompareexcel.txt"
dir "%file2%" /B /S> %tmpfile%
dir "%file1%" /B /S>> %tmpfile%

"C:\Program Files\Microsoft Office\root\Client\AppVLP.exe" "C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE" %tmpfile%

REM If it throws an error of No Excel file found then increase the sleep time as your excel file may be taking time to open
echo Waiting for Excel to open ...
IF "%filex%" EQU "" (sleep 10) else (timeout /t 10)
