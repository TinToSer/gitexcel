@echo off
Rem Developed by TinToSer(https://github.com/tintoser)
Rem It works for office 2019/2021/365
set file1=%1
if "%5" neq "" set file1=%5
set file2=%2

set file1=%file1:/=\%
set file2=%file2:/=\%
set tmpfile="%temp%\tocompareexcel.txt"
dir %file2% /B /S > %tmpfile%
dir %file1% /B /S >> %tmpfile%
"C:\Program Files\Microsoft Office\root\Client\AppVLP.exe" "C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE" %tmpfile%
Rem If it throws an error of No Excel file found then increase the sleep time to 10
sleep 10
