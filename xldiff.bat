@ECHO OFF
REM Paths to original and current spreadsheets
set base=%1
set mine=%2
REM Change forward slash to back slash and remove double quotes
set base=%base:/=\%
set base=%base:"=%
set mine=%mine:/=\%
REM Create paths.txt containing paths to spreadsheets for comparison
ECHO %base% > paths.txt
dir %mine% /B /S >> paths.txt
REM Run spreadsheetcompare 
"C:\Program Files\Microsoft Office\root\Client\AppVLP.exe" "C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE" paths.txt
REM paths.txt is deleted by automatically after it has been read