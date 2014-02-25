@echo off
if "%OS%" == "Windows_NT" setlocal
echo. 2>skip.launch
cscript ".\scripts\run.vbs" rolemapping.accdb WGetAllTables
cscript ".\scripts\run.vbs" rolemapping.accdb OnTest
ping 1.1.1.1 -n 1 -w 10000 > nul
".\thirdparty\PSTools\pskill.exe" msaccess.exe
del skip.launch