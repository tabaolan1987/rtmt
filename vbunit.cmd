@echo off
if "%OS%" == "Windows_NT" setlocal
cscript ".\scripts\run.vbs" rolemapping.accdb OnTest
ping 1.1.1.1 -n 1 -w 20000 > nul
".\thirdparty\PSTools\pskill.exe" msaccess.exe