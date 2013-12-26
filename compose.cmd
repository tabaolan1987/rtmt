@echo off
if "%OS%" == "Windows_NT" setlocal

cscript ".\scripts\compose.vbs" rolemapping.accdb
START /B rolemapping.accdb /decompile &
ping 1.1.1.1 -n 1 -w 5000 > nul
".\thirdparty\PSTools\pskill.exe" msaccess.exe