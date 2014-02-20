@echo off
if "%OS%" == "Windows_NT" setlocal
echo "%1"
echo. 2>skip.launch
if "%1" == "" (
	cscript ".\scripts\compose.vbs" rolemapping.accdb
) else (
	cscript ".\scripts\compose.vbs" rolemapping.accdb "" %1
)
START /B rolemapping.accdb /decompile &
ping 1.1.1.1 -n 1 -w 3000 > nul
".\thirdparty\PSTools\pskill.exe" msaccess.exe
del skip.launch
