@echo off
if "%OS%" == "Windows_NT" setlocal
echo "%1"
if "%1" == "" (
	cscript ".\scripts\compose.vbs" rolemapping.accdb
) else (
	cscript ".\scripts\compose.vbs" rolemapping.accdb "" %1
)
START /B rolemapping.accdb /decompile &
ping 1.1.1.1 -n 1 -w 5000 > nul
".\thirdparty\PSTools\pskill.exe" msaccess.exe

if "%1" == "prod" (
	mkdir ".\target"
	copy ".\rolemapping.accdb" ".\target\rolemapping.accdb"
	cscript ".\scripts\run.vbs" rolemapping.accdb MakeAccde	
	ping 1.1.1.1 -n 1 -w 5000 > nul
	".\thirdparty\PSTools\pskill.exe" msaccess.exe
)

