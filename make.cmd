@echo off
if "%OS%" == "Windows_NT" setlocal
echo "%INNO_HOME%"
if exist "%INNO_HOME%" goto okHome
echo The INNO_HOME environment variable is not defined correctly
echo This environment variable is needed to run this program
goto end
:okHome
echo "%1"
if "%1" == "" (
	cscript ".\scripts\init.vbs" DEVELOP
) else (
	cscript ".\scripts\init.vbs" %1
)

mkdir ".\target"
echo. 2>skip.launch
cscript ".\scripts\run.vbs" rolemapping.accdb WDeleteAllTables
cscript ".\scripts\run.vbs" rolemapping.accdb WDisableShift
copy ".\rolemapping.accdb" ".\target\rolemapping.accdb"
ping 1.1.1.1 -n 1 -w 3000 > nul
".\thirdparty\PSTools\pskill.exe" msaccess.exe

cscript ".\scripts\make.vbs"
iscc ".\target\inno-setup-script.iss"

if "%1" == "" (
	cscript ".\scripts\zip.vbs" DEVELOP
) else (
	cscript ".\scripts\zip.vbs" %1
)

del skip.launch
:end
