@echo off
if "%OS%" == "Windows_NT" setlocal
echo "%1"
if "%1" == "" (
	cscript ".\scripts\init.vbs" DEVELOP
) else (
	cscript ".\scripts\init.vbs" %1

)
echo. 2>skip.launch
cscript ".\scripts\run.vbs" rolemapping.accdb WEnableShift
del skip.launch