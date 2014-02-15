@echo off
if "%OS%" == "Windows_NT" setlocal
echo "%INNO_HOME%"
if exist "%INNO_HOME%" goto okHome
echo The INNO_HOME environment variable is not defined correctly
echo This environment variable is needed to run this program
goto end
:okHome

mkdir ".\target"
cscript ".\scripts\run.vbs" rolemapping.accdb WDeleteAllTables
copy ".\rolemapping.accdb" ".\target\rolemapping.accdb"
ping 1.1.1.1 -n 1 -w 5000 > nul
".\thirdparty\PSTools\pskill.exe" msaccess.exe

cscript ".\scripts\make.vbs
iscc ".\target\inno-setup-script.iss"
:end
