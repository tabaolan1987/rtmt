@echo off
if "%OS%" == "Windows_NT" setlocal
echo "%INNO_HOME%"
if exist "%INNO_HOME%" goto okHome
echo The INNO_HOME environment variable is not defined correctly
echo This environment variable is needed to run this program
goto end
:okHome
cscript ".\scripts\make.vbs
iscc ".\target\inno-setup-script.iss"
pause
:end
