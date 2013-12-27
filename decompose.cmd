@echo off
if "%OS%" == "Windows_NT" setlocal
cscript ".\scripts\decompose.vbs" rolemapping.accdb
pause