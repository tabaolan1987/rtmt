@echo off
if "%OS%" == "Windows_NT" setlocal
echo. 2>skip.launch
cscript ".\scripts\run.vbs" rolemapping.accdb WDeleteAllTables
cscript ".\scripts\decompose.vbs" rolemapping.accdb
del skip.launch