@ECHO off
if "%OS%" == "Windows_NT" setlocal
cls
:start
ECHO RTMT additional commands
ECHO ===========================
ECHO 1. Delete all application data
ECHO 2. Enable hold shift key and click to open database
ECHO 3. Disable hold shift key and click to open database
ECHO 4. Exit
ECHO ===========================
set /p choice=Choose option: 
rem if not '%choice%'=='' set choice=%choice:~0;1% ( don`t use this command, because it takes only first digit in the case you type more digits. After that for example choice 23455666 is choice 2 and you get "bye"
if '%choice%'=='' ECHO "%choice%" is not valid please try again
if '%choice%'=='1' goto delete_all_application_data
if '%choice%'=='2' goto enable_shift_key
if '%choice%'=='3' goto disable_shift_key
if '%choice%'=='4' goto exit
ECHO.
goto start
:delete_all_application_data
echo. 2>skip.launch
ECHO Please wait while the system is processing your request
cscript ".\scripts\run.vbs" rolemapping.accdb WDeleteAllAppData
ECHO Completed
del skip.launch
goto end
:enable_shift_key
echo. 2>skip.launch
ECHO Please wait while the system is processing your request
cscript ".\scripts\run.vbs" rolemapping.accdb WEnableShift
ECHO Completed
del skip.launch
goto end
:disable_shift_key
echo. 2>skip.launch
ECHO Please wait while the system is processing your request
cscript ".\scripts\run.vbs" rolemapping.accdb WDisableShift
ECHO Completed
del skip.launch
goto end
:end
pause
exit

