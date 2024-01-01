@set masver=2.5
@setlocal DisableDelayedExpansion
@echo off



::============================================================================
::
::   This script is a part of 'Microsoft_Activation_Scripts' (MAS) project.
::
::   Homepage: mass grave[.]dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================



::========================================================================================================================================

::  Set Path variable, it helps if it is misconfigured in the system

set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "PATH=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%PATH%"
)

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="r1" set r1=1
if /i "%%#"=="r2" set r2=1
if /i "%%#"=="-qedit" (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "1" /f %nul1%
rem check the code below admin elevation to understand why it's here
)
)

if exist %SystemRoot%\Sysnative\cmd.exe if not defined r1 (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* r1"
exit /b
)

:: Re-launch the script with ARM32 process if it was initiated by x64 process on ARM64 Windows

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 if not defined r2 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* r2"
exit /b
)

::========================================================================================================================================

set "blank="
set "mas=ht%blank%tps%blank%://mass%blank%grave.dev/"

::  Check if Null service is working, it's important for the batch script

sc query Null | find /i "RUNNING"
if %errorlevel% NEQ 0 (
echo:
echo Null service is not running, script may crash...
echo:
echo:
echo Help - %mas%troubleshoot.html
echo:
echo:
ping 127.0.0.1 -n 10
)
cls

::  Check LF line ending

pushd "%~dp0"
>nul findstr /v "$" "%~nx0" && (
echo:
echo Error: Script either has LF line ending issue or an empty line at the end of the script is missing.
echo:
ping 127.0.0.1 -n 6 >nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title Extract $OEM$ Folder %masver%

set _args=
set _elev=

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="-el"                    set _elev=1
)
)

set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"

set psc=powershell.exe
set winbuild=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 %nul2% | find /i "0x0" %nul1% && (set _NCS=0)

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :ex_color %Red% "==== ERROR ====" &echo:"

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected [%winbuild%].
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
goto done2
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto done2
)

::========================================================================================================================================

::  Fix special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%
set "_ttemp=%userprofile%\AppData\Local\Temp"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
if /i not "!_work!"=="!_ttemp!" (
%eline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto done2
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo To do so, right click on this script and select 'Run as administrator'.
goto done2
)

::========================================================================================================================================

::  This code disables QuickEdit for this cmd.exe session only without making permanent changes to the registry
::  It is added because clicking on the script window pauses the operation and leads to the confusion that script stopped due to an error

for %%# in (%_args%) do (if /i "%%#"=="-qedit" set quedit=1)

reg query HKCU\Console /v QuickEdit %nul2% | find /i "0x0" %nul1% || if not defined quedit (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "0" /f %nul1%
start cmd.exe /c ""!_batf!" %_args% -qedit"
rem quickedit reset code is added at the starting of the script instead of here because it takes time to reflect in some cases
exit /b
)

::========================================================================================================================================

::  Check for updates

set -=
set old=

for /f "delims=[] tokens=2" %%# in ('ping -4 -n 1 updatecheck.mass%-%grave.dev') do (
if not [%%#]==[] (echo "%%#" | find "127.69" %nul1% && (echo "%%#" | find "127.69.%masver%" %nul1% || set old=1))
)

if defined old (
echo ________________________________________________
%eline%
echo You are running outdated version MAS %masver%
echo ________________________________________________
echo:
echo [1] Get Latest MAS
echo [0] Continue Anyway
echo:
call :ex_color %_Green% "Enter a menu option in the Keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
)
cls

::========================================================================================================================================

setlocal DisableDelayedExpansion

::  Check desktop location

set desktop=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "desktop=%%b"
if not defined desktop for /f "delims=" %%a in ('%psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"

set "_pdesk=%desktop:'=''%"
set "_dir=%desktop%\$OEM$\$$\Setup\Scripts"

if exist "!desktop!\" (
%eline%
echo Desktop location was not detected, aborting...
goto done2
)

setlocal EnableDelayedExpansion

::========================================================================================================================================

mode con cols=78 lines=30

if exist "!desktop!\$OEM$\" (
echo _____________________________________________________
%eline%
echo $OEM$ folder already exists on the Desktop.
echo _____________________________________________________
goto done2
)

set HWID_Activation.cmd=Activators\HWID_Activation.cmd
set KMS38_Activation.cmd=Activators\KMS38_Activation.cmd
set Online_KMS_Activation.cmd=Activators\Online_KMS_Activation.cmd
set Ohook_Activation_AIO.cmd=Activators\Ohook_Activation_AIO.cmd
pushd "!_work!"

set _nofile=
for %%# in (
%HWID_Activation.cmd%
%KMS38_Activation.cmd%
%Online_KMS_Activation.cmd%
%Ohook_Activation_AIO.cmd%
) do (
if not exist "%%#" set _nofile=1
)

popd

if defined _nofile (
echo _____________________________________________________
%eline%
echo Some files are missing in the 'Activators' folder.
echo _____________________________________________________
goto done2
)

::========================================================================================================================================

:Menu

cls
mode con cols=78 lines=30
echo:
echo:
echo:
echo:
echo:                     Extract $OEM$ folder on the desktop           
echo:           ________________________________________________________
echo:
echo:              [1] HWID
echo:              [2] Ohook
echo:              [3] KMS38
echo:              [4] Online KMS
echo:
echo:              [5] HWID       ^(Windows^) ^+ Ohook      ^(Office^)
echo:              [6] HWID       ^(Windows^) ^+ Online KMS ^(Office^)
echo:              [7] KMS38      ^(Windows^) ^+ Ohook      ^(Office^)
echo:              [8] KMS38      ^(Windows^) ^+ Online KMS ^(Office^)
echo:              [9] Online KMS ^(Windows^) ^+ Ohook      ^(Office^)
echo:
call :ex_color2 %_White% "              [R] " %_Green% "ReadMe"
echo:              [0] Exit
echo:           ________________________________________________________
echo:  
call :ex_color2 %_White% "             " %_Green% "Enter a menu option in the Keyboard :"
choice /C:123456789R0 /N
set _erl=%errorlevel%

if %_erl%==11 exit /b
if %_erl%==10 start %mas%oem-folder.html &goto :Menu
if %_erl%==9 goto:kms_ohook
if %_erl%==8 goto:kms38_kms
if %_erl%==7 goto:kms38_ohook
if %_erl%==6 goto:hwid_kms
if %_erl%==5 goto:hwid_ohook
if %_erl%==4 goto:kms
if %_erl%==3 goto:kms38
if %_erl%==2 goto:ohook
if %_erl%==1 goto:hwid
goto :Menu

::========================================================================================================================================

:hwid

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
popd
call :export hwid_setup

set _error=
if not exist "!_dir!\HWID_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=HWID
goto done

:hwid_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0HWID_Activation.cmd" /HWID

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_setup:

::========================================================================================================================================

:ohook

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b %Ohook_Activation_AIO.cmd% "!_dir!\Ohook_Activation_AIO.cmd" %nul%
popd
call :export ohook_setup

set _error=
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=Ohook
goto done

:ohook_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0Ohook_Activation_AIO.cmd" /Ohook

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:ohook_setup:

::========================================================================================================================================

:kms38

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%KMS38_Activation.cmd%" "!_dir!\KMS38_Activation.cmd" %nul%
popd
call :export kms38_setup

set _error=
if not exist "!_dir!\KMS38_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=KMS38
goto done

:kms38_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0KMS38_Activation.cmd" /KMS38

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_setup:

::========================================================================================================================================

:kms

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
popd
call :export kms_setup

set _error=
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=Online KMS
goto done

:kms_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-WindowsOffice

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms_setup:

::========================================================================================================================================

:hwid_ohook

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
copy /y /b "%Ohook_Activation_AIO.cmd%" "!_dir!\Ohook_Activation_AIO.cmd" %nul%
popd
call :export hwid_ohook_setup

set _error=
if not exist "!_dir!\HWID_Activation.cmd" set _error=1
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=HWID [Windows] + Ohook [Office]
goto done

:hwid_ohook_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0HWID_Activation.cmd" /HWID
endlocal

setlocal
call "%~dp0Ohook_Activation_AIO.cmd" /Ohook
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_ohook_setup:

::========================================================================================================================================

:hwid_kms

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
popd
call :export hwid_kms_setup

set _error=
if not exist "!_dir!\HWID_Activation.cmd" set _error=1
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=HWID [Windows] + Online KMS [Office]
goto done

:hwid_kms_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0HWID_Activation.cmd" /HWID
endlocal

setlocal
call "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-Office
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_kms_setup:

::========================================================================================================================================

:kms38_ohook

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%KMS38_Activation.cmd%" "!_dir!\KMS38_Activation.cmd" %nul%
copy /y /b "%Ohook_Activation_AIO.cmd%" "!_dir!\Ohook_Activation_AIO.cmd" %nul%
popd
call :export kms38_ohook_setup

set _error=
if not exist "!_dir!\KMS38_Activation.cmd" set _error=1
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=KMS38 [Windows] + Ohook [Office]
goto done

:kms38_ohook_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0KMS38_Activation.cmd" /KMS38
endlocal

setlocal
call "%~dp0Ohook_Activation_AIO.cmd" /Ohook
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_ohook_setup:

::========================================================================================================================================

:kms38_kms

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%KMS38_Activation.cmd%" "!_dir!\KMS38_Activation.cmd" %nul%
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
popd
call :export kms38_kms_setup

set _error=
if not exist "!_dir!\KMS38_Activation.cmd" set _error=1
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=KMS38 [Windows] + Online KMS [Office]
goto done

:kms38_kms_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0KMS38_Activation.cmd" /KMS38
endlocal

setlocal
call "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-Office
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_kms_setup:

::========================================================================================================================================

:kms_ohook

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
copy /y /b "%Ohook_Activation_AIO.cmd%" "!_dir!\Ohook_Activation_AIO.cmd" %nul%
popd
call :export kms_ohook_setup

set _error=
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=Online KMS [Windows] + Ohook [Office]
goto done

:kms_ohook_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-Windows
endlocal

setlocal
call "%~dp0Ohook_Activation_AIO.cmd" /Ohook
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms_ohook_setup:

::========================================================================================================================================

:errorfound

%eline%
echo $OEM$ Folder was not created successfully...
goto :done2

:done

echo ______________________________________________________________
echo:
call :ex_color %Blue% "%oem%"
call :ex_color %Green% "$OEM$ folder is successfully created on the Desktop."
echo "%oem%" | find /i "38" %nul% && (
echo:
echo To KMS38 activate Server Cor/Acor editions ^(No GUI Versions^),
echo Check this page %mas%oem-folder
)
echo ______________________________________________________________

:done2

echo:
call :ex_color %_Yellow% "Press any key to exit..."
pause %nul1%
exit /b

::========================================================================================================================================

::  Extract the text from batch script without character and file encoding issue

:export

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('!_pdesk!\$OEM$\$$\Setup\Scripts\SetupComplete.cmd',$f[1].Trim(),[System.Text.Encoding]::ASCII);"
exit /b

::========================================================================================================================================

:ex_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
if not exist %psc% (echo %~3) else (%psc% write-host -back '%1' -fore '%2' '%3')
)
exit /b

:ex_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
if not exist %psc% (echo %~3%~6) else (%psc% write-host -back '%1' -fore '%2' '%3' -NoNewline; write-host -back '%4' -fore '%5' '%6')
)
exit /b

::========================================================================================================================================
:: Leave empty line below
