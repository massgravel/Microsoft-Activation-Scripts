@setlocal DisableDelayedExpansion
@echo off



::============================================================================
::
::   This script is a part of 'Microsoft_Activation_Scripts' (MAS) project.
::
::   Homepage: mass grave.dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================




::========================================================================================================================================

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="r1" set r1=1
if /i "%%#"=="r2" set r2=1
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

::  Set Path variable, it helps if it is misconfigured in the system

set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "PATH=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%PATH%"
)

::  Check LF line ending

pushd "%~dp0"
>nul findstr /rxc:".*" "%~nx0"
if not %errorlevel%==0 (
echo:
echo Error: Script either has LF line ending issue, or it failed to read itself.
echo:
ping 127.0.0.1 -n 6 > nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title Extract $OEM$ Folder

set winbuild=1
set "nul=>nul 2>&1"
set psc=powershell.exe
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 2>nul | find /i "0x0" 1>nul && (set _NCS=0)

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set   "Green="42;97m""
set "Magenta="45;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set   "Green="DarkGreen" "white""
set "Magenta="Darkmagenta" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :ex_color %Red% "==== ERROR ====" &echo:"

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected.
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
goto done2
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto done2
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name

set desktop=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "desktop=%%b"
if not defined desktop for /f "delims=" %%a in ('%psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"
set "_pdesk=%desktop:'=''%"

set _PSarg="""%~f0""" -el %_args%
set "_ttemp=%temp%"

set "_dir=%desktop%\$OEM$\$$\Setup\Scripts"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" 1>nul && (
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

>nul fltmc || (
if not defined _elev %nul% %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script require admin privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto done2
)

::========================================================================================================================================

if not exist "!desktop!\" (
%eline%
echo Desktop location was not detected, aborting...
goto done2
)

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

pushd "!_work!"

set _nofile=
for %%# in (
%HWID_Activation.cmd%
%KMS38_Activation.cmd%
%Online_KMS_Activation.cmd%
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
echo:
echo:                     Extract $OEM$ folder on the desktop           
echo:           ________________________________________________________
echo:                                                                  
echo:              [1] HWID
echo:              [2] KMS38
echo:              [3] Online KMS
echo:          
echo:              [4] HWID  ^(Windows^) ^+ Online KMS ^(Office^)
echo:              [5] KMS38 ^(Windows^) ^+ Online KMS ^(Office^)
echo:          
echo:              [0] Exit                                            
echo:           ________________________________________________________
echo:  
call :ex_color2 %_White% "             " %_Green% "Enter a menu option in the Keyboard [1,2,3,4,5,0]"
choice /C:123450 /N
set _erl=%errorlevel%

if %_erl%==6 exit /b
if %_erl%==5 goto:kms38_kms
if %_erl%==4 goto:hwid_kms
if %_erl%==3 goto:kms
if %_erl%==2 goto:kms38
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

start /b /wait cmd /c "%~dp0HWID_Activation.cmd" /HWID

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_setup:

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

start /b /wait cmd /c "%~dp0KMS38_Activation.cmd" /KMS38

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

start /b /wait cmd /c "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-WindowsOffice

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms_setup:

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

start /b /wait cmd /c "%~dp0HWID_Activation.cmd" /HWID

start /b /wait cmd /c "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-Office

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_kms_setup:

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

start /b /wait cmd /c "%~dp0KMS38_Activation.cmd" /KMS38

start /b /wait cmd /c "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-Office

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_kms_setup:

::========================================================================================================================================

:errorfound

%eline%
echo $OEM$ Folder was not created successfully...
goto :done2

:done

set -=
echo ______________________________________________________________
echo:
call :ex_color %Magenta% "%oem%"
call :ex_color %Green% "$OEM$ folder is successfully created on the Desktop."
echo "%oem%" | find /i "38" %nul% && (
echo:
echo To KMS38 activate Server Cor/Acor editions ^(No GUI Versions^),
echo Check this page https://mass%-%grave.dev/oem-folder
)
echo ______________________________________________________________

:done2

echo:
call :ex_color %_Yellow% "Press any key to exit..."
pause >nul
exit /b

::========================================================================================================================================

::  Extract the text from batch script without character and file encoding issue

:export

%nul% %psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('!_pdesk!\$OEM$\$$\Setup\Scripts\SetupComplete.cmd',$f[1].Trim(),[System.Text.Encoding]::ASCII);"
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