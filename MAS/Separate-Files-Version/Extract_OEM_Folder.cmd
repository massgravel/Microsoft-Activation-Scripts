@set masver=3.4
@echo off



::============================================================================
::
::   Homepage: mass grave[.]dev
::      Email: mas.help@outlook.com
::
::============================================================================



::========================================================================================================================================

::  Set environment variables, it helps if they are misconfigured in the system

setlocal EnableExtensions
setlocal DisableDelayedExpansion

set "PathExt=.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC"

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)

set "ComSpec=%SysPath%\cmd.exe"
set "PSModulePath=%ProgramFiles%\WindowsPowerShell\Modules;%SysPath%\WindowsPowerShell\v1.0\Modules"

set re1=
set re2=
set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="re1" set re1=1
if /i "%%#"=="re2" set re2=1
)

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

if exist %SystemRoot%\Sysnative\cmd.exe if not defined re1 (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* re1"
exit /b
)

:: Re-launch the script with ARM32 process if it was initiated by x64 process on ARM64 Windows

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 if not defined re2 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* re2"
exit /b
)

::========================================================================================================================================

set "blank="
set "mas=ht%blank%tps%blank%://mass%blank%grave.dev/"
set "github=ht%blank%tps%blank%://github.com/massgra%blank%vel/Micro%blank%soft-Acti%blank%vation-Scripts"
set "selfgit=ht%blank%tps%blank%://git.acti%blank%vated.win/massg%blank%rave/Micr%blank%osoft-Act%blank%ivation-Scripts"

::  Check if Null service is working, it's important for the batch script

sc query Null | find /i "RUNNING"
if %errorlevel% NEQ 0 (
echo:
echo Null service is not running, script may crash...
echo:
echo:
echo Check this webpage for help - %mas%fix_service
echo:
echo:
ping 127.0.0.1 -n 20
)
cls

::  Check LF line ending

pushd "%~dp0"
>nul findstr /v "$" "%~nx0" && (
echo:
echo Error - Script either has LF line ending issue or an empty line at the end of the script is missing.
echo:
echo:
echo Check this webpage for help - %mas%troubleshoot
echo:
echo:
ping 127.0.0.1 -n 20 >nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title  Extract $OEM$ Folder %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args set _args=%_args:re1=%
if defined _args set _args=%_args:re2=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="-el"                    set _elev=1
)
)

set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"

call :dk_setvar

::========================================================================================================================================

if %winbuild% EQU 1 (
%eline%
echo Failed to detect Windows build number.
echo:
setlocal EnableDelayedExpansion
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
goto done2
)

if %winbuild% LSS 6001 (
%nceline%
echo Unsupported OS version detected [%winbuild%].
echo MAS only supports Windows Vista/7/8/8.1/10/11 and their Server equivalents.
if %winbuild% EQU 6000 (
echo:
echo Windows Vista RTM is not supported because Powershell cannot be installed.
echo Upgrade to Windows Vista SP1 or SP2.
)
goto done2
)

if %winbuild% LSS 7600 if not exist "%SysPath%\WindowsPowerShell\v1.0\Modules" (
%nceline%
if not exist %ps% (
echo PowerShell is not installed in your system.
)
echo Install PowerShell 2.0 using the following URL.
echo:
echo https://www.catalog.update.microsoft.com/Search.aspx?q=KB968930
if %_unattended%==0 start https://www.catalog.update.microsoft.com/Search.aspx?q=KB968930
goto done2
)

::========================================================================================================================================

::  Fix special character limitations in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%
set _PSarg=%_PSarg:'=''%

set "_ttemp=%userprofile%\AppData\Local\Temp"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
if /i not "!_work!"=="!_ttemp!" (
%eline%
echo The script was launched from the temp folder.
echo You are most likely running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto done2
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo Right click on this script and select 'Run as administrator'.
goto done2
)

::========================================================================================================================================

::  Check PowerShell

::pstst $ExecutionContext.SessionState.LanguageMode :pstst

for /f "delims=" %%a in ('%psc% "if ($PSVersionTable.PSEdition -ne 'Core') {$f=[io.file]::ReadAllText('!_batp!') -split ':pstst';iex ($f[1])}" %nul6%') do (set tstresult=%%a)

if /i not "%tstresult%"=="FullLanguage" (
%eline%
for /f "delims=" %%a in ('%psc% "$ExecutionContext.SessionState.LanguageMode" %nul6%') do (set tstresult2=%%a)
echo Test 1 - %tstresult%
echo Test 2 - !tstresult2!
echo:

REM check LanguageMode

echo: !tstresult2! | findstr /i "ConstrainedLanguage RestrictedLanguage NoLanguage" %nul1% && (
echo FullLanguage mode not found in PowerShell. Aborting...
echo If you have applied restrictions on Powershell then undo those changes.
echo:
set fixes=%fixes% %mas%fix_powershell
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%fix_powershell"
goto done2
)

REM check Powershell core version

cmd /c "%psc% "$PSVersionTable.PSEdition"" | find /i "Core" %nul1% && (
echo Windows Powershell is needed for MAS but it seems to be replaced with Powershell core. Aborting...
goto done2
)

REM check for Mal-ware that may cause issues with Powershell

for /r "%ProgramFiles%\" %%f in (secureboot.exe) do if exist "%%f" (
echo "%%f"
echo Mal%blank%ware found, PowerShell is not working properly.
echo:
set fixes=%fixes% %mas%remove_mal%w%ware
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%remove_mal%w%ware"
goto done2
)

REM check antivirus and other errors

echo PowerShell is not working properly. Aborting...

if /i "!tstresult2!"=="FullLanguage" (
echo:
echo Your antivirus software might be blocking the script, or PowerShell on your system might be corrupted.
cmd /c "%psc% ""$av = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct; $n = @(); foreach ($i in $av) { $n += $i.displayName }; if ($n) { Write-Host ('Installed Antivirus - ' + ($n -join ', '))}"""
)

echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
goto done2
)

::========================================================================================================================================

::  Disable QuickEdit and launch from conhost.exe to avoid Terminal app

if %winbuild% GEQ 17763 (
set terminal=1
) else (
set terminal=
)

::  Check if script is running in Terminal app

if defined terminal (
set lines=0
for /f "skip=2 tokens=2 delims=: " %%A in ('mode con') do if "!lines!"=="0" set lines=%%A
if !lines! GEQ 100 set terminal=
)

if %_unattended%==1 goto :skipQE
for %%# in (%_args%) do (if /i "%%#"=="-qedit" goto :skipQE)

::  Relaunch to disable QuickEdit in the current session and use conhost.exe instead of the Terminal app
::  This code disables QuickEdit for the current cmd.exe session without making permanent registry changes
::  It is included because clicking on the script window can pause execution, causing confusion that the script has stopped due to an error

set resetQE=1
reg query HKCU\Console /v QuickEdit %nul2% | find /i "0x0" %nul1% && set resetQE=0
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d 0 /f %nul1%

if defined terminal (
start conhost.exe "!_batf!" %_args% -qedit
start reg add HKCU\Console /v QuickEdit /t REG_DWORD /d %resetQE% /f %nul1%
exit /b
) else if %resetQE% EQU 1 (
start cmd.exe /c ""!_batf!" %_args% -qedit"
start reg add HKCU\Console /v QuickEdit /t REG_DWORD /d %resetQE% /f %nul1%
exit /b
)

:skipQE

::========================================================================================================================================

::  Check for updates

set -=
set old=
set pingp=
set upver=%masver:.=%

for %%A in (
activ%-%ated.win
mass%-%grave.dev
) do if not defined pingp (
for /f "delims=[] tokens=2" %%B in ('ping -n 1 %%A') do (
if not "%%B"=="" (set old=1& set pingp=1)
for /f "delims=[] tokens=2" %%C in ('ping -n 1 updatecheck%upver%.%%A') do (
if not "%%C"=="" set old=
)
)
)

if defined old (
echo ________________________________________________
%eline%
echo Your version of MAS [%masver%] is outdated.
echo ________________________________________________
echo:
if not %_unattended%==1 (
echo [1] Get Latest MAS
echo [0] Continue Anyway
echo:
call :dk_color %_Green% "Choose a menu option using your keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start %selfgit% & start %github% & start %mas% & exit /b)
)
)

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
echo Unable to detect Desktop location, aborting...
goto done2
)

setlocal EnableDelayedExpansion

::========================================================================================================================================

if not defined terminal mode 78, 30

if exist "!desktop!\$OEM$\" (
echo _____________________________________________________
%eline%
echo The $OEM$ folder already exists on your Desktop.
echo _____________________________________________________
goto done2
)

set HWID_Activation.cmd=Activators\HWID_Activation.cmd
set KMS38_Activation.cmd=Activators\KMS38_Activation.cmd
set Online_KMS_Activation.cmd=Activators\Online_KMS_Activation.cmd
set Ohook_Activation_AIO.cmd=Activators\Ohook_Activation_AIO.cmd
set TSforge_Activation.cmd=Activators\TSforge_Activation.cmd
pushd "!_work!"

set _nofile=
for %%# in (
%HWID_Activation.cmd%
%KMS38_Activation.cmd%
%Online_KMS_Activation.cmd%
%Ohook_Activation_AIO.cmd%
%TSforge_Activation.cmd%
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
if not defined terminal mode 78, 30
echo:
echo:
echo:
echo:
echo:                     Extract $OEM$ folder on the desktop           
echo:         ____________________________________________________________
echo:
echo:            [1] HWID             [Windows]
echo:            [2] Ohook            [Office]
echo:            [3] TSforge          [Windows / ESU / Office]
echo:            [4] KMS38            [Windows]
echo:            [5] Online KMS       [Windows / Office]
echo:
echo:            [6] HWID    [Windows] ^+ Ohook [Office]
echo:            [7] HWID    [Windows] ^+ Ohook [Office] ^+ TSforge [ESU]
echo:            [8] TSforge [Windows] ^+ Online KMS [Office]
echo:
call :dk_color2 %_White% "            [R] " %_Green% "ReadMe"
echo:            [0] Exit
echo:         ____________________________________________________________
echo:  
call :dk_color2 %_White% "             " %_Green% "Choose a menu option using your keyboard :"
choice /C:12345678R0 /N
set _erl=%errorlevel%

if %_erl%==10 exit /b
if %_erl%==9 start %mas%oem-folder &goto :Menu
if %_erl%==8 goto:tsforge_kms
if %_erl%==7 goto:hwid_ohook_tsforge
if %_erl%==6 goto:hwid_ohook
if %_erl%==5 goto:kms
if %_erl%==4 goto:kms38
if %_erl%==3 goto:tsforge
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

:tsforge

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%TSforge_Activation.cmd%" "!_dir!\TSforge_Activation.cmd" %nul%
popd
call :export tsforge_setup

set _error=
if not exist "!_dir!\TSforge_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=TSforge
goto done

:tsforge_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0TSforge_Activation.cmd" /Z-WindowsESUOffice

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:tsforge_setup:

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

call "%~dp0Online_KMS_Activation.cmd" /K-WindowsOffice

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

:hwid_ohook_tsforge

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
copy /y /b "%Ohook_Activation_AIO.cmd%" "!_dir!\Ohook_Activation_AIO.cmd" %nul%
copy /y /b "%TSforge_Activation.cmd%" "!_dir!\TSforge_Activation.cmd" %nul%
popd
call :export hwid_ohook_tsforge_setup

set _error=
if not exist "!_dir!\HWID_Activation.cmd" set _error=1
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\TSforge_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=HWID [Windows] + Ohook [Office] + TSforge [ESU]
goto done

:hwid_ohook_tsforge_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0HWID_Activation.cmd" /HWID
endlocal

setlocal
call "%~dp0Ohook_Activation_AIO.cmd" /Ohook
endlocal

setlocal
call "%~dp0TSforge_Activation.cmd" /Z-ESU
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_ohook_tsforge_setup:

::========================================================================================================================================

:tsforge_kms

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%TSforge_Activation.cmd%" "!_dir!\TSforge_Activation.cmd" %nul%
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
popd
call :export tsforge_kms_setup

set _error=
if not exist "!_dir!\TSforge_Activation.cmd" set _error=1
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=TSforge [Windows] + Online KMS [Office]
goto done

:tsforge_kms_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0TSforge_Activation.cmd" /Z-Windows
endlocal

setlocal
call "%~dp0Online_KMS_Activation.cmd" /K-Office
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:tsforge_kms_setup:

::========================================================================================================================================

:errorfound

%eline%
echo The script failed to create the $OEM$ folder.
goto :done2

:done

echo ______________________________________________________________
echo:
call :dk_color %Blue% "%oem%"
call :dk_color %Green% "$OEM$ folder was successfully created on your Desktop."
echo "%oem%" | find /i "38" %nul% && (
echo:
echo To KMS38 activate Server Cor/Acor editions [No GUI Versions],
echo Check this page %mas%oem-folder
)
echo ______________________________________________________________

:done2

echo:
if defined fixes (
call :dk_color %White% "Follow ALL the ABOVE blue lines.   "
call :dk_color2 %Blue% "Press [1] to Open Support Webpage " %Gray% " Press [0] to Ignore"
choice /C:10 /N
if !errorlevel!==2 exit /b
if !errorlevel!==1 (start %selfgit% & start %github% & for %%# in (%fixes%) do (start %%#))
)

if defined terminal (
call :dk_color %_Yellow% "Press [0] key to %_exitmsg%..."
choice /c 0 /n
) else (
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
pause %nul1%
)
exit /b

::========================================================================================================================================

::  Set variables

:dk_setvar

set ps=%SysPath%\WindowsPowerShell\v1.0\powershell.exe
set psc=%ps% -nop -c
set winbuild=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _slexe=sppsvc.exe& set _slser=sppsvc
if %winbuild% LEQ 6300 (set _slexe=SLsvc.exe& set _slser=SLsvc)
if %winbuild% LSS 7600 if exist "%SysPath%\SLsvc.exe" (set _slexe=SLsvc.exe& set _slser=SLsvc)
if %_slexe%==SLsvc.exe set _vis=1

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 %nul2% | find /i "0x0" %nul1% && (set _NCS=0)

echo "%PROCESSOR_ARCHITECTURE% %PROCESSOR_ARCHITEW6432%" | find /i "ARM64" %nul1% && (if %winbuild% LSS 21277 set ps32onArm=1)

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set   "White="107;91m""
set    "_Red="40;91m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set   "White="White" "Red""
set    "_Red="Black" "Red""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
if %~z0 GEQ 200000 (
set "_exitmsg=Go back"
set "_fixmsg=Go back to Main Menu, select Troubleshoot and run Fix Licensing option."
) else (
set "_exitmsg=Exit"
set "_fixmsg=In MAS folder, run Troubleshoot script and select Fix Licensing option."
)
exit /b

::========================================================================================================================================

::  Extract the text from batch script without character and file encoding issue

:export

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('!_pdesk!\$OEM$\$$\Setup\Scripts\SetupComplete.cmd',$f[1].Trim(),[System.Text.Encoding]::ASCII);"
exit /b

::========================================================================================================================================

:dk_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else if exist %ps% (
%psc% write-host -back '%1' -fore '%2' '%3'
) else if not exist %ps% (
echo %~3
)
exit /b

:dk_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else if exist %ps% (
%psc% write-host -back '%1' -fore '%2' '%3' -NoNewline; write-host -back '%4' -fore '%5' '%6'
) else if not exist %ps% (
echo %~3 %~6
)
exit /b

::========================================================================================================================================
:: Leave empty line below
