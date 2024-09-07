@set masver=2.7
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

set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="r1" set r1=1
if /i "%%#"=="r2" set r2=1
)

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

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
echo Help - %mas%troubleshoot
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
echo Help - %mas%troubleshoot
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

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected [%winbuild%].
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalents.
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

::  Check PowerShell

REM :PowerShellTest: $ExecutionContext.SessionState.LanguageMode :PowerShellTest:

cmd /c "%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':PowerShellTest:\s*';iex ($f[1])"" | find /i "FullLanguage" %nul1% || (
%eline%
cmd /c "%psc% "$ExecutionContext.SessionState.LanguageMode""
echo:
cmd /c "%psc% "$ExecutionContext.SessionState.LanguageMode"" | find /i "FullLanguage" %nul1% && (
echo Failed to run Powershell command but Powershell is working.
call :dk_color %Blue% "Check if your antivirus is blocking the script."
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
) || (
echo PowerShell is not working. Aborting...
echo If you have applied restrictions on Powershell then undo those changes.
echo:
set fixes=%fixes% %mas%fix_powershell
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%fix_powershell"
)
goto done2
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

::  Disable QuickEdit and launch from conhost.exe to avoid Terminal app

if %winbuild% GEQ 17763 (
set terminal=1
) else (
set terminal=
)

::  Check if script is running in Terminal app

set r1=$TB = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0);
set r2=%r1% [void]$TB.DefinePInvokeMethod('GetConsoleWindow', 'kernel32.dll', 22, 1, [IntPtr], @(), 1, 3).SetImplementationFlags(128);
set r3=%r2% [void]$TB.DefinePInvokeMethod('SendMessageW', 'user32.dll', 22, 1, [IntPtr], @([IntPtr], [UInt32], [IntPtr], [IntPtr]), 1, 3).SetImplementationFlags(128);
set d1=%r3% $hIcon = $TB.CreateType(); $hWnd = $hIcon::GetConsoleWindow();
set d2=%d1% echo $($hIcon::SendMessageW($hWnd, 127, 0, 0) -ne [IntPtr]::Zero);

if defined terminal (
%psc% "%d2%" %nul2% | find /i "True" %nul1% && set terminal=
)

if %_unattended%==1 goto :skipQE
for %%# in (%_args%) do (if /i "%%#"=="-qedit" goto :skipQE)

if defined terminal (
set "launchcmd=start conhost.exe %psc%"
) else (
set "launchcmd=%psc%"
)

::  Disable QuickEdit in current session

set "d1=$t=[AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0);"
set "d2=$t.DefinePInvokeMethod('GetStdHandle', 'kernel32.dll', 22, 1, [IntPtr], @([Int32]), 1, 3).SetImplementationFlags(128);"
set "d3=$t.DefinePInvokeMethod('SetConsoleMode', 'kernel32.dll', 22, 1, [Boolean], @([IntPtr], [Int32]), 1, 3).SetImplementationFlags(128);"
set "d4=$k=$t.CreateType(); $b=$k::SetConsoleMode($k::GetStdHandle(-10), 0x0080);"

%launchcmd% "%d1% %d2% %d3% %d4% & cmd.exe '/c' '!_PSarg! -qedit'" && (exit /b) || (set terminal=1)
:skipQE

::========================================================================================================================================

::  Check for updates

set -=
set old=

for /f "delims=[] tokens=2" %%# in ('ping -4 -n 1 updatecheck.mass%-%grave.dev') do (
if not "%%#"=="" (echo "%%#" | find "127.69" %nul1% && (echo "%%#" | find "127.69.%masver%" %nul1% || set old=1))
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
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
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
if not defined terminal mode 78, 30
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
call :dk_color2 %_White% "              [R] " %_Green% "ReadMe"
echo:              [0] Exit
echo:           ________________________________________________________
echo:  
call :dk_color2 %_White% "             " %_Green% "Choose a menu option using your keyboard :"
choice /C:123456789R0 /N
set _erl=%errorlevel%

if %_erl%==11 exit /b
if %_erl%==10 start %mas%oem-folder &goto :Menu
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
call "%~dp0Online_KMS_Activation.cmd" /K-Office
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
call "%~dp0Online_KMS_Activation.cmd" /K-Office
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
call "%~dp0Online_KMS_Activation.cmd" /K-Windows
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
echo The script failed to create the $OEM$ folder.
goto :done2

:done

echo ______________________________________________________________
echo:
call :dk_color %Blue% "%oem%"
call :dk_color %Green% "$OEM$ folder was successfully created on your Desktop."
echo "%oem%" | find /i "38" %nul% && (
echo:
echo To KMS38 activate Server Cor/Acor editions ^(No GUI Versions^),
echo Check this page %mas%oem-folder
)
echo ______________________________________________________________

:done2

echo:
if defined fixes (
call :dk_color2 %Blue% "Press [1] to Open Troubleshoot Page " %Gray% " Press [0] to Ignore"
choice /C:10 /N
if !errorlevel!==1 (for %%# in (%fixes%) do (start %%#))
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
set    "_Red="40;91m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
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
) else (
%psc% write-host -back '%1' -fore '%2' '%3'
)
exit /b

:dk_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
%psc% write-host -back '%1' -fore '%2' '%3' -NoNewline; write-host -back '%4' -fore '%5' '%6'
)
exit /b

::========================================================================================================================================
:: Leave empty line below
