@::63489fhty3-random
@set masver=3.0
@setlocal DisableDelayedExpansion
@echo off



::  For command line switches, check mass<>grave<.>dev/command_line_switches
::  If you want to better understand script, read from MAS separate files version. 



::============================================================================
::
::   Homepage: mass<>grave<.>dev
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

::  Check if Null service is working, it's important for the batch script

sc query Null | find /i "RUNNING"
if %errorlevel% NEQ 0 (
echo:
echo Null service is not running, script may crash...
echo:
echo:
echo Help - %mas%fix_service
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
title  Microsoft_Activation_Scripts %masver%

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

if defined _args echo "%_args%" | find /i "/" >nul && set _unattended=1

::========================================================================================================================================

set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"

call :dk_setvar

if %winbuild% EQU 1 (
%eline%
echo Failed to detect Windows build number.
echo:
setlocal EnableDelayedExpansion
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected [%winbuild%].
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalents.
goto dk_done
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
goto dk_done
)
)

::========================================================================================================================================

::  Check PowerShell

REM :PStest: $ExecutionContext.SessionState.LanguageMode :PStest:

cmd /c "%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':PStest:\s*';iex ($f[1])"" | find /i "FullLanguage" %nul1% || (
%eline%
cmd /c "%psc% "$ExecutionContext.SessionState.LanguageMode""
echo:
cmd /c "%psc% "$ExecutionContext.SessionState.LanguageMode"" | find /i "FullLanguage" %nul1% && (
echo Failed to run Powershell command but Powershell is working.
echo:
cmd /c "%psc% ""$av = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct; $n = @(); foreach ($i in $av) { if ($i.displayName -notlike '*windows*') { $n += $i.displayName } }; if ($n) { Write-Host ('Installed 3rd party Antivirus might be blocking the script - ' + ($n -join ', ')) -ForegroundColor White -BackgroundColor Blue }"""
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
goto dk_done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo Right click on this script and select 'Run as administrator'.
goto dk_done
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

if defined ps32onArm goto :skipQE
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
set upver=%masver:.=%

for /f "delims=[] tokens=2" %%# in ('ping -4 -n 1 activ%-%ated.win') do (
if not "%%#"=="" set old=1
for /f "delims=[] tokens=2" %%# in ('ping -4 -n 1 updatecheck%upver%.activ%-%ated.win') do (
if not "%%#"=="" set old=
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
if !errorlevel!==1 (start %mas% & exit /b)
)
)

::========================================================================================================================================

if not exist "%SystemRoot%\Temp\" mkdir "%SystemRoot%\Temp" %nul%

::  Run script with parameters in unattended mode

set _elev=
if defined _args echo "%_args%" | find /i "/S" %nul% && (set "_silent=%nul%") || (set _silent=)
if defined _args echo "%_args%" | find /i "/" %nul% && (
echo "%_args%" | find /i "/HWID"   %nul% && (setlocal & cls & (call :HWIDActivation    %_args% %_silent%) & endlocal)
echo "%_args%" | find /i "/KMS38"  %nul% && (setlocal & cls & (call :KMS38Activation   %_args% %_silent%) & endlocal)
echo "%_args%" | find /i "/Z-"     %nul% && (setlocal & cls & (call :TSforgeActivation %_args% %_silent%) & endlocal)
echo "%_args%" | find /i "/K-"     %nul% && (setlocal & cls & (call :KMSActivation     %_args% %_silent%) & endlocal)
echo "%_args%" | find /i "/Ohook"  %nul% && (setlocal & cls & (call :OhookActivation   %_args% %_silent%) & endlocal)
exit /b
)

::========================================================================================================================================

setlocal DisableDelayedExpansion

::  Check desktop location

set desktop=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "desktop=%%b"
if not defined desktop for /f "delims=" %%a in ('%psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"
set "_pdesk=%desktop:'=''%"

setlocal EnableDelayedExpansion

if not defined desktop (
%eline%
echo Unable to detect Desktop location, aborting...
goto dk_done
)

::========================================================================================================================================

:MainMenu

cls
color 07
title  Microsoft %blank%Activation %blank%Scripts %masver%
if not defined terminal mode 76, 34

if %winbuild% GEQ 10240 if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _hwidgo=1
if %winbuild% GTR 14393 if exist "%SysPath%\spp\tokens\skus\EnterpriseSN\" set _hwidgo=
if not defined _hwidgo set _tsforgego=1
if %winbuild% GEQ 9200 set _ohookgo=1
if %winbuild% LSS 9200 set _okmsgo=1

echo:
echo:
echo:
echo:
echo:       ______________________________________________________________
echo:
echo:                 Activation Methods:
echo:
if defined _hwidgo (
call :dk_color3 %_White% "             [1] " %_Green% "HWID" %_White% "                - Windows"
) else (
echo:             [1] HWID                - Windows
)
if defined _ohookgo (
call :dk_color3 %_White% "             [2] " %_Green% "Ohook" %_White% "               - Office"
) else (
echo:             [2] Ohook               - Office
)
if defined _tsforgego (
call :dk_color3 %_White% "             [3] " %_Green% "TSforge" %_White% "             - Windows / Office / ESU"
) else (
echo:             [3] TSforge             - Windows / Office / ESU
)
echo:             [4] KMS38               - Windows
if defined _okmsgo (
call :dk_color3 %_White% "             [5] " %_Green% "Online KMS" %_White% "          - Windows / Office"
) else (
echo:             [5] Online KMS          - Windows / Office
)
echo:             __________________________________________________ 
echo:
echo:             [6] Check Activation Status
echo:             [7] Change Windows Edition
echo:             [8] Change Office Edition
echo:             __________________________________________________      
echo:
echo:             [9] Troubleshoot
echo:             [E] Extras
echo:             [H] Help
echo:             [0] Exit
echo:       ______________________________________________________________
echo:
call :dk_color2 %_White% "         " %_Green% "Choose a menu option using your keyboard [1,2,3...E,H,0] :"
choice /C:123456789EH0 /N
set _erl=%errorlevel%

if %_erl%==12 exit /b
if %_erl%==11 start %mas%troubleshoot & goto :MainMenu
if %_erl%==10 goto :Extras
if %_erl%==9 setlocal & call :troubleshoot      & cls & endlocal & goto :MainMenu
if %_erl%==8 setlocal & call :change_offedition & cls & endlocal & goto :MainMenu
if %_erl%==7 setlocal & call :change_winedition & cls & endlocal & goto :MainMenu
if %_erl%==6 setlocal & call :check_actstatus   & cls & endlocal & goto :MainMenu
if %_erl%==5 setlocal & call :KMSActivation     & cls & endlocal & goto :MainMenu
if %_erl%==4 setlocal & call :KMS38Activation   & cls & endlocal & goto :MainMenu
if %_erl%==3 setlocal & call :TSforgeActivation & cls & endlocal & goto :MainMenu
if %_erl%==2 setlocal & call :OhookActivation   & cls & endlocal & goto :MainMenu
if %_erl%==1 setlocal & call :HWIDActivation    & cls & endlocal & goto :MainMenu
goto :MainMenu

:dk_color3

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[%~5%~6%esc%[0m
) else (
%psc% write-host -back '%1' -fore '%2' '%3' -NoNewline; write-host -back '%4' -fore '%5' '%6'-NoNewline; write-host -back '%7' -fore '%8' '%9'
)
exit /b

::========================================================================================================================================

:Extras

cls
title  Extras
if not defined terminal mode 76, 30
echo:
echo:
echo:
echo:
echo:
echo:           ______________________________________________________
echo:           
echo:                [1] Extract $OEM$ Folder
echo:                  
echo:                [2] Download Genuine Windows / Office 
echo:                ____________________________________________      
echo:                                                                          
echo:                [0] Go to Main Menu
echo:           ______________________________________________________
echo:
call :dk_color2 %_White% "             " %_Green% "Choose a menu option using your keyboard [1,2,0] :"
choice /C:120 /N
set _erl=%errorlevel%

if %_erl%==3 goto :MainMenu
if %_erl%==2 start %mas%genuine-installation-media & goto :Extras
if %_erl%==1 goto :Extract$OEM$
goto :Extras

::========================================================================================================================================

:Extract$OEM$

cls
title  Extract $OEM$ Folder
if not defined terminal mode 76, 30

if exist "!desktop!\$OEM$\" (
%eline%
echo $OEM$ folder already exists on the Desktop.
echo _____________________________________________________
echo:
call :dk_color %_Yellow% "Press [0] key to %_exitmsg%..."
choice /c 0 /n
goto :Extras
)

:Extract$OEM$2

cls
title  Extract $OEM$ Folder
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
echo:            [0] Go Back
echo:         ____________________________________________________________
echo:  
call :dk_color2 %_White% "             " %_Green% "Choose a menu option using your keyboard :"
choice /C:12345678R0 /N
set _erl=%errorlevel%

if %_erl%==10 goto:Extras
if %_erl%==9 start %mas%oem-folder &goto:Extract$OEM$2
if %_erl%==8 (set "_oem=TSforge [Windows] + Online KMS [Office]" & set "para=/Z-Windows /K-Office" &goto:Extract$OEM$3)
if %_erl%==7 (set "_oem=HWID [Windows] + Ohook [Office] + TSforge [ESU]" & set "para=/HWID /Ohook /Z-ESU" &goto:Extract$OEM$3)
if %_erl%==6 (set "_oem=HWID [Windows] + Ohook [Office]" & set "para=/HWID /Ohook" &goto:Extract$OEM$3)
if %_erl%==5 (set "_oem=Online KMS" & set "para=/K-WindowsOffice" &goto:Extract$OEM$3)
if %_erl%==4 (set "_oem=KMS38" & set "para=/KMS38" &goto:Extract$OEM$3)
if %_erl%==3 (set "_oem=TSforge" & set "para=/Z-WindowsESUOffice" &goto:Extract$OEM$3)
if %_erl%==2 (set "_oem=Ohook" & set "para=/Ohook" &goto:Extract$OEM$3)
if %_erl%==1 (set "_oem=HWID" & set "para=/HWID" &goto:Extract$OEM$3)
goto :Extract$OEM$2

::========================================================================================================================================

:Extract$OEM$3

cls
set "_dir=!desktop!\$OEM$\$$\Setup\Scripts"
md "!_dir!\"

:: Add random data on top to create unique file which helps in avoiding AV's detections
%psc% "$f=[io.file]::ReadAllText('!_batp!'); [io.file]::WriteAllText('!_pdesk!\$OEM$\$$\Setup\Scripts\MAS_AIO.cmd', '@::RANDOM-' + [Guid]::NewGuid().Guid + [Environment]::NewLine + $f, [System.Text.Encoding]::ASCII)"

(
echo @echo off
echo fltmc ^>nul ^|^| exit /b
echo call "%%~dp0MAS_AIO.cmd" %para%
echo cd \
echo ^(goto^) 2^>nul ^& ^(if "%%~dp0"=="%%SystemRoot%%\Setup\Scripts\" rd /s /q "%%~dp0"^)
)>"!_dir!\SetupComplete.cmd"

set _error=
if not exist "!_dir!\MAS_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1

if defined _error (
%eline%
echo The script failed to create the $OEM$ folder.
if exist "!desktop!\$OEM$\.*" rmdir /s /q "!desktop!\$OEM$\" %nul%
) else (
echo:
call :dk_color %Blue% "%_oem%"
call :dk_color %Green% "$OEM$ folder was successfully created on your Desktop."
)
echo "%_oem%" | find /i "KMS38" 1>nul && (
echo:
echo To KMS38 activate Server Cor/Acor editions ^(No GUI Versions^),
echo Check this page %mas%oem-folder
)
echo ___________________________________________________________________
echo:
call :dk_color %_Yellow% "Press [0] key to %_exitmsg%..."
choice /c 0 /n
goto Extras

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:HWIDActivation

::  To activate, run the script with "/HWID" parameter or change 0 to 1 in below line
set _act=0

::  To disable changing edition if current edition doesn't support HWID activation, change the value to 1 from 0 or run the script with "/HWID-NoEditionChange" parameter
set _NoEditionChange=0

::  If value is changed in above lines or parameter is used then script will run in unattended mode

::========================================================================================================================================

cls
color 07
title  HWID Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/HWID"                  set _act=1
if /i "%%A"=="/HWID-NoEditionChange"  set _NoEditionChange=1
if /i "%%A"=="-el"                    set _elev=1
)
)

for %%A in (%_act% %_NoEditionChange%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

if %winbuild% LSS 10240 (
%eline%
echo Unsupported OS version detected [%winbuild%].
echo HWID Activation is only supported on Windows 10/11.
echo:
call :dk_color %Blue% "Use TSforge activation option from the main menu."
goto dk_done
)

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo HWID Activation is not supported on Windows Server.
call :dk_color %Blue% "Use TSforge activation option from the main menu."
goto dk_done
)

setlocal EnableDelayedExpansion

::========================================================================================================================================

cls
if not defined terminal (
mode 110, 34
if exist "%SysPath%\spp\store_test\" mode 134, 34
)
title  HWID Activation %masver%

echo:
echo Initializing...
call :dk_chkmal

for %%# in (
sppsvc.exe
ClipUp.exe
) do (
if not exist %SysPath%\%%# (
%eline%
echo [%SysPath%\%%#] file is missing, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)
)

::========================================================================================================================================

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_ckeckwmic
call :dk_checksku
call :dk_product
call :dk_sppissue

::========================================================================================================================================

::  Check if system is permanently activated or not

call :dk_checkperm
if defined _perm (
cls
echo ___________________________________________________________________________________________
echo:
call :dk_color2 %_White% "     " %Green% "%winos% is already permanently activated."
echo ___________________________________________________________________________________________
if %_unattended%==1 goto dk_done
echo:
choice /C:10 /N /M ">    [1] Activate Anyway [0] %_exitmsg% : "
if errorlevel 2 exit /b
)
cls

::========================================================================================================================================

::  Check Evaluation version

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% && (
%eline%
echo [%winos% ^| %winbuild%]
echo:
echo Evaluation editions cannot be activated outside of their evaluation period.
call :dk_color %Blue% "Use TSforge activation option from the main menu to reset evaluation period."
echo:
set fixes=%fixes% %mas%evaluation_editions
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%evaluation_editions"
goto dk_done
)
)

::========================================================================================================================================

set error=

cls
echo:
call :dk_showosinfo

::  Check Internet connection

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not "%%#"=="" set _int=1)
)

if not defined _int (
%psc% "If([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet){Exit 0}Else{Exit 1}"
if !errorlevel!==0 (set _int=1&set ping_f= But Ping Failed)
)

if defined _int (
echo Checking Internet Connection            [Connected%ping_f%]
) else (
set error=1
call :dk_color %Red% "Checking Internet Connection            [Not Connected]"
call :dk_color %Blue% "Internet is required for HWID activation."
)

::========================================================================================================================================

echo Initiating Diagnostic Tests...

set "_serv=ClipSVC wlidsvc sppsvc KeyIso LicenseManager Winmgmt"

::  Client License Service (ClipSVC)
::  Microsoft Account Sign-in Assistant
::  Software Protection
::  CNG Key Isolation
::  Windows License Manager Service
::  Windows Management Instrumentation

call :dk_errorcheck

::========================================================================================================================================

::  Detect Key

set key=
set altkey=
set changekey=
set altapplist=
set altedition=
set notworking=

call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f
if defined allapps call :hwiddata key
if not defined key (
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':getactivationid\:.*';iex ($f[1])"') do (set altapplist=%%a)
if defined altapplist call :hwiddata key
)

if defined notworking call :hwidfallback
if not defined key call :hwidfallback

if defined altkey (set key=%altkey%&set changekey=1&set notworking=)

if defined notworking if defined notfoundaltactID (
call :dk_color %Red% "Checking Alternate Edition For HWID     [%altedition% Activation ID Not Found]"
)

if not defined key (
%eline%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
if not defined skunotfound (
echo This product does not support HWID activation.
echo Make sure you are using the latest version of the script.
echo If you are, then try TSforge activation option from the main menu.
set fixes=%fixes% %mas%
echo %mas%
) else (
echo Required license files not found in %SysPath%\spp\tokens\skus\
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)
echo:
goto dk_done
)

if defined notworking set error=1

::========================================================================================================================================

::  Install key

echo:
if defined changekey (
call :dk_color %Blue% "[%altedition%] edition product key will be used to enable HWID activation."
echo:
)

if defined winsub (
call :dk_color %Blue% "Windows Subscription [SKU ID-%slcSKU%] detected. Script will activate base edition [SKU ID-%regSKU%]."
echo:
)

call :dk_inskey "[%key%]"

::========================================================================================================================================

::  Change Windows region to USA to avoid activation issues as Windows store license is not available in many countries 

for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Control Panel\International\Geo" /v Name %nul6%') do set "name=%%b"
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Control Panel\International\Geo" /v Nation %nul6%') do set "nation=%%b"

set regionchange=
if not "%name%"=="US" (
set regionchange=1
%psc% "Set-WinHomeLocation -GeoId 244" %nul%
if !errorlevel! EQU 0 (
echo Changing Windows Region To USA          [Successful]
) else (
call :dk_color %Red% "Changing Windows Region To USA          [Failed]"
)
)

::==========================================================================================================================================

::  Generate GenuineTicket.xml and apply
::  In some cases clipup -v -o method fails and in some cases service restart method fails as well
::  To maximize success rate and get better error details, script will install tickets two times (service restart + clipup -v -o)

set "tdir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if not exist "%tdir%\" md "%tdir%\" %nul%

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" del /f /q "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" %nul%

call :hwiddata ticket

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "Generating GenuineTicket.xml            [Failed, aborting...]"
echo [%encoded%]
if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
goto :dl_final
) else (
echo Generating GenuineTicket.xml            [Successful]
)

set "_xmlexist=if exist "%tdir%\GenuineTicket.xml""

%_xmlexist% (
%psc% "Start-Job { Restart-Service ClipSVC } | Wait-Job -Timeout 20 | Out-Null"
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

%_xmlexist% (
set error=1
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
call :dk_color %Gray% "Installing GenuineTicket.xml            [Failed with ClipSVC service restart, wait...]"
)
)

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%
clipup -v -o

set rebuildinfo=

if not exist %ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat (
set error=1
set rebuildinfo=1
call :dk_color %Red% "Checking ClipSVC tokens.dat             [Not Found]"
)

%_xmlexist% (
set error=1
set rebuildinfo=1
call :dk_color %Red% "Installing GenuineTicket.xml            [Failed With clipup -v -o]"
)

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*.xml" (
set error=1
set rebuildinfo=1
call :dk_color %Red% "Checking Ticket Migration               [Failed]"
)

if not defined altapplist if not defined showfix if defined rebuildinfo (
set showfix=1
call :dk_color %Blue% "%_fixmsg%"
)

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%

::==========================================================================================================================================

call :dk_product

echo:
echo Activating...

call :dk_act
call :dk_checkperm
if defined _perm (
echo:
call :dk_color %Green% "%winos% is permanently activated with a digital license."
goto :dl_final
)

::==========================================================================================================================================

::  Clear store ID related registry to fix activation if Internet is connected

set "_ident=HKU\S-1-5-19\SOFTWARE\Microsoft\IdentityCRL"

if %keyerror% EQU 0 if defined _int (
reg delete "%_ident%" /f %nul%
for %%# in (wlidsvc LicenseManager sppsvc) do (%psc% "Start-Job { Restart-Service %%# } | Wait-Job -Timeout 20 | Out-Null")
call :dk_refresh
call :dk_act
call :dk_checkperm

reg query "%_ident%" %nul% || (
set error=1
echo:
call :dk_color %Red% "Generating New IdentityCRL Registry     [Failed] [%_ident%]"
)
)

::==========================================================================================================================================

::  Extended licensing servers tests incase error not found and activation failed

if %keyerror% EQU 0 if not defined _perm if defined _int (
ipconfig /flushdns %nul%
set "tls=[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;"

for %%# in (
licensing.mp.microsoft.com/v7.0/licenses/content
login.live.com/ppsecure/deviceaddcredential.srf
purchase.mp.microsoft.com/v7.0/users/me/orders
) do if not defined resfail (
%psc% "try { !tls! irm https://%%# -Method POST } catch { if ($_.Exception.Response -eq $null) { Write-Host """"[%%#] $($_.Exception.Message)"""" -ForegroundColor Red -BackgroundColor Black; exit 3 } }"
if !errorlevel!==3 set resfail=1
)
)

if defined resfail (
set error=1
for %%# in (
live.com
microsoft.com
login.live.com
purchase.mp.microsoft.com
licensing.mp.microsoft.com
) do (
findstr /i "%%#" "%SysPath%\drivers\etc\hosts" %nul1% && set "hosfail= [%%# Blocked in Hosts]"
)
call :dk_color %Red% "Checking Licensing Servers              [Failed to Connect]!hosfail!"
set fixes=%fixes% %mas%licensing-servers-issue
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%licensing-servers-issue"
)

::==========================================================================================================================================

::  Windows update and store block check

if %keyerror% EQU 0 if not defined _perm if defined _int (

reg query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /v DisableWindowsUpdateAccess %nul2% | find /i "0x1" %nul% && set wublock=1
reg query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /v DoNotConnectToWindowsUpdateInternetLocations %nul2% | find /i "0x1" %nul% && set wublock=1
if defined wublock (
call :dk_color %Red% "Checking Update Blocker In Registry     [Found]"
call :dk_color %Blue% "HWID activation needs working Windows updates, if you have used any tool to block updates, undo it."
)

reg query "HKLM\SOFTWARE\Policies\Microsoft\WindowsStore" /v DisableStoreApps %nul2% | find /i "0x1" %nul% && (
set storeblock=1
call :dk_color %Red% "Checking Store Blocker In Registry      [Found]"
call :dk_color %Blue% "If you have used any tool to block Store, undo it."
)

set wcount=0
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type ServiceSidType RequiredPrivileges FailureActions) do (
reg query HKLM\SYSTEM\CurrentControlSet\Services\wuauserv /v %%G %nul% || (set wucorrupt=1&set /a wcount+=1)
)

for %%G in (Parameters Security) do (
reg query HKLM\SYSTEM\CurrentControlSet\Services\wuauserv\%%G %nul% || (set wucorrupt=1&set /a wcount+=1)
)

if defined wucorrupt (
set error=1
call :dk_color %Red% "Checking Windows Update Registry        [Corruption Found]"
if !wcount! GTR 2 (
call :dk_color %Red% "Windows seems to be infected with Mal%w%ware."
set fixes=%fixes% %mas%remove_mal%w%ware
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%remove_mal%w%ware"
) else (
call :dk_color %Blue% "HWID activation needs working Windows updates, if you have used any tool to block updates, undo it."
)
) else (
%psc% "Start-Job { Start-Service wuauserv } | Wait-Job -Timeout 20 | Out-Null"
sc query wuauserv | find /i "RUNNING" %nul% || (
set error=1
set wuerror=1
sc start wuauserv %nul%
call :dk_color %Red% "Starting Windows Update Service         [Failed] [!errorlevel!]"
call :dk_color %Blue% "HWID activation needs working Windows updates, if you have used any tool to block updates, undo it."
)
)
)

::==========================================================================================================================================

::  Check Internet related error codes

if %keyerror% EQU 0 if not defined _perm if defined _int (
if not defined wucorrupt if not defined wublock if not defined wuerror if not defined storeblock if not defined resfail (
echo "%error_code%" | findstr /i "0x80072e 0x80072f 0x800704cf 0x87e10bcf 0x800705b4" %nul% && (
call :dk_color %Red% "Checking Internet Issues                [Found] %error_code%"
set fixes=%fixes% %mas%licensing-servers-issue
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%licensing-servers-issue"
)
)
)

::==========================================================================================================================================

echo:
if defined _perm (
call :dk_color %Green% "%winos% is permanently activated with a digital license."
) else (
call :dk_color %Red% "Activation Failed %error_code%"
if defined notworking (
call :dk_color %Blue% "At the time of writing, HWID Activation is not supported for this product."
call :dk_color %Blue% "Use TSforge activation option from the main menu instead."
) else (
if not defined error call :dk_color %Blue% "%_fixmsg%"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)
)

::========================================================================================================================================

:dl_final

echo:

if defined regionchange (
%psc% "Set-WinHomeLocation -GeoId %nation%" %nul%
if !errorlevel! EQU 0 (
echo Restoring Windows Region                [Successful]
) else (
call :dk_color %Red% "Restoring Windows Region                [Failed] [%name% - %nation%]"
)
)

REM if %osSKU%==175 call :dk_color %Red% "%winos% does not support activation on non-azure platforms."

::  Trigger reevaluation of SPP's Scheduled Tasks

if defined _perm (
call :dk_reeval %nul%
)
goto :dk_done

::========================================================================================================================================

::  Set variables

:dk_setvar

set psc=powershell.exe
set winbuild=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

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

::  Show OS info

:dk_showosinfo

for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set osarch=%%b

for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if not "%%j"=="" (
set fullbuild=%%i.%%j
) else (
for /f "tokens=3" %%G in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR" %nul6%') do if not errorlevel 1 set /a "UBR=%%G"
for /f "skip=2 tokens=3,4 delims=. " %%G in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildLabEx') do (
if defined UBR (set "fullbuild=%%G.!UBR!") else (set "fullbuild=%%G.%%H")
)
)

echo Checking OS Info                        [%winos% ^| %fullbuild% ^| %osarch%]
exit /b

::  Check SKU value

:dk_checksku

call :dk_reflection

set osSKU=
set slcSKU=
set wmiSKU=
set regSKU=
set winsub=

if %winbuild% GEQ 14393 (set info=Kernel-BrandingInfo) else (set info=Kernel-ProductInfo)
set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SLGetWindowsInformationDWORD', 'slc.dll', 'Public, Static', 1, [int], @([String], [int].MakeByRefType()), 1, 3);
set d1=%d1% $Sku = 0; [void]$TypeBuilder.CreateType()::SLGetWindowsInformationDWORD('%info%', [ref]$Sku); $Sku
for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set slcSKU=%%s)
set slcSKU=%slcSKU: =%
if "%slcSKU%"=="0" set slcSKU=
for /f "tokens=* delims=0123456789" %%a in ("%slcSKU%") do (if not "[%%a]"=="[]" set slcSKU=)

for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn %nul6%') do set "regSKU=%%a"
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" %nul6%') do if not errorlevel 1 set "wmiSKU=%%a"
if %_wmic% EQU 0 for /f "tokens=1" %%a in ('%psc% "([WMI]'Win32_OperatingSystem=@').OperatingSystemSKU" %nul6%') do if not errorlevel 1 set "wmiSKU=%%a"

if %winbuild% GEQ 15063 %psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':winsubstatus\:.*';iex ($f[1])" %nul2% | find /i "Subscription_is_activated" %nul% && (
if defined regSKU if defined slcSKU if not "%regSKU%"=="%slcSKU%" (
set winsub=1
set osSKU=%regSKU%
)
)

if not defined osSKU set osSKU=%slcSKU%
if not defined osSKU set osSKU=%wmiSKU%
if not defined osSKU set osSKU=%regSKU%
exit /b

::  Get Windows Subscription status

:winsubstatus:
$DM = [AppDomain]::CurrentDomain.DefineDynamicAssembly(6, 1).DefineDynamicModule(4).DefineType(2)
[void]$DM.DefinePInvokeMethod('ClipGetSubscriptionStatus', 'Clipc.dll', 22, 1, [Int32], @([IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
$m = [System.Runtime.InteropServices.Marshal]
$p = $m::AllocHGlobal(12)
$r = $DM.CreateType()::ClipGetSubscriptionStatus([ref]$p)
if ($r -eq 0) {
  $enabled = $m::ReadInt32($p)
  if ($enabled -ge 1) {
    $state = $m::ReadInt32($p, 8)
    if ($state -eq 1) {
        "Subscription_is_activated."
    }
  }
}
:winsubstatus:

::  Get Windows permanent activation status

:dk_checkperm

if %_wmic% EQU 1 wmic path %spp% where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL AND LicenseDependsOn is NULL) get Name /value %nul2% | findstr /i "Windows" %nul1% && set _perm=1||set _perm=
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Name FROM %spp% WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').Get()).Name | %% {echo ('Name='+$_)}" %nul2% | findstr /i "Windows" %nul1% && set _perm=1||set _perm=
exit /b

::  Refresh license status

:dk_refresh

if %_wmic% EQU 1 wmic path %sps% where __CLASS='%sps%' call RefreshLicenseStatus %nul%
if %_wmic% EQU 0 %psc% "$null=(([WMICLASS]'%sps%').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::  Install Key

:dk_inskey

if %_wmic% EQU 1 wmic path %sps% where __CLASS='%sps%' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "try { $null=(([WMISEARCHER]'SELECT Version FROM %sps%').Get()).InstallProductKey('%key%'); exit 0 } catch { exit $_.Exception.InnerException.HResult }" %nul%
set keyerror=%errorlevel%
cmd /c exit /b %keyerror%
if %keyerror% NEQ 0 set "keyerror=[0x%=ExitCode%]"

if %keyerror% EQU 0 (
if %sps%==SoftwareLicensingService call :dk_refresh
echo Installing Generic Product Key          %~1 [Successful]
) else (
call :dk_color %Red% "Installing Generic Product Key          %~1 [Failed] %keyerror%"
if not defined error (
if defined altapplist call :dk_color %Red% "Activation ID not found for this key."
call :dk_color %Blue% "%_fixmsg%"
set showfix=1
)
set error=1
)

exit /b

::  Activation command

:dk_act

set error_code=
if %_wmic% EQU 1 wmic path %spp% where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL" call Activate %nul%
if %_wmic% EQU 0 %psc% "try {$null=(([WMISEARCHER]'SELECT ID FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').Get()).Activate(); exit 0} catch { exit $_.Exception.InnerException.HResult }" %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 (set "error_code=[Error Code: 0x%=ExitCode%]") else (set error_code=)
exit /b

::  Get all products Activation IDs

:dk_actids

set allapps=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='%1') get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %spp% WHERE ApplicationID=''%1''').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined allapps (call set "allapps=!allapps! %%a") else (call set "allapps=%%a"))

::  Check potential script crash issue when user manually installs way too many licenses for Office (length limit in variable)

if defined allapps if %1==0ff1ce15-a989-479d-af46-f275c6370663 (
set len=0
echo:!allapps!> %SystemRoot%\Temp\chklen
for %%A in (%SystemRoot%\Temp\chklen) do (set len=%%~zA)
del %SystemRoot%\Temp\chklen %nul%

if !len! GTR 6000 (
%eline%
echo Too many licenses are installed, the script may crash.
call :dk_color %Blue% "%_fixmsg%"
timeout /t 30
)
)
exit /b

::  Get installed products Activation IDs

:dk_actid

set apps=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='%1' and PartialProductKey is not null) get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %spp% WHERE ApplicationID=''%1'' AND PartialProductKey IS NOT NULL').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined apps (call set "apps=!apps! %%a") else (call set "apps=%%a"))
exit /b

::  Trigger reevaluation, it helps in updating SPP tasks

:dk_reeval

::  This key is left by the system in rearm process and sppsvc sometimes fails to delete it, it causes issues in working of the Scheduled Tasks of SPP

set "ruleskey=HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\PersistedSystemState"
reg delete "%ruleskey%" /v "State" /f %nul%
reg delete "%ruleskey%" /v "SuppressRulesEngine" /f %nul%

set r1=$TB = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0);
set r2=%r1% [void]$TB.DefinePInvokeMethod('SLpTriggerServiceWorker', 'sppc.dll', 22, 1, [Int32], @([UInt32], [IntPtr], [String], [UInt32]), 1, 3);
set d1=%r2% [void]$TB.CreateType()::SLpTriggerServiceWorker(0, 0, 'reeval', 0)
%psc% "Start-Job { Stop-Service sppsvc -force } | Wait-Job -Timeout 20 | Out-Null; %d1%"
exit /b

::  Get Activation IDs from licensing files if not found through WMI

:getactivationid:
$folderPath = "$env:SysPath\spp\tokens\skus"
$files = Get-ChildItem -Path $folderPath -Recurse -Filter "*.xrm-ms"
$guids = @()
foreach ($file in $files) {
    $content = Get-Content -Path $file.FullName -Raw
    $matches = [regex]::Matches($content, 'name="productSkuId">\{([0-9a-fA-F\-]+)\}')
    foreach ($match in $matches) {
        $guids += $match.Groups[1].Value
    }
}
$guids = $guids | Select-Object -Unique
$guidsString = $guids -join " "
$guidsString
:getactivationid:

::  Install License files using Powershell/WMI instead of slmgr.vbs

:xrm:
function InstallLicenseFile($Lsc) {
    try {
        $null = $sls.InstallLicense([IO.File]::ReadAllText($Lsc))
    } catch {
        $host.SetShouldExit($_.Exception.HResult)
    }
}
function InstallLicenseArr($Str) {
    $a = $Str -split ';'
    ForEach ($x in $a) {InstallLicenseFile "$x"}
}
function InstallLicenseDir($Loc) {
    dir $Loc *.xrm-ms -af -s | select -expand FullName | % {InstallLicenseFile "$_"}
}
function ReinstallLicenses() {
    $Oem = "$env:SysPath\oem"
    $Spp = "$env:SysPath\spp\tokens"
    InstallLicenseDir "$Spp"
    If (Test-Path $Oem) {InstallLicenseDir "$Oem"}
}
:xrm:

::  Check wmic.exe

:dk_ckeckwmic

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
cmd /c "wmic path Win32_ComputerSystem get CreationClassName /value" %nul2% | find /i "computersystem" %nul1% && set _wmic=1
)
exit /b

::  Show info for potential script stuck scenario

:dk_sppissue

sc start sppsvc %nul%
set spperror=%errorlevel%

if %spperror% NEQ 1056 if %spperror% NEQ 0 (
%eline%
echo sc start sppsvc [Error Code: %spperror%]
)

echo:
%psc% "$job = Start-Job { (Get-WmiObject -Query 'SELECT * FROM %sps%').Version }; if (-not (Wait-Job $job -Timeout 30)) {write-host 'sppsvc is not working correctly. Help - %mas%troubleshoot'}"
exit /b

::  Get Product name (WMI/REG methods are not reliable in all conditions, hence winbrand.dll method is used)

:dk_product

set d1=%ref% $meth = $TypeBuilder.DefinePInvokeMethod('BrandingFormatString', 'winbrand.dll', 'Public, Static', 1, [String], @([String]), 1, 3);
set d1=%d1% $meth.SetImplementationFlags(128); $TypeBuilder.CreateType()::BrandingFormatString('%%WINDOWS_LONG%%')

set winos=
for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set winos=%%s)
echo "%winos%" | find /i "Windows" %nul1% || (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName %nul6%') do set "winos=%%b"
if %winbuild% GEQ 22000 (
set winos=!winos:Windows 10=Windows 11!
)
)

if not defined winsub exit /b

::  Check base edition product name if Windows subscription license is found

for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b
set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('GetEditionNameFromId', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([int], [IntPtr].MakeByRefType()), 1, 3);
set d1=%d1% $out = 0; [void]$TypeBuilder.CreateType()::GetEditionNameFromId(%regSKU%, [ref]$out);$s=[Runtime.InteropServices.Marshal]::PtrToStringUni($out); $s

for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (
if %winbuild% GEQ 22000 (
set winos=Windows 11 %%a
) else (
set winos=Windows 10 %%a
)
)
exit /b

::  Common lines used in PowerShell reflection code

:dk_reflection

set ref=$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1);
set ref=%ref% $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False);
set ref=%ref% $TypeBuilder = $ModuleBuilder.DefineType(0);
exit /b

::========================================================================================================================================

:dk_chkmal

::  Many users unknowingly download mal-ware by using activators found through Google search.
::  This code aims to notify users that their system has been affected by mal-ware.

set w=
set results=
if exist "%ProgramFiles%\KM%w%Spico" set pupfound= KM%w%Spico 
if not defined pupfound (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks" /f Path /s | find /i "AutoPico" %nul% && set pupfound= KM%w%Spico 
)

set hcount=0
for %%# in (avira.com kaspersky.com virustotal.com mcafee.com) do (
find /i "%%#" %SysPath%\drivers\etc\hosts %nul% && set /a hcount+=1)
if %hcount%==4 set "results=[Antivirus URLs are blocked in hosts]"

sc start sppsvc %nul%
echo "%errorlevel%" | findstr "577 225" %nul% && (
set "results=%results%[Likely File Infector]"
) || (
if not exist %SysPath%\sppsvc.exe if not exist %SysPath%\alg.exe (set "results=%results%[Likely File Infector]")
)

if not "%results%%pupfound%"=="" (
if defined pupfound call :dk_color %Gray% "Checking PUP Activators                 [Found%pupfound%]"
if defined results call :dk_color %Red% "Checking Probable Mal%w%ware Infection..."
if defined results call :dk_color %Red% "%results%"
set fixes=%fixes% %mas%remove_mal%w%ware
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%remove_mal%w%ware"
echo:
)

::  Remove the scheduled task of R@1n-KMS (old version) that runs the activation command every minute, as it leads to high CPU usage.

if exist %SysPath%\Tasks\R@1n-KMS (
for /f %%A in ('dir /b /a:-d %SysPath%\Tasks\R@1n-KMS %nul6%') do (schtasks /delete /tn \R@1n-KMS\%%A /f %nul%)
)

exit /b

::========================================================================================================================================

:dk_errorcheck

set showfix=
call :dk_chkmal

::  Check Sandboxing

sc query Null %nul% || (
set error=1
set showfix=1
call :dk_color %Red% "Checking Sandboxing                     [Found, script may not work properly.]"
call :dk_color %Blue% "If you are using any third-party antivirus, check if it is blocking the script."
echo:
)

::========================================================================================================================================

::  Check corrupt services

set serv_cor=
for %%# in (%_serv%) do (
set _corrupt=
sc start %%# %nul%
if !errorlevel! EQU 1060 set _corrupt=1
sc query %%# %nul% || set _corrupt=1
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type) do if not defined _corrupt (
reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v %%G %nul% || set _corrupt=1
)

if defined _corrupt (if defined serv_cor (set "serv_cor=!serv_cor! %%#") else (set "serv_cor=%%#"))
)

if defined serv_cor (
set error=1
set showfix=1
call :dk_color %Red% "Checking Corrupt Services               [%serv_cor%]"
)

::========================================================================================================================================

::  Check disabled services

set serv_ste=
for %%# in (%_serv%) do (
sc start %%# %nul%
if !errorlevel! EQU 1058 (if defined serv_ste (set "serv_ste=!serv_ste! %%#") else (set "serv_ste=%%#"))
)

::  Change disabled services startup type to default

set serv_csts=
set serv_cste=

if defined serv_ste (
for %%# in (%serv_ste%) do (
if /i %%#==ClipSVC          (reg add "HKLM\SYSTEM\CurrentControlSet\Services\%%#" /v "Start" /t REG_DWORD /d "3" /f %nul% & sc config %%# start= demand %nul%)
if /i %%#==wlidsvc          sc config %%# start= demand %nul%
if /i %%#==sppsvc           (reg add "HKLM\SYSTEM\CurrentControlSet\Services\%%#" /v "Start" /t REG_DWORD /d "2" /f %nul% & sc config %%# start= delayed-auto %nul%)
if /i %%#==KeyIso           sc config %%# start= demand %nul%
if /i %%#==LicenseManager   sc config %%# start= demand %nul%
if /i %%#==Winmgmt          sc config %%# start= auto %nul%
if !errorlevel!==0 (
if defined serv_csts (set "serv_csts=!serv_csts! %%#") else (set "serv_csts=%%#")
) else (
if defined serv_cste (set "serv_cste=!serv_cste! %%#") else (set "serv_cste=%%#")
)
)
)

if defined serv_csts call :dk_color %Gray% "Enabling Disabled Services              [Successful] [%serv_csts%]"

if defined serv_cste (
set error=1
call :dk_color %Red% "Enabling Disabled Services              [Failed] [%serv_cste%]"
)

::========================================================================================================================================

::  Check if the services are able to run or not
::  Workarounds are added to get correct status and error code because sc query doesn't output correct results in some conditions

set serv_e=
for %%# in (%_serv%) do (
set errorcode=
set checkerror=

sc query %%# | find /i "RUNNING" %nul% || (
%psc% "Start-Job { Start-Service %%# } | Wait-Job -Timeout 20 | Out-Null"
set errorcode=!errorlevel!
sc query %%# | find /i "RUNNING" %nul% || set checkerror=1
)

sc start %%# %nul%
if !errorlevel! NEQ 1056 if !errorlevel! NEQ 0 (set errorcode=!errorlevel!&set checkerror=1)
if defined checkerror if defined serv_e (set "serv_e=!serv_e!, %%#-!errorcode!") else (set "serv_e=%%#-!errorcode!")
)

if defined serv_e (
set error=1
call :dk_color %Red% "Starting Services                       [Failed] [%serv_e%]"
echo %serv_e% | findstr /i "ClipSVC-1058 sppsvc-1058" %nul% && (
call :dk_color %Blue% "Reboot your machine using the restart option to fix this error."
set showfix=1
)
echo %serv_e% | findstr /i "sppsvc-1060" %nul% && (
set fixes=%fixes% %mas%fix_service
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%fix_service"
set showfix=1
)
)

::========================================================================================================================================

::  Various error checks

if defined safeboot_option (
set error=1
set showfix=1
call :dk_color2 %Red% "Checking Boot Mode                      [%safeboot_option%] " %Blue% "[Safe mode found. Run in normal mode.]"
)


::  https://learn.microsoft.com/windows-hardware/manufacture/desktop/windows-setup-states

for /f "skip=2 tokens=2*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State" /v ImageState') do (set imagestate=%%B)

if /i not "%imagestate%"=="IMAGE_STATE_COMPLETE" (
call :dk_color %Gray% "Checking Windows Setup State            [%imagestate%]"
echo "%imagestate%" | find /i "RESEAL" %nul% && (
set error=1
set showfix=1
call :dk_color %Blue% "You need to run it in normal mode in case you are running it in Audit Mode."
)
echo "%imagestate%" | find /i "UNDEPLOYABLE" %nul% && (
set fixes=%fixes% %mas%in-place_repair_upgrade
call :dk_color2 %Blue% "If the activation fails, do this - " %_Yellow% " %mas%in-place_repair_upgrade"
)
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinPE" /v InstRoot %nul% && (
set error=1
set showfix=1
call :dk_color2 %Red% "Checking WinPE                          " %Blue% "[WinPE mode found. Run in normal mode.]"
)


set wpainfo=
set wpaerror=
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':wpatest\:.*';iex ($f[1])" %nul6%') do (set wpainfo=%%a)
echo "%wpainfo%" | find /i "Error Found" %nul% && (
set error=1
set wpaerror=1
call :dk_color %Red% "Checking WPA Registry Errors            [%wpainfo%]"
) || (
echo Checking WPA Registry Count             [%wpainfo%]
)


if not defined notwinact if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% || (
call :dk_color %Red% "Checking Eval Packages                  [Non-Eval Licenses are installed in Eval Windows]"
set fixes=%fixes% %mas%evaluation_editions
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%evaluation_editions"
)
)


set osedition=0
if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL) get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
%chkedi% do if not errorlevel 1 (call set "osedition=%%a")

if %osedition%==0 for /f "skip=2 tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul6%') do set "osedition=%%a"

::  Workaround for an issue in builds between 1607 and 1709 where ProfessionalEducation is shown as Professional

if not %osedition%==0 (
if "%osSKU%"=="164" set osedition=ProfessionalEducation
if "%osSKU%"=="165" set osedition=ProfessionalEducationN
)

if not defined notwinact (
if %osedition%==0 (
call :dk_color %Red% "Checking Edition Name                   [Not Found In Registry]"
) else (

if not exist "%SysPath%\spp\tokens\skus\%osedition%\%osedition%*.xrm-ms" if not exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%osedition%\*-%osedition%-*.xrm-ms" (
set skunotfound=1
call :dk_color %Red% "Checking License Files                  [Not Found] [%osedition%]"
)

if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*-%osedition%-*.mum" (
call :dk_color %Red% "Checking Package Files                  [Not Found] [%osedition%]"
)
)
)


%psc% "try { $null=([WMISEARCHER]'SELECT * FROM %sps%').Get().Version; exit 0 } catch { exit $_.Exception.InnerException.HResult }" %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 set "error_code=0x%=ExitCode%"
if %error_code% NEQ 0 (
set error=1
call :dk_color %Red% "Checking SoftwareLicensingService       [Not Working] %error_code%"
)


set wmifailed=
if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value %nul2% | find /i "computersystem" %nul1%
if %_wmic% EQU 0 %psc% "Get-WmiObject -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" %nul2% | find /i "computersystem" %nul1%

if %errorlevel% NEQ 0 set wmifailed=1
echo "%error_code%" | findstr /i "0x800410 0x800440 0x80131501" %nul1% && set wmifailed=1& ::  https://learn.microsoft.com/en-us/windows/win32/wmisdk/wmi-error-constants
if defined wmifailed (
set error=1
call :dk_color %Red% "Checking WMI                            [Not Working]"
if not defined showfix call :dk_color %Blue% "Go back to Main Menu, select Troubleshoot and run Fix WMI option."
set showfix=1
)


if not defined notwinact (
if %winbuild% GEQ 10240 (
%nul% set /a "sum=%slcSKU%+%regSKU%+%wmiSKU%"
set /a "sum/=3"
if not "!sum!"=="%slcSKU%" (
call :dk_color %Gray% "Checking SLC/WMI/REG SKU                [Difference Found - SLC:%slcSKU% WMI:%wmiSKU% Reg:%regSKU%]"
)
) else (
%nul% set /a "sum=%slcSKU%+%wmiSKU%"
set /a "sum/=2"
if not "!sum!"=="%slcSKU%" (
call :dk_color %Gray% "Checking SLC/WMI SKU                    [Difference Found - SLC:%slcSKU% WMI:%wmiSKU%]"
)
)
)

reg query "HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\PersistedTSReArmed" %nul% && (
set error=1
set showfix=1
call :dk_color2 %Red% "Checking Rearm                          " %Blue% "[System Restart Is Required]"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState" %nul% && (
set error=1
set showfix=1
call :dk_color2 %Red% "Checking ClipSVC                        " %Blue% "[System Restart Is Required]"
)


::  This "WLMS" service was included in previous Eval editions (which were activable) to automatically shut down the system every hour after the evaluation period expired and prevent SPPSVC from stopping.

if exist "%SysPath%\wlms\wlms.exe" (
echo Checking Eval WLMS Service              [Found]
)


reg query "HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion" %nul% || (
set error=1
set showfix=1
call :dk_color %Red% "Checking HKU\S-1-5-20 Registry          [Not Found]"
set fixes=%fixes% %mas%in-place_repair_upgrade
call :dk_color2 %Blue% "In case of activation issues, do this - " %_Yellow% " %mas%in-place_repair_upgrade"
)


for %%# in (SppEx%w%tComObj.exe sppsvc.exe sppsvc.exe\PerfOptions) do (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%w%ge File Execu%w%tion Options\%%#" %nul% && (if defined _sppint (set "_sppint=!_sppint!, %%#") else (set "_sppint=%%#"))
)
if defined _sppint (
echo %_sppint% | find /i "PerfOptions" %nul% && (
call :dk_color %Red% "Checking SPP Interference In IFEO       [%_sppint% - System might deactivate later]"
if not defined showfix call :dk_color %Blue% "%_fixmsg%"
set showfix=1
) || (
echo Checking SPP In IFEO                    [%_sppint%]
)
)


for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" %nul6%') do if /i %%b NEQ 0x0 (
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" /t REG_DWORD /d "0" /f %nul%
call :dk_color %Red% "Checking SkipRearm                      [Default 0 Value Not Found. Changing To 0]"
%psc% "Start-Job { Stop-Service sppsvc -force } | Wait-Job -Timeout 20 | Out-Null"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Plugins\Objects\msft:rm/algorithm/hwid/4.0" /f ba02fed39662 /d %nul% || (
call :dk_color %Red% "Checking SPP Registry Key               [Incorrect ModuleId Found]"
set fixes=%fixes% %mas%issues_due_to_gaming_spoofers
call :dk_color2 %Blue% "Most likely caused by gaming spoofers. Help - " %_Yellow% " %mas%issues_due_to_gaming_spoofers"
set error=1
set showfix=1
)


set tokenstore=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore %nul6%') do call set "tokenstore=%%b"
if %winbuild% LSS 9200 set "tokenstore=%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform"
if %winbuild% GEQ 9200 if /i not "%tokenstore%"=="%SysPath%\spp\store" if /i not "%tokenstore%"=="%SysPath%\spp\store\2.0" if /i not "%tokenstore%"=="%SysPath%\spp\store_test\2.0" (
set toerr=1
set error=1
set showfix=1
call :dk_color %Red% "Checking TokenStore Registry Key        [Correct Path Not Found] [%tokenstore%]"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)


::  This code creates token folder only if it's missing and sets default permission for it

if not defined toerr if not exist "%tokenstore%\" (
mkdir "%tokenstore%" %nul%
if %winbuild% LSS 9200 set "d=$sddl = 'O:NSG:NSD:AI(A;OICIID;FA;;;SY)(A;OICIID;FA;;;BA)(A;OICIID;FA;;;NS)';"
if %winbuild% GEQ 9200 set "d=$sddl = 'O:BAG:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICIIO;GR;;;BU)(A;;FR;;;BU)(A;OICI;FA;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)';"
set "d=!d! $AclObject = New-Object System.Security.AccessControl.DirectorySecurity;"
set "d=!d! $AclObject.SetSecurityDescriptorSddlForm($sddl);"
set "d=!d! Set-Acl -Path %tokenstore% -AclObject $AclObject;"
%psc% "!d!" %nul%
if exist "%tokenstore%\" (
call :dk_color %Gray% "Checking SPP Token Folder               [Not Found, Created Now] [%tokenstore%\]"
) else (
call :dk_color %Red% "Checking SPP Token Folder               [Not Found, Failed to Create] [%tokenstore%\]"
set error=1
set showfix=1
)
)


if not defined notwinact (
call :dk_actid 55c92734-d682-4d71-983e-d6ec3f16059f
if not defined apps (
%psc% "Start-Job { Stop-Service sppsvc -force } | Wait-Job -Timeout 20 | Out-Null; $sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
call :dk_actid 55c92734-d682-4d71-983e-d6ec3f16059f
if not defined apps (
set "_notfoundids=Key Not Installed / Act ID Not Found"
call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f
if not defined allapps (
set error=1
set "_notfoundids=Not found"
)
call :dk_color %Red% "Checking Activation IDs                 [!_notfoundids!]"
)
)
)


if exist "%tokenstore%\" if not exist "%tokenstore%\tokens.dat" (
set error=1
call :dk_color %Red% "Checking SPP tokens.dat                 [Not Found] [%tokenstore%\]"
)


if %winbuild% GEQ 9200 if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
%psc% "Get-WmiObject -Query 'SELECT Description FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL AND LicenseDependsOn IS NULL' | Select-Object -Property Description" %nul2% | findstr /i "KMS_" %nul1% || (
for /f "delims=" %%a in ('%psc% "(Get-ScheduledTask -TaskName 'SvcRestartTask' -TaskPath '\Microsoft\Windows\SoftwareProtectionPlatform\').State" %nul6%') do (set taskinfo=%%a)
echo !taskinfo! | find /i "Ready" %nul% || (
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "actionlist" /f %nul%
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\SoftwareProtectionPlatform\SvcRestartTask" %nul% || set taskinfo=Removed
if "!taskinfo!"=="" set "taskinfo=Not Found"
call :dk_color %Red% "Checking SvcRestartTask Status          [!taskinfo!, System might deactivate later]"
if not defined error call :dk_color %Blue% "Reboot your machine using the restart option."
)
)
)


::  This code checks if SPP has permission access to tokens folder and required registry keys. It's often caused by gaming spoofers.

set permerror=
if %winbuild% GEQ 9200 if not defined ps32onArm (
for %%# in (
"%tokenstore%+FullControl"
"HKLM:\SYSTEM\WPA+QueryValues, EnumerateSubKeys, WriteKey"
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform+SetValue"
) do for /f "tokens=1,2 delims=+" %%A in (%%#) do if not defined permerror (
%psc% "$acl = (Get-Acl '%%A' | fl | Out-String); if (-not ($acl -match 'NT SERVICE\\sppsvc Allow  %%B') -or ($acl -match 'NT SERVICE\\sppsvc Deny')) {Exit 2}" %nul%
if !errorlevel!==2 (
if "%%A"=="%tokenstore%" (
set "permerror=Error Found In Token Folder"
) else (
set "permerror=Error Found In SPP Registries"
)
)
)

REM  https://learn.microsoft.com/office/troubleshoot/activation/license-issue-when-start-office-application

if not defined permerror (
reg query "HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion" %nul% && (
set "pol=HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies"
reg query "!pol!" %nul% || reg add "!pol!" %nul%
%psc% "$netServ = (New-Object Security.Principal.SecurityIdentifier('S-1-5-20')).Translate([Security.Principal.NTAccount]).Value; $aclString = Get-Acl 'Registry::!pol!' | Format-List | Out-String; if (-not ($aclString.Contains($netServ + ' Allow  FullControl') -or $aclString.Contains('NT SERVICE\sppsvc Allow  FullControl')) -or ($aclString.Contains('Deny'))) {Exit 3}" %nul%
if !errorlevel!==3 set "permerror=Error Found In S-1-5-20 SPP"
)
)

if defined permerror (
set error=1
call :dk_color %Red% "Checking SPP Permissions                [!permerror!]"
if not defined showfix call :dk_color %Blue% "%_fixmsg%"
set showfix=1
)
)


::  If required services are not disabled or corrupted + if there is any error + SoftwareLicensingService errorlevel is not Zero + no fix was shown before

if not defined serv_cor if not defined serv_cste if defined error if /i not %error_code%==0 if not defined showfix (
if not defined permerror if defined wpaerror (call :dk_color %Blue% "Go back to Main Menu, select Troubleshoot and run Fix WPA Registry option." & set showfix=1)
if not defined showfix (
set showfix=1
call :dk_color %Blue% "%_fixmsg%"
if not defined permerror call :dk_color %Blue% "If activation still fails then run Fix WPA Registry option."
)
)

if not defined showfix if defined wpaerror (
set showfix=1
call :dk_color %Blue% "If activation fails then go back to Main Menu, select Troubleshoot and run Fix WPA Registry option."
)

exit /b

::  This code checks for invalid registry keys in HKLM\SYSTEM\WPA. This issue may appear even on healthy systems

:wpatest:
$wpaKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $env:COMPUTERNAME).OpenSubKey("SYSTEM\\WPA")
$count = 0
foreach ($subkeyName in $wpaKey.GetSubKeyNames()) {
    if ($subkeyName -match '.*-.*-.*-.*-.*-') {
        $count++
    }
}
$osVersion = [System.Environment]::OSVersion.Version
$minBuildNumber = 14393
if ($osVersion.Build -ge $minBuildNumber) {
    $subkeyHashTable = @{}
    foreach ($subkeyName in $wpaKey.GetSubKeyNames()) {
        if ($subkeyName -match '.*-.*-.*-.*-.*-') {
            $keyNumber = $subkeyName -replace '.*-', ''
            $subkeyHashTable[$keyNumber] = $true
        }
    }
    for ($i=1; $i -le $count; $i++) {
        if (-not $subkeyHashTable.ContainsKey("$i")) {
            Write-Output "Total Keys $count. Error Found - $i key does not exist."
			$wpaKey.Close()
			exit
        }
    }
}
$wpaKey.GetSubKeyNames() | ForEach-Object {
    if ($_ -match '.*-.*-.*-.*-.*-') {
        if ($PSVersionTable.PSVersion.Major -lt 3) {
            cmd /c "reg query "HKLM\SYSTEM\WPA\$_" /ve /t REG_BINARY >nul 2>&1"
			if ($LASTEXITCODE -ne 0) {
            Write-Host "Total Keys $count. Error Found - Binary Data is corrupt."
			$wpaKey.Close()
			exit
			}
        } else {
            $subkey = $wpaKey.OpenSubKey($_)
            $p = $subkey.GetValueNames()
            if (($p | Where-Object { $subkey.GetValueKind($_) -eq [Microsoft.Win32.RegistryValueKind]::Binary }).Count -eq 0) {
                Write-Host "Total Keys $count. Error Found - Binary Data is corrupt."
				$wpaKey.Close()
				exit
            }
        }
    }
}
$count
$wpaKey.Close()
:wpatest:

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

:dk_done

echo:
if %_unattended%==1 timeout /t 2 & exit /b

if defined fixes (
call :dk_color %White% "Follow ALL the ABOVE blue lines.   "
call :dk_color2 %Blue% "Press [1] to Open Support Webpage " %Gray% " Press [0] to Ignore"
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

::  1st column = Activation ID
::  2nd column = Generic Retail/OEM/MAK Key
::  3rd column = SKU ID
::  4th column = Key part number
::  5th column = Ticket signature value. It's as it is, it's not encoded. (Check mass grave[.]dev/hwid#manual-activation to see how it's generated)
::  6th column = 1 = activation is not working (at the time of writing this), 0 = activation is working
::  7th column = Key Type
::  8th column = WMI Edition ID (For reference only)
::  9th column = Version name incase same Edition ID is used in different OS versions with different key
::  Separator  = _


:hwiddata

set f=
for %%# in (
8b351c9c-f398-4515-9900-09df49427262_XGVPP-NMH47-7TTHJ-W3FW7-8H%f%V2C___4_X19-99683_HGNKjkKcKQHO6n8srMUrDh/MElffBZarLqCMD9rWtgFKf3YzYOLDPEMGhuO/auNMKCeiU7ebFbQALS/MyZ7TvidMQ2dvzXeXXKzPBjfwQx549WJUU7qAQ9Txg9cR9SAT8b12Pry2iBk+nZWD9VtHK3kOnEYkvp5WTCTsrSi6Re4_0_OEM:NONSLP_Enterprise
c83cef07-6b72-4bbc-a28f-a00386872839_3V6Q6-NQXCX-V8YXR-9QCYV-QP%f%FCT__27_X19-98746_NHn2n0N1UfVf00CfaI5LCDMDsKdVAWpD/HAfUrcTAKsw9d2Sks4h5MhyH/WUx+B6dFi8ol7D3AHorR8y9dqVS1Bd2FdZNJl/tTR1PGwYn6KL88NS19aHmFNdX8s4438vaa+Ty8Qk8EDcwm/wscC8lQmi3/RgUKYdyGFvpbGSVlk_0_Volume:MAK_EnterpriseN
4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7JG-NPHTM-C97JM-9MPGT-3V%f%66T__48_X19-98841_Yl/jNfxJ1SnaIZCIZ4m6Pf3ySNoQXifNeqfltNaNctx+onwiivOx7qcSn8dFtURzgMzSOFnsRQzb5IrvuqHoxWWl1S3JIQn56FvKsvSx7aFXIX3+2Q98G1amPV/WEQ0uHA5d7Ya6An+g0Z0zRP7evGoomTs4YuweaWiZQjQzSpA_0_____Retail_Professional
9fbaf5d6-4d83-4422-870d-fdda6e5858aa_2B87N-8KFHP-DKV6R-Y2C8J-PK%f%CKT__49_X19-98859_Ge0mRQbW8ALk7T09V+1k1yg66qoS0lhkgPIROOIOgxKmWPAvsiLAYPKDqM4+neFCA/qf1dHFmdh0VUrwFBPYsK251UeWuElj4bZFVISL6gUt1eZwbGfv5eurQ0i+qZiFv+CcQOEFsd5DD4Up6xPLLQS3nAXODL5rSrn2sHRoCVY_0_____Retail_ProfessionalN
f742e4ff-909d-4fe9-aacb-3231d24a0c58_4CPRK-NM3K3-X6XXQ-RXX86-WX%f%CHW__98_X19-98877_vel4ytVtnE8FhvN87Cflz9sbh5QwHD1YGOeej9QP7hF3vlBR4EX2/S/09gRneeXVbQnjDOCd2KFMKRUWHLM7ZhFBk8AtlG+kvUawPZ+CIrwrD3mhi7NMv8UX/xkLK3HnBupMEuEwsMJgCUD8Pn6om1mEiQebHBAqu4cT7GN9Y0g_0_____Retail_CoreN
1d1bac85-7365-4fea-949a-96978ec91ae0_N2434-X9D7W-8PF6X-8DV9T-8T%f%YMD__99_X19-99652_Nv17eUTrr1TmUX6frlI7V69VR6yWb7alppCFJPcdjfI+xX4/Cf2np3zm7jmC+zxFb9nELUs477/ydw2KCCXFfM53bKpBQZKHE5+MdGJGxebOCcOtJ3hrkDJtwlVxTQmUgk5xnlmpk8PHg82M2uM5B7UsGLxGKK4d3hi0voSyKeI_0_____Retail_CoreCountrySpecific
3ae2cc14-ab2d-41f4-972f-5e20142771dc_BT79Q-G7N6G-PGBYW-4YWX6-6F%f%4BT_100_X19-99661_FV2Eao/R5v8sGrfQeOjQ4daokVlNOlqRCDZXuaC45bQd5PsNU3t1b4AwWeYM8TAwbHauzr4tPG0UlsUqUikCZHy0poROx35bBBMBym6Zbm9wDBVyi7nCzBtwS86eOonQ3cU6WfZxhZRze0POdR33G3QTNPrnVIM2gf6nZJYqDOA_0_____Retail_CoreSingleLanguage
2b1f36bb-c1cd-4306-bf5c-a0367c2d97d8_YTMG3-N6DKC-DKB77-7M9GH-8H%f%VX7_101_X19-98868_GH/jwFxIcdQhNxJIlFka8c1H48PF0y7TgJwaryAUzqSKXynONLw7MVciDJFVXTkCjbXSdxLSWpPIC50/xyy1rAf8aC7WuN/9cRNAvtFPC1IVAJaMeq1vf4mCqRrrxJQP6ZEcuAeHFzLe/LLovGWCd8rrs6BbBwJXCvAqXImvycQ_0_____Retail_Core
2a6137f3-75c0-4f26-8e3e-d83d802865a4_XKCNC-J26Q9-KFHD2-FKTHY-KD%f%72Y_119_X19-99606_hci78IRWDLBtdbnAIKLDgV9whYgtHc1uYyp9y6FszE9wZBD5Nc8CUD2pI2s2RRd3M04C4O7M3tisB3Ov/XVjpAbxlX3MWfUR5w4MH0AphbuQX0p5MuHEDYyfqlRgBBRzOKePF06qfYvPQMuEfDpKCKFwNojQxBV8O0Arf5zmrIw_0_OEM:NONSLP_PPIPro
e558417a-5123-4f6f-91e7-385c1c7ca9d4_YNMGQ-8RYV3-4PGQ3-C8XTP-7C%f%FBY_121_X19-98886_x9tPFDZmjZMf29zFeHV5SHbXj8Wd8YAcCn/0hbpLcId4D7OWqkQKXxXHIegRlwcWjtII0sZ6WYB0HQV2KH3LvYRnWKpJ5SxeOgdzBIJ6fhegYGGyiXsBv9sEb3/zidPU6ZK9LugVGAcRZ6HQOiXyOw+Yf5H35iM+2oDZXSpjvJw_0_____Retail_Education
c5198a66-e435-4432-89cf-ec777c9d0352_84NGF-MHBT6-FXBX8-QWJK7-DR%f%R8H_122_X19-98892_jkL4YZkmBCJtvL1fT30ZPBcjmzshBSxjwrE0Q00AZ1hYnhrH+npzo1MPCT6ZRHw19ZLTz7wzyBb0qqcBVbtEjZW0Xs2MYLxgriyoONkhnPE6KSUJBw7C0enFVLHEqnVu/nkaOFfockN3bc+Eouw6W2lmHjklPHc9c6Clo04jul0_0_____Retail_EducationN
f6e29426-a256-4316-88bf-cc5b0f95ec0c_PJB47-8PN2T-MCGDY-JTY3D-CB%f%CPV_125_X23-50331_OPGhsyx+Ctw7w/KLMRNrY+fNBmKPjUG0R9RqkWk4e8ez+ExSJxSLLex5WhO5QSNgXLmEra+cCsN6C638aLjIdH2/L7D+8z/C6EDgRvbHMmidHg1lX3/O8lv0JudHkGtHJYewjorn/xXGY++vOCTQdZNk6qzEgmYSvPehKfdg8js_1_Volume:MAK_EnterpriseS_Ge
cce9d2de-98ee-4ce2-8113-222620c64a27_KCNVH-YKWX8-GJJB9-H9FDT-6F%f%7W2_125_X22-66075_GCqWmJOsTVun9z4QkE9n2XqBvt3ZWSPl9QmIh9Q2mXMG/QVt2IE7S+ES/NWlyTSNjLVySr1D2sGjxgEzy9kLwn7VENQVJ736h1iOdMj/3rdqLMSpTa813+nPSQgKpqJ3uMuvIvRP0FdB7Y4qt8qf9kNKK25A1QknioD/6YubL/4_1_Volume:MAK_EnterpriseS_VB
d06934ee-5448-4fd1-964a-cd077618aa06_43TBQ-NH92J-XKTM7-KT3KK-P3%f%9PB_125_X21-83233_EpB6qOCo8pRgO5kL4vxEHck2J1vxyd9OqvxUenDnYO9AkcGWat/D74ZcFg5SFlIya1U8l5zv+tsvZ4wAvQ1IaFW1PwOKJLOaGgejqZ41TIMdFGGw+G+s1RHsEnrWr3UOakTodby1aIMUMoqf3NdaM5aWFo8fOmqWC5/LnCoighs_0_OEM:NONSLP_EnterpriseS_RS5
706e0cfd-23f4-43bb-a9af-1a492b9f1302_NK96Y-D9CD8-W44CQ-R8YTK-DY%f%JWX_125_X21-05035_ntcKmazIvLpZOryft28gWBHu1nHSbR+Gp143f/BiVe+BD2UjHBZfSR1q405xmQZsygz6VRK6+zm8FPR++71pkmArgCLhodCQJ5I4m7rAJNw/YX99pILphi1yCRcvHsOTGa825GUVXgf530tHT6hr0HQ1lGeGgG1hPekpqqBbTlg_0_OEM:NONSLP_EnterpriseS_RS1
faa57748-75c8-40a2-b851-71ce92aa8b45_FWN7H-PF93Q-4GGP8-M8RF3-MD%f%WWW_125_X19-99617_Fe9CDClilrAmwwT7Yhfx67GafWRQEpwyj8R+a4eaTqbpPcAt7d1hv1rx8Sa9AzopEGxIrb7IhiPoDZs0XaT1HN0/olJJ/MnD73CfBP4sdQdLTsSJE3dKMWYTQHpnjqRaS/pNBYRr8l9Mv8yfcP8uS2MjIQ1cRTqRmC7WMpShyCg_0_OEM:NONSLP_EnterpriseS_TH
3d1022d8-969f-4222-b54b-327f5a5af4c9_2DBW3-N2PJG-MVHW3-G7TDK-9H%f%KR4_126_X21-04921_zLPNvcl1iqOefy0VLg+WZgNtRNhuGpn8+BFKjMqjaNOSKiuDcR6GNDS5FF1Aqk6/e6shJ+ohKzuwrnmYq3iNQ3I2MBlYjM5kuNfKs8Vl9dCjSpQr//GBGps6HtF2xrG/2g/yhtYC7FbtGDIE16uOeNKFcVg+XMb0qHE/5Etyfd8_0_Volume:MAK_EnterpriseSN_RS1
60c243e1-f90b-4a1b-ba89-387294948fb6_NTX6B-BRYC2-K6786-F6MVQ-M7%f%V2X_126_X19-98770_kbXfe0z9Vi1S0yfxMWzI5+UtWsJKzxs7wLGUDLjrckFDn1bDQb4MvvuCK1w+Qrq33lemiGpNDspa+ehXiYEeSPFcCvUBpoMlGBFfzurNCHWiv3o1k3jBoawJr/VoDoVZfxhkps0fVoubf9oy6C6AgrkZ7PjCaS58edMcaUWvYYg_0_Volume:MAK_EnterpriseSN_TH
01eb852c-424d-4060-94b8-c10d799d7364_3XP6D-CRND4-DRYM2-GM84D-4G%f%G8Y_139_X23-37869_PVW0XnRJnsWYjTqxb6StCi2tge/uUwegjdiFaFUiZpwdJ620RK+MIAsSq5S+egXXzIWNntoy2fB6BO8F1wBFmxP/mm/3rn5C33jtF5QrbNqY7X9HMbqSiC7zhs4v4u2Xa4oZQx8JQkwr8Q2c/NgHrOJKKRASsSckhunxZ+WVEuM_1_____Retail_ProfessionalCountrySpecific_Zn
eb6d346f-1c60-4643-b960-40ec31596c45_DXG7C-N36C4-C4HTG-X4T3X-2Y%f%V77_161_X21-43626_MaVqTkRrGnOqYizl15whCOKWzx01+BZTVAalvEuHXM+WV55jnIfhWmd/u1GqCd5OplqXdU959zmipK2Iwgu2nw/g91nW//sQiN/cUcvg1Lxo6pC3gAo1AjTpHmGIIf9XlZMYlD+Vl6gXsi/Auwh3yrSSFh5s7gOczZoDTqQwHXA_0_____Retail_ProfessionalWorkstation
89e87510-ba92-45f6-8329-3afa905e3e83_WYPNQ-8C467-V2W6J-TX4WX-WT%f%2RQ_162_X21-43644_JVGQowLiCcPtGY9ndbBDV+rTu/q5ljmQTwQWZgBIQsrAeQjLD8jLEk/qse7riZ7tMT6PKFVNXeWqF7PhLAmACbE8O3Lvp65XMd/Oml9Daynj5/4n7unsffFHIHH8TGyO5j7xb4dkFNqC5TX3P8/1gQEkTIdZEOTQQXFu0L2SP5c_0_____Retail_ProfessionalWorkstationN
62f0c100-9c53-4e02-b886-a3528ddfe7f6_8PTT6-RNW4C-6V7J2-C2D3X-MH%f%BPB_164_X21-04955_CEDgxI8f/fxMBiwmeXw5Of55DG32sbGALzHihXkdbYTDaE3pY37oAA4zwGHALzAFN/t254QImGPYR6hATgl+Cp804f7serJqiLeXY965Zy67I4CKIMBm49lzHLFJeDnVTjDB0wVyN29pvgO3+HLhZ22KYCpkRHFFMy2OKxS68Yc_0_____Retail_ProfessionalEducation
13a38698-4a49-4b9e-8e83-98fe51110953_GJTYN-HDMQY-FRR76-HVGC7-QP%f%F8P_165_X21-04956_r35zp9OfxKSBcTxKWon3zFtbOiCufAPo6xRGY5DJqCRFKdB0jgZalNQitvjmaZ/Rlez2vjRJnEart4LrvyW4d9rrukAjR3+c3UkeTKwoD3qBl9AdRJbXCa2BdsoXJs1WVS4w4LuVzpB/SZDuggZt0F2DlMB427F5aflook/n1pY_0_____Retail_ProfessionalEducationN
df96023b-dcd9-4be2-afa0-c6c871159ebe_NJCF7-PW8QT-3324D-688JX-2Y%f%V66_175_X21-41295_rVpetYUmiRB48YJfCvJHiaZapJ0bO8gQDRoql+rq5IobiSRu//efV1VXqVpBkwILQRKgKIVONSTUF5y2TSxlDLbDSPKp7UHfbz17g6vRKLwOameYEz0ZcK3NTbApN/cMljHvvF/mBag1+sHjWu+eoFzk8H89k9nw8LMeVOPJRDc_0_____Retail_ServerRdsh
d4ef7282-3d2c-4cf0-9976-8854e64a8d1e_V3WVW-N2PV2-CGWC3-34QGF-VM%f%J2C_178_X21-32983_Xzme9hDZR6H0Yx0deURVdE6LiTOkVqWng5W/OTbkxRc0rq+mSYpo/f/yqhtwYlrkBPWx16Yok5Bvcb34vbKHvEAtxfYp4te20uexLzVOtBcoeEozARv4W/6MhYfl+llZtR5efsktj4N4/G4sVbuGvZ9nzNfQO9TwV6NGgGEj2Ec_0_____Retail_Cloud
af5c9381-9240-417d-8d35-eb40cd03e484_NH9J3-68WK7-6FB93-4K3DF-DJ%f%4F6_179_X21-32987_QGRDZOU/VZhYLOSdp2xDnFs8HInNZctcQlWCIrORVnxTQr55IJwN4vK3PJHjkfRLQ/bgUrcEIhyFbANqZFUq8yD1YNubb2bjNORgI/m8u85O9V7nDGtxzO/viEBSWyEHnrzLKKWYqkRQKbbSW3ungaZR0Ti5O2mAUI4HzAFej50_0_____Retail_CloudN
8ab9bdd1-1f67-4997-82d9-8878520837d9_XQQYW-NFFMW-XJPBH-K8732-CK%f%FFD_188_X21-99378_djy0od0uuKd2rrIl+V1/2+MeRltNgW7FEeTNQsPMkVSL75NBphgoso4uS0JPv2D7Y1iEEvmVq6G842Kyt52QOwXgFWmP/IQ6Sq1dr+fHK/4Et7bEPrrGBEZoCfWqk0kdcZRPBij2KN6qCRWhrk1hX2g+U40smx/EYCLGh9HCi24_0_____OEM:DM_IoTEnterprise
ed655016-a9e8-4434-95d9-4345352c2552_QPM6N-7J2WJ-P88HH-P3YRH-YY%f%74H_191_X21-99682_qHs/PzfhYWdtSys2edzcz4h+Qs8aDqb8BIiQ/mJ/+0uyoJh1fitbRCIgiFh2WAGZXjdgB8hZeheNwHibd8ChXaXg4u+0XlOdFlaDTgTXblji8fjETzDBk9aGkeMCvyVXRuUYhTSdp83IqGHz7XuLwN2p/6AUArx9JZCoLGV8j3w_0_OEM:NONSLP_IoTEnterpriseS_VB
6c4de1b8-24bb-4c17-9a77-7b939414c298_CGK42-GYN6Y-VD22B-BX98W-J8%f%JXD_191_X23-12617_J/fpIRynsVQXbp4qZNKp6RvOgZ/P2klILUKQguMlcwrBZybwNkHg/kM5LNOF/aDzEktbPnLnX40GEvKkYT6/qP4cMhn/SOY0/hYOkIdR34ilzNlVNq5xP7CMjCjaUYJe+6ydHPK6FpOuEoWOYYP5BZENKNGyBy4w4shkMAw19mA_0_OEM:NONSLP_IoTEnterpriseS_Ge
d4bdc678-0a4b-4a32-a5b3-aaa24c3b0f24_K9VKN-3BGWV-Y624W-MCRMQ-BH%f%DCD_202_X22-53884_kyoNx2s93U6OUSklB1xn+GXcwCJO1QTEtACYnChi8aXSoxGQ6H2xHfUdHVCwUA1OR0UeNcRrMmOzZBOEUBtdoGWSYPg9AMjvxlxq9JOzYAH+G6lT0UbCWgMSGGrqdcIfmshyEak3aUmsZK6l+uIAFCCZZ/HbbCRkkHC5rWKstMI_0_____Retail_CloudEditionN
92fb8726-92a8-4ffc-94ce-f82e07444653_KY7PN-VR6RX-83W6Y-6DDYQ-T6%f%R4W_203_X22-53847_gD6HnT4jP4rcNu9u83gvDiQq1xs7QSujcDbo60Di5iSVa9/ihZ7nlhnA0eDEZfnoDXriRiPPqc09T6AhSnFxLYitAkOuPJqL5UMobIrab9dwTKlowqFolxoHhLOO4V92Hsvn/9JLy7rEzoiAWHhX/0cpMr3FCzVYPeUW1OyLT1A_0_____Retail_CloudEdition
5a85300a-bfce-474f-ac07-a30983e3fb90_N979K-XWD77-YW3GB-HBGH6-D3%f%2MH_205_X23-15042_blZopkUuayCTgZKH4bOFiisH9GTAHG5/js6UX/qcMWWc3sWNxKSX1OLp1k3h8Xx1cFuvfG/fNAw/I83ssEtPY+A0Gx1JF4QpRqsGOqJ5ruQ2tGW56CJcCVHkB+i46nJAD759gYmy3pEYMQbmpWbhLx3MJ6kvwxKfU+0VCio8k50_0_____OEM:DM_IoTEnterpriseSK
80083eae-7031-4394-9e88-4901973d56fe_P8Q7T-WNK7X-PMFXY-VXHBG-RR%f%K69_206_X23-62084_habUJ0hhAG0P8iIKaRQ74/wZQHyAdFlwHmrejNjOSRG08JeqilJlTM6V8G9UERLJ92/uMDVHIVOPXfN8Zdh8JuYO8oflPnqymIRmff/pU+Gpb871jV2JDA4Cft5gmn+ictKoN4VoSfEZRR+R5hzF2FsoCExDNNw6gLdjtiX94uA_0_____OEM:DM_IoTEnterpriseK
) do (
for /f "tokens=1-9 delims=_" %%A in ("%%#") do (

REM Detect key

if %1==key if %osSKU%==%%C if not defined key (
echo "!allapps! !altapplist!" | find /i "%%A" %nul1% && (
if %%F==1 set notworking=1
set key=%%B
)
)

REM Generate ticket

if %1==ticket if "%key%"=="%%B" (
set "string=OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;Pfn=Microsoft.Windows.%%C.%%D_8wekyb3d8bbwe;PKeyIID=465145217131314304264339481117862266242033457260311819664735280;$([char]0)"
for /f "tokens=* delims=" %%i in ('%psc% [conv%f%ert]::ToBas%f%e64String([Text.En%f%coding]::Uni%f%code.GetBytes("""!string!"""^)^)') do set "encoded=%%i"
echo "!encoded!" | find "AAAA" %nul1% || exit /b

<nul set /p "=<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=!encoded!;TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">%%E=</signature></signatures></genuineProperties></genuineAuthorization>" >"%tdir%\GenuineTicket"
)

)
)
exit /b

::========================================================================================================================================

::  Below code is used to get alternate edition name and key if current edition doesn't support HWID activation

::  1st column = Current SKU ID
::  2nd column = Current Edition Name
::  3rd column = Current Edition Activation ID
::  4th column = Alternate Edition Activation ID
::  5th column = Alternate Edition HWID Key
::  6th column = Alternate Edition Name
::  Separator  = _


:hwidfallback

set notfoundaltactID=
if %_NoEditionChange%==1 exit /b

for %%# in (
125_EnterpriseS-2021_______________cce9d2de-98ee-4ce2-8113-222620c64a27_ed655016-a9e8-4434-95d9-4345352c2552_QPM6N-7J2WJ-P88HH-P3YRH-YY%f%74H_IoTEnterpriseS-2021
125_EnterpriseS-2024_______________f6e29426-a256-4316-88bf-cc5b0f95ec0c_6c4de1b8-24bb-4c17-9a77-7b939414c298_CGK42-GYN6Y-VD22B-BX98W-J8%f%JXD_IoTEnterpriseS-2024
138_ProfessionalSingleLanguage_____a48938aa-62fa-4966-9d44-9f04da3f72f2_4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7JG-NPHTM-C97JM-9MPGT-3V%f%66T_Professional
139_ProfessionalCountrySpecific____f7af7d09-40e4-419c-a49b-eae366689ebd_4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7JG-NPHTM-C97JM-9MPGT-3V%f%66T_Professional
139_ProfessionalCountrySpecific-Zn_01eb852c-424d-4060-94b8-c10d799d7364_4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7JG-NPHTM-C97JM-9MPGT-3V%f%66T_Professional
) do (
for /f "tokens=1-6 delims=_" %%A in ("%%#") do if %osSKU%==%%A (
echo "!allapps! !altapplist!" | find /i "%%C" %nul1% && (
echo "!allapps!" | find /i "%%D" %nul1% && (
set altkey=%%E
set altedition=%%F
) || (
set altedition=%%F
set notfoundaltactID=1
)
)
)
)
exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:OhookActivation

::  To activate Office with Ohook activation, run the script with "/Ohook" parameter or change 0 to 1 in below line
set _act=0

::  To remove Ohook activation, run the script with /Ohook-Uninstall parameter or change 0 to 1 in below line
set _rem=0

::  If value is changed in above lines or parameter is used then script will run in unattended mode

::========================================================================================================================================

cls
color 07
title  Ohook Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/Ohook"                  set _act=1
if /i "%%A"=="/Ohook-Uninstall"        set _rem=1
if /i "%%A"=="-el"                     set _elev=1
)
)

for %%A in (%_act% %_rem%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

if %winbuild% LSS 9200 (
%eline%
echo Unsupported OS version detected [%winbuild%].
echo Ohook Activation is supported only on Windows 8/10/11 and their server equivalents.
echo:
call :dk_color %Blue% "Use Online KMS activation option instead."
goto dk_done
)

::========================================================================================================================================

if %_rem%==1 goto :oh_uninstall

:oh_menu

if %_unattended%==0 (
cls
if not defined terminal mode 76, 25
title  Ohook Activation %masver%
call :oh_checkapps
echo:
echo:
echo:
echo:
if defined checknames (call :dk_color %_Yellow% "                Close [!checknames!] before proceeding...")
echo         ____________________________________________________________
echo:
echo                 [1] Install Ohook Office Activation
echo:
echo                 [2] Uninstall Ohook
echo                 ____________________________________________
echo:
echo                 [3] Download Office
echo:
echo                 [0] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "             " %_Green% "Choose a menu option using your keyboard [1,2,3,0]"
choice /C:1230 /N
set _el=!errorlevel!
if !_el!==4  exit /b
if !_el!==3  start %mas%genuine-installation-media &goto :oh_menu
if !_el!==2  goto :oh_uninstall
if !_el!==1  goto :oh_menu2
goto :oh_menu
)

::========================================================================================================================================

:oh_menu2

cls
if not defined terminal (
mode 130, 32
if exist "%SysPath%\spp\store_test\" mode 134, 32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=32;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)
title  Ohook Activation %masver%

echo:
echo Initializing...
call :dk_chkmal

if not exist %SysPath%\sppsvc.exe (
%eline%
echo [%SysPath%\sppsvc.exe] file is missing, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_reflection
call :dk_ckeckwmic
call :dk_product
call :dk_sppissue

::========================================================================================================================================

set error=

cls
echo:
call :dk_showosinfo

::========================================================================================================================================

echo Initiating Diagnostic Tests...

set "_serv=sppsvc Winmgmt"

::  Software Protection
::  Windows Management Instrumentation

set notwinact=1
set ohookact=1
call :dk_errorcheck

::  Check unsupported office versions

set o14msi=
set o14c2r=
set o16uwp=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o14msi=Office 2010 MSI )
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o14msi=Office 2010 MSI )
%nul% reg query %_68%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 
%nul% reg query %_86%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 

if %winbuild% GEQ 10240 (
for /f "delims=" %%a in ('%psc% "(Get-AppxPackage -name 'Microsoft.Office.Desktop' | Select-Object -ExpandProperty InstallLocation)" %nul6%') do (if exist "%%a\Integration\Integrator.exe" set o16uwp=Office UWP )
)

if not "%o14msi%%o14c2r%%o16uwp%"=="" (
echo:
call :dk_color %Red% "Checking Unsupported Office Install     [ %o14msi%%o14c2r%%o16uwp%]"
if not "%o14msi%%o16uwp%"=="" call :dk_color %Blue% "Use Online KMS option to activate it."
)

if %winbuild% GEQ 10240 %psc% "Get-AppxPackage -name "Microsoft.MicrosoftOfficeHub"" | find /i "Office" %nul1% && (
set ohub=1
)

::========================================================================================================================================

::  Check supported office versions

call :oh_getpath

sc query ClickToRunSvc %nul%
set error1=%errorlevel%

if defined o16c2r if %error1% EQU 1060 (
call :dk_color %Red% "Checking ClickToRun Service             [Not found, Office 16.0 files found]"
set o16c2r=
set error=1
)

sc query OfficeSvc %nul%
set error2=%errorlevel%

if defined o15c2r if %error1% EQU 1060 if %error2% EQU 1060 (
call :dk_color %Red% "Checking ClickToRun Service             [Not found, Office 15.0 files found]"
set o15c2r=
set error=1
)

if "%o16c2r%%o15c2r%%o16msi%%o15msi%"=="" (
set error=1
echo:
if not "%o14msi%%o14c2r%%o16uwp%"=="" (
call :dk_color %Red% "Checking Supported Office Install       [Not Found]"
) else (
call :dk_color %Red% "Checking Installed Office               [Not Found]"
)

if defined ohub (
echo:
echo You only have the Office dashboard app installed, you need to install the full version of Office.
)
echo:
call :dk_color %Blue% "Download and install Office from the below URL and then try again."
echo:
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto dk_done
)

set multioffice=
if not "%o16c2r%%o15c2r%%o16msi%%o15msi%"=="1" set multioffice=1
if not "%o14msi%%o14c2r%%o16uwp%"=="" set multioffice=1

if defined multioffice (
call :dk_color %Gray% "Checking Multiple Office Install        [Found, its recommended to install only one version]"
)

::========================================================================================================================================

::  Check Windows Server

set winserver=
reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v ProductType %nul2% | find /i "WinNT" %nul1% || set winserver=1
if not defined winserver (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Server" %nul1% && set winserver=1
)

::========================================================================================================================================

::  Process Office 15.0 C2R

if not defined o15c2r goto :starto16c2r

call :oh_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=15
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v ProductReleaseIds" %nul6%') do (set "_prids=%o15c2r_reg%\Configuration /v ProductReleaseIds" & set "_config=%o15c2r_reg%\Configuration")
if not defined _oArch   for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v Platform" %nul6%') do (set "_oArch=%%b")
if not defined _version for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v version" %nul6%') do (set "_version=%%b")
if not defined _prids   for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v ProductReleaseId" %nul6%') do (set "_prids=%o15c2r_reg%\propertyBag /v ProductReleaseId" & set "_config=%o15c2r_reg%\propertyBag")

echo "%o15c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=10) || (set _tok=9)
for /f "tokens=%_tok% delims=\" %%a in ('reg query %o15c2r_reg%\ProductReleaseIDs\Active %nul6% ^| findstr /i "Retail Volume"') do (
echo "!_oIds!" | find /i " %%a " %nul1% || (set "_oIds= !_oIds! %%a ")
)

set "_oLPath=%_oRoot%\Licenses"
set "_oIntegrator=%_oRoot%\integration\integrator.exe"

if /i "%_oArch%"=="x64" (set "_hookPath=%_oRoot%\vfs\System"    & set "_hook=sppc64.dll")
if /i "%_oArch%"=="x86" (set "_hookPath=%_oRoot%\vfs\SystemX86" & set "_hook=sppc32.dll")
if not "%osarch%"=="x86" (
if /i "%_oArch%"=="x64" set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if /i "%_oArch%"=="x86" set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

echo:
echo Activating Office...                    [C2R ^| %_version% ^| %_oArch%]

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :starto16c2r
)

call :oh_fixprids
call :oh_process
call :oh_hookinstall

::========================================================================================================================================

:starto16c2r

::  Process Office 16.0 C2R

if not defined o16c2r goto :startmsi

call :oh_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=16
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v AudienceData" %nul6%') do (set "_AudienceData=^| %%b ")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ProductReleaseIds" %nul6%') do (set "_prids=%o16c2r_reg%\Configuration /v ProductReleaseIds" & set "_config=%o16c2r_reg%\Configuration")

echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=9) || (set _tok=8)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs" /s /f ".16" /k %nul6% ^| findstr /i "Retail Volume"') do (
echo "!_oIds!" | find /i " %%a " %nul1% || (set "_oIds= !_oIds! %%a ")
)
set _oIds=%_oIds:.16=%
set _o16c2rIds=%_oIds%

set "_oLPath=%_oRoot%\Licenses16"
set "_oIntegrator=%_oRoot%\integration\integrator.exe"

if /i "%_oArch%"=="x64" (set "_hookPath=%_oRoot%\vfs\System"    & set "_hook=sppc64.dll")
if /i "%_oArch%"=="x86" (set "_hookPath=%_oRoot%\vfs\SystemX86" & set "_hook=sppc32.dll")
if not "%osarch%"=="x86" (
if /i "%_oArch%"=="x64" set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if /i "%_oArch%"=="x86" set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

echo:
echo Activating Office...                    [C2R ^| %_version% %_AudienceData%^| %_oArch%]

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :startmsi
)

call :oh_fixprids
call :oh_process
call :oh_hookinstall

::========================================================================================================================================

::  Old version (16.0.9xxxx and below) of Office with subscription license key may show a banner to sign in to fix license issue.
::  Although script applies a Resiliency registry entry to fix that but it doesn't work on old office versions.
::  Below code checks that condition and informs the user to update the Office.

if defined _sublic (
if not exist "%_oLPath%\Word2019VL_KMS_Client_AE*.xrm-ms" (
call :dk_color %Gray% "Checking Old Office With Sub License    [Found. Update Office, otherwise, it may show a licensing issue-related banner.]"
)
)

::========================================================================================================================================

::  mass grave[.]dev/office-license-is-not-genuine
::  Add registry keys for volume products so that 'non-genuine' banner won't appear 
::  Script already is using MAK instead of GVLK so it won't appear anyway, but registry keys are added incase Office installs default GVLK grace key for volume products

set "kmskey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663"
echo "%_oIds%" | find /i "Volume" %nul1% && (
if %winbuild% GEQ 9200 (
if not "%osarch%"=="x86" (
reg delete "%kmskey%" /f /reg:32 %nul%
reg add "%kmskey%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" /reg:32 %nul%
)
reg delete "%kmskey%" /f %nul%
reg add "%kmskey%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" %nul%
echo Adding a Registry to Prevent Banner     [Successful]
)
)

::========================================================================================================================================

:startmsi

if defined o15msi call :oh_processmsi 15 %o15msi_reg%
if defined o16msi call :oh_processmsi 16 %o16msi_reg%

::========================================================================================================================================

call :oh_clearblock
call :oh_uninstkey
call :oh_licrefresh

::========================================================================================================================================

echo:
if not defined error (
call :dk_color %Green% "Office is permanently activated."
if defined ohub call :dk_color %Gray% "Office apps such as Word, Excel are activated, use them directly. Ignore 'Buy' button in Office dashboard app."
echo Help: %mas%troubleshoot
) else (
call :dk_color %Red% "Some errors were detected."
if not defined ierror if not defined showfix if not defined serv_cor if not defined serv_cste call :dk_color %Blue% "%_fixmsg%"
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

goto :dk_done

::========================================================================================================================================

:oh_uninstall

cls
if not defined terminal mode 99, 32
title  Uninstall Ohook Activation %masver%

set _present=
set _unerror=
call :oh_reset
call :oh_getpath

echo:
echo Uninstalling Ohook activation...
echo:

if defined o16c2r_reg (for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_16CHook=%%b\root\vfs"))
if defined o15c2r_reg (for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg% /v InstallPath" %nul6%') do (set "_15CHook=%%b\root\vfs"))
if defined o16msi_reg (for /f "skip=2 tokens=2*" %%a in ('"reg query %o16msi_reg%\Common\InstallRoot /v Path" %nul6%') do (set "_16MHook=%%b"))
if defined o15msi_reg (for /f "skip=2 tokens=2*" %%a in ('"reg query %o15msi_reg%\Common\InstallRoot /v Path" %nul6%') do (set "_15MHook=%%b"))

if defined _16CHook (if exist "%_16CHook%\System\sppc*dll"    (set _present=1& del /s /f /q "%_16CHook%\System\sppc*dll"    & if exist "%_16CHook%\System\sppc*dll"    set _unerror=1))
if defined _16CHook (if exist "%_16CHook%\SystemX86\sppc*dll" (set _present=1& del /s /f /q "%_16CHook%\SystemX86\sppc*dll" & if exist "%_16CHook%\SystemX86\sppc*dll" set _unerror=1))
if defined _15CHook (if exist "%_15CHook%\System\sppc*dll"    (set _present=1& del /s /f /q "%_15CHook%\System\sppc*dll"    & if exist "%_15CHook%\System\sppc*dll"    set _unerror=1))
if defined _15CHook (if exist "%_15CHook%\SystemX86\sppc*dll" (set _present=1& del /s /f /q "%_15CHook%\SystemX86\sppc*dll" & if exist "%_15CHook%\SystemX86\sppc*dll" set _unerror=1))
if defined _16MHook (if exist "%_16MHook%sppc*dll"            (set _present=1& del /s /f /q "%_16MHook%sppc*dll"            & if exist "%_16MHook%sppc*dll"            set _unerror=1))
if defined _15MHook (if exist "%_15MHook%sppc*dll"            (set _present=1& del /s /f /q "%_15MHook%sppc*dll"            & if exist "%_15MHook%sppc*dll"            set _unerror=1))

for %%# in (15 16) do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" (set _present=1& del /s /f /q "%%~A\Microsoft Office\Office%%#\sppc*dll" & if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" set _unerror=1)
)
)

for %%# in (System SystemX86) do (
for %%G in ("Office 15" "Office") do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" (set _present=1& del /s /f /q "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" & if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" set _unerror=1)
)
)
)

reg query HKCU\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency %nul% && (
echo:
echo Deleting - Registry keys for skipping license check

reg load HKU\DEF_TEMP %SystemDrive%\Users\Default\NTUSER.DAT %nul%
reg query HKU\DEF_TEMP\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency %nul% && reg delete HKU\DEF_TEMP\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /f
reg unload HKU\DEF_TEMP %nul%

set _sidlist=
for /f "tokens=* delims=" %%a in ('%psc% "$p = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'; Get-ChildItem $p | ForEach-Object { $pi = (Get-ItemProperty """"$p\$($_.PSChildName)"""").ProfileImagePath; if ($pi -like '*\Users\*' -and (Test-Path """"$pi\NTUSER.DAT"""") -and -not ($_.PSChildName -match '\.bak$')) { Split-Path $_.PSPath -Leaf } }" %nul6%') do (if defined _sidlist (set _sidlist=!_sidlist! %%a) else (set _sidlist=%%a))

if not defined _sidlist (
for /f "delims=" %%a in ('%psc% "$explorerProc = Get-Process -Name explorer | Where-Object {$_.SessionId -eq (Get-Process -Id $pid).SessionId} | Select-Object -First 1; $sid = (gwmi -Query ('Select * From Win32_Process Where ProcessID=' + $explorerProc.Id)).GetOwnerSid().Sid; $sid" %nul6%') do (set _sidlist=%%a)
)

for %%# in (!_sidlist!) do (

reg query HKU\%%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency %nul% && reg delete HKU\%%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /f

reg query HKU\%%#\Software %nul% || (
for /f "skip=2 tokens=2*" %%a in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\%%#" /v ProfileImagePath" %nul6%') do (
reg load HKU\%%# "%%b\NTUSER.DAT" %nul%
reg query HKU\%%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency %nul% && reg delete HKU\%%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /f
reg unload HKU\%%# %nul%
)
)
)
)

set "kmskey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663"
reg query "%kmskey%" %nul% && (
echo:
echo Deleting - Registry keys for preventing non-genuine banner
reg delete "%kmskey%" /f
)

reg query "%kmskey%" /reg:32 %nul% && (
reg delete "%kmskey%" /f /reg:32
)

echo __________________________________________________________________________________________
echo:

if not defined _present (
echo Ohook activation is not installed.
) else (
if defined _unerror (
call :dk_color %Red% "Failed to uninstall Ohook activation."
call :oh_checkapps
if defined checknames (
call :dk_color %Blue% "Close [!checknames!] and try again."
call :dk_color %Blue% "If it is still not fixed, reboot your machine using the restart option and try again."
) else (
call :dk_color %Blue% "Reboot your machine using the restart option and try again."
)
) else (
call :dk_color %Green% "Successfully uninstalled Ohook activation."
)
)
echo __________________________________________________________________________________________

goto :dk_done

::========================================================================================================================================

:oh_reset

set key=
set _oRoot=
set _oArch=
set _oIds=
set _oLPath=
set _hookPath=
set _hook=
set _sppcPath=
set _actid=
set _prod=
set _lic=
set _arr=
set _prids=
set _config=
set _version=
set _License=
set _oBranding=
exit /b

::========================================================================================================================================

:oh_getpath

set o16c2r=
set o15c2r=
set o16msi=
set o15msi=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set o16c2r=1&set o16c2r_reg=%_86%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set o16c2r=1&set o16c2r_reg=%_68%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set o15c2r=1&set o15c2r_reg=%_86%\15.0\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set o15c2r=1&set o15c2r_reg=%_68%\15.0\ClickToRun)

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o16msi=1&set o16msi_reg=%_86%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o16msi=1&set o16msi_reg=%_68%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o15msi=1&set o15msi_reg=%_86%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o15msi=1&set o15msi_reg=%_68%\15.0)

exit /b

::========================================================================================================================================

::  Some Office Retail to Volume converter tools may edit the ProductReleaseIds to add VL products. This code restores it because it may affect features.

:oh_fixprids

if not defined _prids (
call :dk_color %Gray% "Checking ProductReleaseIds In Registry  [Not Found]"
exit /b
)

set _pridsR=
set _pridsE=
for /f "skip=2 tokens=2*" %%a in ('"reg query %_prids%" %nul6%') do (set "_pridsR=%%b")

set _pridsR=%_pridsR:,= %
for %%# in (%_pridsR%) do (echo %%# | findstr /I "%_oIds%" %nul1% || set _pridsE=1)
for %%# in (%_oIds%) do (echo %%# | findstr /I "%_pridsR%" %nul1% || set _pridsE=1)

if not defined _pridsE exit /b
reg add %_prids% /t REG_SZ /d "" /f %nul1%

for %%# in (%_oIds%) do (
for /f "skip=2 tokens=2*" %%a in ('reg query %_prids%') do if not "%%b"=="" (
reg add %_prids% /t REG_SZ /d "%%b,%%#" /f %nul1%
) else (
reg add %_prids% /t REG_SZ /d "%%#" /f %nul1%
)
)

exit /b

::========================================================================================================================================

:oh_installlic

if not defined _oLPath exit /b

if defined _oIntegrator (
if %oVer%==16 (
"!_oIntegrator!" /I /License PRIDName=%_License%.16 PidKey=%key% %nul%
) else (
"!_oIntegrator!" /I /License PRIDName=%_License% PidKey=%key% %nul%
)
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663
echo "!allapps!" | find /i "!_actid!" %nul1% && exit /b
)

::  Fallback to manual method to install licenses incase integrator.exe is not working

set _License=%_License:XVolume=XC2RVL_%

set _License=%_License:O365EduCloudRetail=O365EduCloudEDUR_%

set _License=%_License:ProjectProRetail=ProjectProO365R_%
set _License=%_License:ProjectStdRetail=ProjectStdO365R_%
set _License=%_License:VisioProRetail=VisioProO365R_%
set _License=%_License:VisioStdRetail=VisioStdO365R_%

if defined _preview set _License=%_License:Volume=PreviewVL_%

set _License=%_License:Retail=R_%
set _License=%_License:Volume=VL_%

for %%# in ("!_oLPath!\client-issuance-*.xrm-ms") do (
if defined _arr (set "_arr=!_arr!;"!_oLPath!\%%~nx#"") else (set "_arr="!_oLPath!\%%~nx#"")
)

for %%# in ("!_oLPath!\%_License%*.xrm-ms") do (
if defined _arr (set "_arr=!_arr!;"!_oLPath!\%%~nx#"") else (set "_arr="!_oLPath!\%%~nx#"")
)

%psc% "$sls = Get-WmiObject %sps%; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); InstallLicenseArr '!_arr!'; InstallLicenseFile '"!_oLPath!\pkeyconfig-office.xrm-ms"'" %nul%

call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663
echo "!allapps!" | find /i "!_actid!" %nul1% || (
set error=1
call :dk_color %Red% "Installing Missing License Files        [Office %oVer%.0 %_prod%] [Failed]"
)

exit /b

::========================================================================================================================================

:oh_hookinstall

set ierror=
set hasherror=

if %_hook%==sppc32.dll set offset=2564
if %_hook%==sppc64.dll set offset=3076

del /s /q "%_hookPath%\sppcs.dll" %nul%
del /s /q "%_hookPath%\sppc.dll" %nul%

if exist "%_hookPath%\sppcs.dll" set "ierror=Remove Previous Ohook Install"
if exist "%_hookPath%\sppc.dll" set "ierror=Remove Previous Ohook Install"

mklink "%_hookPath%\sppcs.dll" "%_sppcPath%" %nul%
if not %errorlevel%==0 (
if not defined ierror set ierror=mklink
)

set exhook=
if exist "!_work!\BIN\%_hook%" set exhook=1

if not exist "%_hookPath%\sppc.dll" (
if defined exhook (
pushd "!_work!\BIN\"
copy /y /b "%_hook%" "%_hookPath%\sppc.dll" %nul%
popd
) else (
call :oh_extractdll "%_hookPath%\sppc.dll" "%offset%"
)
)
if not exist "%_hookPath%\sppc.dll" (if not defined ierror set ierror=Copy)

echo:
if not defined ierror (
echo Symlinking System's sppc.dll to         ["%_hookPath%\sppcs.dll"] [Successful]
if defined exhook (
echo Copying Custom %_hook% to            ["%_hookPath%\sppc.dll"] [Successful]
) else (
echo Extracting Custom %_hook% to         ["%_hookPath%\sppc.dll"] [Successful]
)
) else (
set error=1
call :dk_color %Red% "Installing Ohook                        [Failed to %ierror%]"
echo:
call :oh_checkapps
if defined checknames (
call :dk_color %Blue% "Close [!checknames!] and try again."
call :dk_color %Blue% "If it is still not fixed, reboot your machine using the restart option and try again."
) else (
if /i not "%ierror%"=="Copy" call :dk_color %Blue% "Reboot your machine using the restart option and try again."
if /i "%ierror%"=="Copy" call :dk_color %Blue% "If you are using any third-party antivirus, check if it is blocking the script."
)
echo:
)

if not defined exhook if not defined ierror (
if defined hasherror (
set error=1
set ierror=1
call :dk_color %Red% "Modifying Hash of Custom %_hook%     [Failed]"
) else (
echo Modifying Hash of Custom %_hook%     [Successful]
)
)

exit /b

::========================================================================================================================================

:oh_process

for %%# in (%_oIds%) do (

set key=
set _actid=
set _lic=
set _preview=
set _License=%%#

echo %%# | find /i "2024" %nul% && (
if exist "!_oLPath!\ProPlus2024PreviewVL_*.xrm-ms" if not exist "!_oLPath!\ProPlus2024VL_*.xrm-ms" set _preview=-Preview
)
set _prod=%%#!_preview!

call :ohookdata getinfo !_prod!

if not "!key!"=="" (
echo "!allapps!" | find /i "!_actid!" %nul1% || call :oh_installlic
call :dk_inskey "[!key!] [!_prod!] [!_lic!]"
) else (
set error=1
call :dk_color %Red% "Checking Product In Script              [Office %oVer%.0 !_prod! not found in script]"
call :dk_color %Blue% "Make sure you are using the latest version of MAS."
set fixes=%fixes% %mas%
call :dk_color %_Yellow% "%mas%"
)
)

::  Add SharedComputerLicensing registry key if Retail Office C2R is installed on Windows Server
::  https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/click-to-run-office-on-terminal-server

if defined winserver if defined _config (
echo %_oIds% | find /i "Retail" %nul1% && (
set scaIsNeeded=1
reg add %_config% /v SharedComputerLicensing /t REG_SZ /d "1" /f %nul1%
echo Adding SharedComputerLicensing Reg      [Successful] [Needed on Server With Retail Office]"
)
)

exit /b

::========================================================================================================================================

:oh_processmsi

::  Process Office MSI Version

call :oh_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=%1
for /f "skip=2 tokens=2*" %%a in ('"reg query %2\Common\InstallRoot /v Path" %nul6%') do (set "_oRoot=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %2\Common\ProductVersion /v LastProduct" %nul6%') do (set "_version=%%b")
if "%_oRoot:~-1%"=="\" set "_oRoot=%_oRoot:~0,-1%"

echo "%2" | find /i "Wow6432Node" %nul1% && set _oArch=x86
if not "%osarch%"=="x86" if not defined _oArch set _oArch=x64
if "%osarch%"=="x86" set _oArch=x86

if /i "%_oArch%"=="x64" (set "_hookPath=%_oRoot%" & set "_hook=sppc64.dll")
if /i "%_oArch%"=="x86" (set "_hookPath=%_oRoot%" & set "_hook=sppc32.dll")
if not "%osarch%"=="x86" (
if /i "%_oArch%"=="x64" set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if /i "%_oArch%"=="x86" set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

set "_common=%CommonProgramFiles%"
if defined PROCESSOR_ARCHITEW6432 set "_common=%CommonProgramW6432%"
set "_common2=%CommonProgramFiles(x86)%"

for /r "%_common%\Microsoft Shared\OFFICE%oVer%\" %%f in (BRANDING.XML) do if exist "%%f" set "_oBranding=%%f"
if not defined _oBranding for /r "%_common2%\Microsoft Shared\OFFICE%oVer%\" %%f in (BRANDING.XML) do if exist "%%f" set "_oBranding=%%f"

call :ohookdata getmsiprod %2

echo:
echo Activating Office...                    [MSI ^| %_version% ^| %_oArch%]

if not defined _oBranding (
set error=1
call :dk_color %Red% "Checking BRANDING.XML                   [Not Found, aborting activation...]"
exit /b
)

if not defined _oIds (
set error=1
call :dk_color %Red% "Checking Installed Products             [Product IDs not found, aborting activation...]"
exit /b
)

call :oh_process
call :oh_hookinstall

exit /b

::========================================================================================================================================

:oh_clearblock

::  Find remnants of Office vNext/shared/device license block and remove it because it stops other licenses from appearing
::  https://learn.microsoft.com/office/troubleshoot/activation/reset-office-365-proplus-activation-state

set _sidlist=
for /f "tokens=* delims=" %%a in ('%psc% "$p = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'; Get-ChildItem $p | ForEach-Object { $pi = (Get-ItemProperty """"$p\$($_.PSChildName)"""").ProfileImagePath; if ($pi -like '*\Users\*' -and (Test-Path """"$pi\NTUSER.DAT"""") -and -not ($_.PSChildName -match '\.bak$')) { Split-Path $_.PSPath -Leaf } }" %nul6%') do (if defined _sidlist (set _sidlist=!_sidlist! %%a) else (set _sidlist=%%a))

if not defined _sidlist (
for /f "delims=" %%a in ('%psc% "$explorerProc = Get-Process -Name explorer | Where-Object {$_.SessionId -eq (Get-Process -Id $pid).SessionId} | Select-Object -First 1; $sid = (gwmi -Query ('Select * From Win32_Process Where ProcessID=' + $explorerProc.Id)).GetOwnerSid().Sid; $sid" %nul6%') do (set _sidlist=%%a)
)

::==========================

::  Load the unloaded useraccounts registry

set loadedsids=
set alrloadedsids=

for %%# in (%_sidlist%) do (
reg query HKU\%%#\Software %nul% && (
call set "alrloadedsids=%%alrloadedsids%% %%#"
) || (
for /f "skip=2 tokens=2*" %%a in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\%%#" /v ProfileImagePath" %nul6%') do (
reg load HKU\%%# "%%b\NTUSER.DAT" %nul%
reg query HKU\%%#\Software %nul% && (
call set "loadedsids=%%loadedsids%% %%#"
) || (
reg unload HKU\%%# %nul%
)
)
)
)

::==========================

set "_sidlist=%loadedsids% %alrloadedsids%"

set /a counter=0
for %%# in (%_sidlist%) do set /a counter+=1

if %counter% EQU 0 (
set error=1
call :dk_color %Red% "Checking User Accounts SID              [Not Found]"
exit /b
)

if %counter% GTR 10 (
call :dk_color %Gray% "Checking Total User Accounts            [%counter%]"
)

::==========================

::  Clear the vNext/shared/device license blocks which may prevent ohook activation

rmdir /s /q "%ProgramData%\Microsoft\Office\Licenses\" %nul%

for %%x in (15 16) do (
for %%# in (%_sidlist%) do (
reg delete HKU\%%#\Software\Microsoft\Office\%%x.0\Common\Licensing /f %nul%

for /f "skip=2 tokens=2*" %%a in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\%%#" /v ProfileImagePath" %nul6%') do (
rmdir /s /q "%%b\AppData\Local\Microsoft\Office\Licenses\" %nul%
rmdir /s /q "%%b\AppData\Local\Microsoft\Office\%%x.0\Licensing\" %nul%
)
)
reg delete "HKLM\SOFTWARE\Microsoft\Office\%%x.0\Common\Licensing" /f %nul%
reg delete "HKLM\SOFTWARE\Microsoft\Office\%%x.0\Common\Licensing" /f /reg:32 %nul%
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Office\%%x.0\Common\Licensing" /f %nul%
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Office\%%x.0\Common\Licensing" /f /reg:32 %nul%
)

::  Clear vNext in UWP Office

if defined o16uwpapplist (
for %%# in (%_sidlist%) do (
for /f "skip=2 tokens=2*" %%a in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\%%#" /v ProfileImagePath" %nul6%') do (
rmdir /s /q "%%b\AppData\Local\Packages\Microsoft.Office.Desktop_8wekyb3d8bbwe\LocalCache\Local\Microsoft\Office\Licenses\" %nul%
if exist "%%b\AppData\Local\Packages\Microsoft.Office.Desktop_8wekyb3d8bbwe\SystemAppData\Helium\User.dat" (
set defname=DEFTEMP-%%#
reg load HKU\!defname! "%%b\AppData\Local\Packages\Microsoft.Office.Desktop_8wekyb3d8bbwe\SystemAppData\Helium\User.dat" %nul%
reg delete HKU\!defname!\Software\Microsoft\Office\16.0\Common\Licensing /f %nul%
reg unload HKU\!defname! %nul%
)
)
)
)

::  Clear SharedComputerLicensing for office
::  https://learn.microsoft.com/en-us/deployoffice/overview-shared-computer-activation

if not defined scaIsNeeded (
reg delete HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v SharedComputerLicensing /f %nul%
reg delete HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v SharedComputerLicensing /f /reg:32 %nul%
reg delete HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v SharedComputerLicensing /f %nul%
reg delete HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v SharedComputerLicensing /f /reg:32 %nul%
)

::  Clear device-based-licensing
::  https://learn.microsoft.com/deployoffice/device-based-licensing

for /f %%# in ('reg query "%o16c2r_reg%\Configuration" /f *.DeviceBasedLicensing %nul6% ^| findstr REG_') do reg delete "%o16c2r_reg%\Configuration" /v %%# /f %nul%

::  Remove OEM registry key
::  https://support.microsoft.com/office/office-repeatedly-prompts-you-to-activate-on-a-new-pc-a9a6b05f-f6ce-4d1f-8d49-eb5007b64ba1

for %%# in (15 16) do (
reg delete "HKLM\SOFTWARE\Microsoft\Office\%%#.0\Common\OEM" /f %nul%
reg delete "HKLM\SOFTWARE\Microsoft\Office\%%#.0\Common\OEM" /f /reg:32 %nul%
)

reg delete "HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies\0ff1ce15-a989-479d-af46-f275c6370663" /f %nul%
reg delete "HKU\S-1-5-20\Software\Microsoft\OfficeSoftwareProtectionPlatform\Policies\0ff1ce15-a989-479d-af46-f275c6370663" /f %nul%
reg delete "HKU\S-1-5-20\Software\Microsoft\OfficeSoftwareProtectionPlatform\Policies\59a52881-a989-479d-af46-f275c6370663" /f %nul%

echo Clearing Office License Blocks          [Successfully cleared from all %counter% user accounts]

::==========================

::  Some retail products attempt to validate the license and may show a banner "There was a problem checking this device's license status."
::  Resiliency registry entry can skip this check

set defname=DEFTEMP-%random%
for /f "skip=2 tokens=2*" %%a in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" /v Default" %nul6%') do call set "defdat=%%b"

if defined o16c2r if defined ohookact (
if exist "%defdat%\NTUSER.DAT" (
reg load HKU\%defname% "%defdat%\NTUSER.DAT" %nul%
reg query HKU\%defname%\Software %nul% && (
reg add HKU\%defname%\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /v "TimeOfLastHeartbeatFailure" /t REG_SZ /d "2040-01-01T00:00:00Z" /f %nul%
)
reg unload HKU\%defname% %nul%
)

for %%# in (%_sidlist%) do (
reg delete HKU\%%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /f %nul%
reg add HKU\%%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /v "TimeOfLastHeartbeatFailure" /t REG_SZ /d "2040-01-01T00:00:00Z" /f %nul%
)
echo Adding Registry to Skip License Check   [Successfully added to all %counter% ^& future new user accounts]
)

::==========================

::  Unload the loaded useraccounts registry

for %%# in (%loadedsids%) do (
reg unload HKU\%%# %nul%
)

exit /b

::========================================================================================================================================

::  Uninstall other / grace Keys

:oh_uninstkey

set upk_result=0
call :dk_actid 0ff1ce15-a989-479d-af46-f275c6370663

if "%_actprojvis%"=="1" (
for /f "delims=" %%a in ('%psc% "Get-WmiObject -Query 'SELECT ID, Description, LicenseFamily FROM %spp% WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663'' AND PartialProductKey IS NOT NULL' | Where-Object { $_.LicenseFamily -notmatch 'Project' -and $_.LicenseFamily -notmatch 'Visio' } | Select-Object -ExpandProperty ID" %nul6%') do call set "_allactid=%%a !_allactid!"
for /f "delims=" %%a in ('%psc% "Get-WmiObject -Query 'SELECT ID, Description, LicenseFamily FROM %spp% WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663'' AND PartialProductKey IS NOT NULL' | Where-Object { '!_allactid!' -contains $_.ID -and ($_.LicenseFamily -match 'Project' -or $_.LicenseFamily -match 'Visio') } | Select-Object -ExpandProperty ID" %nul6%') do call set "_allactid=%%a !_allactid!"
)

for %%# in (%apps%) do (
echo "%_allactid%" | find /i "%%#" %nul1% || (

if %_wmic% EQU 1 wmic path %spp% where ID='%%#' call UninstallProductKey %nul%
if %_wmic% EQU 0 %psc% "$null=([WMI]'%spp%=''%%#''').UninstallProductKey()" %nul%

if !errorlevel!==0 (
set upk_result=1
) else (
set error=1
set upk_result=2
)
)
)

if defined ohookact if not %upk_result%==0 echo:
if %upk_result%==1 echo Uninstalling Other/Grace Keys           [Successful]
if %upk_result%==2 call :dk_color %Red% "Uninstalling Other/Grace Keys           [Failed]"
exit /b

::========================================================================================================================================

::  Refresh Windows Insider Preview Licenses
::  It required in Insider versions otherwise office may not activate

:oh_licrefresh

if exist "%SysPath%\spp\store_test\2.0\tokens.dat" (
%psc% "Stop-Service sppsvc -force; $sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
if !errorlevel! NEQ 0 %psc% "$sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
)
exit /b

::========================================================================================================================================

::  Check running office apps and notify user

:oh_checkapps

set checkapps=
set checknames=
for /f "tokens=1" %%i in ('tasklist ^| findstr /I ".exe" %nul6%') do (set "checkapps=!checkapps! -%%i-")

for %%# in (
Access_msaccess.exe
Excel_excel.exe
Groove_groove.exe
Lync_lync.exe
OneNote_onenote.exe
Outlook_outlook.exe
PowerPoint_powerpnt.exe
Project_winproj.exe
Publisher_mspub.exe
Visio_visio.exe
Word_winword.exe
Lime_lime.exe
) do (
for /f "tokens=1-2 delims=_" %%A in ("%%#") do (
echo !checkapps! | find /i "-%%B-" %nul1% && (if defined checknames (set "checknames=!checknames! %%A") else (set "checknames=%%A"))
)
)
exit /b

::========================================================================================================================================

::  1st column = Office version number
::  2nd column = Activation ID
::  3rd column = Generic key. Preference is given in this order, Retail:TB:Sub > Retail > OEM:NONSLP > Volume:MAK > Volume:GVLK
::  4th column = Last part of license description
::  5th column = Edition
::  Separator  = "_"

:ohookdata

set f=
for %%# in (
:: Office 2013
15_ab4d047b-97cf-4126-a69f-34df08e2f254_B7RFY-7NXPK-Q4342-Y9X2H-3J%f%X4X_Retail________AccessRetail
15_259de5be-492b-44b3-9d78-9645f848f7b0_X3XNB-HJB7K-66THH-8DWQ3-XH%f%GJP_Bypass________AccessRuntimeRetail
15_4374022d-56b8-48c1-9bb7-d8f2fc726343_9MF9G-CN32B-HV7XT-9XJ8T-9K%f%VF4_MAK___________AccessVolume
15_1b1d9bd5-12ea-4063-964c-16e7e87d6e08_NT889-MBH4X-8MD4H-X8R2D-WQ%f%HF8_Retail________ExcelRetail
15_ac1ae7fd-b949-4e04-a330-849bc40638cf_Y3N36-YCHDK-XYWBG-KYQVV-BD%f%TJ2_MAK___________ExcelVolume
15_cfaf5356-49e3-48a8-ab3c-e729ab791250_BMK4W-6N88B-BP9QR-PHFCK-MG%f%7GF_Retail________GrooveRetail
15_4825ac28-ce41-45a7-9e6e-1fed74057601_RN84D-7HCWY-FTCBK-JMXWM-HT%f%7GJ_MAK___________GrooveVolume
15_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_2WQNF-GBK4B-XVG6F-BBMX7-M4%f%F2Y_OEM-Perp______HomeBusinessPipcRetail
15_a2b90e7a-a797-4713-af90-f0becf52a1dd_YWD4R-CNKVT-VG8VJ-9333B-RC%f%W9F_Subscription__HomeBusinessRetail
15_1fdfb4e4-f9c9-41c4-b055-c80daf00697d_B92QY-NKYFQ-6KTKH-VWW2Q-3P%f%B3B_OEM-ARM_______HomeStudentARMRetail
15_ebef9f05-5273-404a-9253-c5e252f50555_QPG96-CNT7M-KH36K-KY4HQ-M7%f%TBR_OEM-ARM_______HomeStudentPlusARMRetail
15_f2de350d-3028-410a-bfae-283e00b44d0e_6WW3N-BDGM9-PCCHD-9QPP9-P3%f%4QG_Subscription__HomeStudentRetail
15_44984381-406e-4a35-b1c3-e54f499556e2_RV7NQ-HY3WW-7CKWH-QTVMW-29%f%VHC_Retail________InfoPathRetail
15_9e016989-4007-42a6-8051-64eb97110cf2_C4TGN-QQW6Y-FYKXC-6WJW7-X7%f%3VG_MAK___________InfoPathVolume
15_9103f3ce-1084-447a-827e-d6097f68c895_6MDN4-WF3FV-4WH3Q-W699V-RG%f%CMY_PrepidBypass__LyncAcademicRetail
15_ff693bf4-0276-4ddb-bb42-74ef1a0c9f4d_N42BF-CBY9F-W2C7R-X397X-DY%f%FQW_PrepidBypass__LyncEntryRetail
15_fada6658-bfc6-4c4e-825a-59a89822cda8_89P23-2NK2R-JXM2M-3Q8R8-BW%f%M3Y_Retail________LyncRetail
15_e1264e10-afaf-4439-a98b-256df8bb156f_3WKCD-RN489-4M7XJ-GJ2GQ-YB%f%FQ6_MAK___________LyncVolume
15_69ec9152-153b-471a-bf35-77ec88683eae_VNWHF-FKFBW-Q2RGD-HYHWF-R3%f%HH2_Subscription__MondoRetail
15_f33485a0-310b-4b72-9a0e-b1d605510dbd_2YNYQ-FQMVG-CB8KW-6XKYD-M7%f%RRJ_MAK___________MondoVolume
15_3391e125-f6e4-4b1e-899c-a25e6092d40d_4TGWV-6N9P6-G2H8Y-2HWKB-B4%f%FF4_Bypass________OneNoteFreeRetail
15_8b524bcc-67ea-4876-a509-45e46f6347e8_3KXXQ-PVN2C-8P7YY-HCV88-GV%f%GQ6_Retail________OneNoteRetail
15_b067e965-7521-455b-b9f7-c740204578a2_JDMWF-NJC7B-HRCHY-WFT8G-BP%f%XD9_MAK___________OneNoteVolume
15_12004b48-e6c8-4ffa-ad5a-ac8d4467765a_9N4RQ-CF8R2-HBVCB-J3C9V-94%f%P4D_Retail________OutlookRetail
15_8d577c50-ae5e-47fd-a240-24986f73d503_HNG29-GGWRG-RFC8C-JTFP4-2J%f%9FH_MAK___________OutlookVolume
15_5aab8561-1686-43f7-9ff5-2c861da58d17_9CYB3-NFMRW-YFDG6-XC7TF-BY%f%36J_OEM-Perp______PersonalPipcRetail
15_17e9df2d-ed91-4382-904b-4fed6a12caf0_2NCQJ-MFRMH-TXV83-J7V4C-RV%f%RWC_Retail________PersonalRetail
15_31743b82-bfbc-44b6-aa12-85d42e644d5b_HVMN2-KPHQH-DVQMK-7B3CM-FG%f%BFC_Retail________PowerPointRetail
15_e40dcb44-1d5c-4085-8e8f-943f33c4f004_47DKN-HPJP7-RF9M3-VCYT2-TM%f%Q4G_MAK___________PowerPointVolume
15_064383fa-1538-491c-859b-0ecab169a0ab_N3QMM-GKDT3-JQGX6-7X3MQ-4G%f%BG3_Retail________ProPlusRetail
15_2b88c4f2-ea8f-43cd-805e-4d41346e18a7_QKHNX-M9GGH-T3QMW-YPK4Q-QR%f%P9V_MAK___________ProPlusVolume
15_4e26cac1-e15a-4467-9069-cb47b67fe191_CF9DD-6CNW2-BJWJQ-CVCFX-Y7%f%TXD_OEM-Perp______ProfessionalPipcRetail
15_44bc70e2-fb83-4b09-9082-e5557e0c2ede_MBQBN-CQPT6-PXRMC-TYJFR-3C%f%8MY_Retail________ProfessionalRetail
15_2f72340c-b555-418d-8b46-355944fe66b8_WPY8N-PDPY4-FC7TF-KMP7P-KW%f%YFY_Subscription__ProjectProRetail
15_ed34dc89-1c27-4ecd-8b2f-63d0f4cedc32_WFCT2-NBFQ7-JD7VV-MFJX6-6F%f%2CM_MAK___________ProjectProVolume
15_58d95b09-6af6-453d-a976-8ef0ae0316b1_NTHQT-VKK6W-BRB87-HV346-Y9%f%6W8_Subscription__ProjectStdRetail
15_2b9e4a37-6230-4b42-bee2-e25ce86c8c7a_3CNQX-T34TY-99RH4-C4YD2-KW%f%YGV_MAK___________ProjectStdVolume
15_c3a0814a-70a4-471f-af37-2313a6331111_TWNCJ-YR84W-X7PPF-6DPRP-D6%f%7VC_Retail________PublisherRetail
15_38ea49f6-ad1d-43f1-9888-99a35d7c9409_DJPHV-NCJV6-GWPT6-K26JX-C7%f%GX6_MAK___________PublisherVolume
15_ba3e3833-6a7e-445a-89d0-7802a9a68588_3NY6J-WHT3F-47BDV-JHF36-23%f%43W_PrepidBypass__SPDRetail
15_32255c0a-16b4-4ce2-b388-8a4267e219eb_V6VWN-KC2HR-YYDD6-9V7HQ-7T%f%7VP_Retail________StandardRetail
15_a24cca51-3d54-4c41-8a76-4031f5338cb2_9TN6B-PCYH4-MCVDQ-KT83C-TM%f%Q7T_MAK___________StandardVolume
15_a56a3b37-3a35-4bbb-a036-eee5f1898eee_NVK2G-2MY4G-7JX2P-7D6F2-VF%f%QBR_Subscription__VisioProRetail
15_3e4294dd-a765-49bc-8dbd-cf8b62a4bd3d_YN7CF-XRH6R-CGKRY-GKPV3-BG%f%7WF_MAK___________VisioProVolume
15_980f9e3e-f5a8-41c8-8596-61404addf677_NCRB7-VP48F-43FYY-62P3R-36%f%7WK_Subscription__VisioStdRetail
15_44a1f6ff-0876-4edb-9169-dbb43101ee89_RX63Y-4NFK2-XTYC8-C6B3W-YP%f%XPJ_MAK___________VisioStdVolume
15_191509f2-6977-456f-ab30-cf0492b1e93a_NB77V-RPFQ6-PMMKQ-T87DV-M4%f%D84_Retail________WordRetail
15_9cedef15-be37-4ff0-a08a-13a045540641_RPHPB-Y7NC4-3VYFM-DW7VD-G8%f%YJ8_MAK___________WordVolume
:: Office 365 - 15.0 version
15_742178ed-6b28-42dd-b3d7-b7c0ea78741b_Y9NF9-M2QWD-FF6RJ-QJW36-RR%f%F2T_SubTest_______O365BusinessRetail
15_a96f8dae-da54-4fad-bdc6-108da592707a_3NMDC-G7C3W-68RGP-CB4MH-4C%f%XCH_SubTest1______O365HomePremRetail
15_e3dacc06-3bc2-4e13-8e59-8e05f3232325_H8DN8-Y2YP3-CR9JT-DHDR9-C7%f%GP3_Subscription2_O365ProPlusRetail
15_0bc1dae4-6158-4a1c-a893-807665b934b2_2QCNB-RMDKJ-GC8PB-7QGQV-7Q%f%TQJ_Subscription2_O365SmallBusPremRetail
:: Office 365 - 16.0 version
16_742178ed-6b28-42dd-b3d7-b7c0ea78741b_Y9NF9-M2QWD-FF6RJ-QJW36-RR%f%F2T_SubTest_______O365BusinessRetail
16_2f5c71b4-5b7a-4005-bb68-f9fac26f2ea3_W62NQ-267QR-RTF74-PF2MH-JQ%f%MTH_Subscription__O365EduCloudRetail
16_a96f8dae-da54-4fad-bdc6-108da592707a_3NMDC-G7C3W-68RGP-CB4MH-4C%f%XCH_SubTest1______O365HomePremRetail
16_e3dacc06-3bc2-4e13-8e59-8e05f3232325_H8DN8-Y2YP3-CR9JT-DHDR9-C7%f%GP3_Subscription2_O365ProPlusRetail
16_0bc1dae4-6158-4a1c-a893-807665b934b2_2QCNB-RMDKJ-GC8PB-7QGQV-7Q%f%TQJ_Subscription2_O365SmallBusPremRetail
:: Office 2016
16_bfa358b0-98f1-4125-842e-585fa13032e6_WHK4N-YQGHB-XWXCC-G3HYC-6J%f%F94_Retail________AccessRetail
16_9d9faf9e-d345-4b49-afce-68cb0a539c7c_RNB7V-P48F4-3FYY6-2P3R3-63%f%BQV_PrepidBypass__AccessRuntimeRetail
16_3b2fa33f-cd5a-43a5-bd95-f49f3f546b0b_JJ2Y4-N8KM3-Y8KY3-Y22FR-R3%f%KVK_MAK___________AccessVolume
16_424d52ff-7ad2-4bc7-8ac6-748d767b455d_RKJBN-VWTM2-BDKXX-RKQFD-JT%f%YQ2_Retail________ExcelRetail
16_685062a7-6024-42e7-8c5f-6bb9e63e697f_FVGNR-X82B2-6PRJM-YT4W7-8H%f%V36_MAK___________ExcelVolume
16_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_2WQNF-GBK4B-XVG6F-BBMX7-M4%f%F2Y_OEM-Perp______HomeBusinessPipcRetail
16_86834d00-7896-4a38-8fae-32f20b86fa2b_HM6FM-NVF78-KV9PM-F36B8-D9%f%MXD_Retail________HomeBusinessRetail
16_090896a0-ea98-48ac-b545-ba5da0eb0c9c_PBQPJ-NC22K-69MXD-KWMRF-WF%f%G77_OEM-ARM_______HomeStudentARMRetail
16_6bbe2077-01a4-4269-bf15-5bf4d8efc0b2_6F2NY-7RTX4-MD9KM-TJ43H-94%f%TBT_OEM-ARM_______HomeStudentPlusARMRetail
16_c28acdb8-d8b3-4199-baa4-024d09e97c99_PNPRV-F2627-Q8JVC-3DGR9-WT%f%YRK_Retail________HomeStudentRetail
16_e2127526-b60c-43e0-bed1-3c9dc3d5a468_YWD4R-CNKVT-VG8VJ-9333B-RC%f%3B8_Retail________HomeStudentVNextRetail
16_69ec9152-153b-471a-bf35-77ec88683eae_VNWHF-FKFBW-Q2RGD-HYHWF-R3%f%HH2_Subscription__MondoRetail
16_2cd0ea7e-749f-4288-a05e-567c573b2a6c_FMTQQ-84NR8-2744R-MXF4P-PG%f%YR3_MAK___________MondoVolume
16_436366de-5579-4f24-96db-3893e4400030_XYNTG-R96FY-369HX-YFPHY-F9%f%CPM_Bypass________OneNoteFreeRetail
16_83ac4dd9-1b93-40ed-aa55-ede25bb6af38_FXF6F-CNC26-W643C-K6KB7-6X%f%XW3_Retail________OneNoteRetail
16_23b672da-a456-4860-a8f3-e062a501d7e8_9TYVN-D76HK-BVMWT-Y7G88-9T%f%PPV_MAK___________OneNoteVolume
16_5a670809-0983-4c2d-8aad-d3c2c5b7d5d1_7N4KG-P2QDH-86V9C-DJFVF-36%f%9W9_Retail________OutlookRetail
16_50059979-ac6f-4458-9e79-710bcb41721a_7QPNR-3HFDG-YP6T9-JQCKQ-KK%f%XXC_MAK___________OutlookVolume
16_5aab8561-1686-43f7-9ff5-2c861da58d17_9CYB3-NFMRW-YFDG6-XC7TF-BY%f%36J_OEM-Perp______PersonalPipcRetail
16_a9f645a1-0d6a-4978-926a-abcb363b72a6_FT7VF-XBN92-HPDJV-RHMBY-6V%f%KBF_Retail________PersonalRetail
16_f32d1284-0792-49da-9ac6-deb2bc9c80b6_N7GCB-WQT7K-QRHWG-TTPYD-7T%f%9XF_Retail________PowerPointRetail
16_9b4060c9-a7f5-4a66-b732-faf248b7240f_X3RT9-NDG64-VMK2M-KQ6XY-DP%f%FGV_MAK___________PowerPointVolume
16_de52bd50-9564-4adc-8fcb-a345c17f84f9_GM43N-F742Q-6JDDK-M622J-J8%f%GDV_Retail________ProPlusRetail
16_c47456e3-265d-47b6-8ca0-c30abbd0ca36_FNVK8-8DVCJ-F7X3J-KGVQB-RC%f%2QY_MAK___________ProPlusVolume
16_4e26cac1-e15a-4467-9069-cb47b67fe191_CF9DD-6CNW2-BJWJQ-CVCFX-Y7%f%TXD_OEM-Perp______ProfessionalPipcRetail
16_d64edc00-7453-4301-8428-197343fafb16_NXFTK-YD9Y7-X9MMJ-9BWM6-J2%f%QVH_Retail________ProfessionalRetail
16_2f72340c-b555-418d-8b46-355944fe66b8_WPY8N-PDPY4-FC7TF-KMP7P-KW%f%YFY_Subscription__ProjectProRetail
16_82f502b5-b0b0-4349-bd2c-c560df85b248_PKC3N-8F99H-28MVY-J4RYY-CW%f%GDH_MAK___________ProjectProVolume
16_16728639-a9ab-4994-b6d8-f81051e69833_JBNPH-YF2F7-Q9Y29-86CTG-C9%f%YGV_MAKC2R________ProjectProXVolume
16_58d95b09-6af6-453d-a976-8ef0ae0316b1_NTHQT-VKK6W-BRB87-HV346-Y9%f%6W8_Subscription__ProjectStdRetail
16_82e6b314-2a62-4e51-9220-61358dd230e6_4TGWV-6N9P6-G2H8Y-2HWKB-B4%f%G93_MAK___________ProjectStdVolume
16_431058f0-c059-44c5-b9e7-ed2dd46b6789_N3W2Q-69MBT-27RD9-BH8V3-JT%f%2C8_MAKC2R________ProjectStdXVolume
16_6e0c1d99-c72e-4968-bcb7-ab79e03e201e_WKWND-X6G9G-CDMTV-CPGYJ-6M%f%VBF_Retail________PublisherRetail
16_fcc1757b-5d5f-486a-87cf-c4d6dedb6032_9QVN2-PXXRX-8V4W8-Q7926-TJ%f%GD8_MAK___________PublisherVolume
16_9103f3ce-1084-447a-827e-d6097f68c895_6MDN4-WF3FV-4WH3Q-W699V-RG%f%CMY_PrepidBypass__SkypeServiceBypassRetail
16_971cd368-f2e1-49c1-aedd-330909ce18b6_4N4D8-3J7Y3-YYW7C-73HD2-V8%f%RHY_PrepidBypass__SkypeforBusinessEntryRetail
16_418d2b9f-b491-4d7f-84f1-49e27cc66597_PBJ79-77NY4-VRGFG-Y8WYC-CK%f%CRC_Retail________SkypeforBusinessRetail
16_03ca3b9a-0869-4749-8988-3cbc9d9f51bb_DMTCJ-KNRKR-JV8TQ-V2CR2-VF%f%TFH_MAK___________SkypeforBusinessVolume
16_4a31c291-3a12-4c64-b8ab-cd79212be45e_2FPWN-4H6CM-KD8QQ-8HCHC-P9%f%XYW_Retail________StandardRetail
16_0ed94aac-2234-4309-ba29-74bdbb887083_WHGMQ-JNMGT-MDQVF-WDR69-KQ%f%BWC_MAK___________StandardVolume
16_a56a3b37-3a35-4bbb-a036-eee5f1898eee_NVK2G-2MY4G-7JX2P-7D6F2-VF%f%QBR_Subscription__VisioProRetail
16_295b2c03-4b1c-4221-b292-1411f468bd02_NRKT9-C8GP2-XDYXQ-YW72K-MG%f%92B_MAK___________VisioProVolume
16_0594dc12-8444-4912-936a-747ca742dbdb_G98Q2-B6N77-CFH9J-K824G-XQ%f%CC4_MAKC2R________VisioProXVolume
16_980f9e3e-f5a8-41c8-8596-61404addf677_NCRB7-VP48F-43FYY-62P3R-36%f%7WK_Subscription__VisioStdRetail
16_44151c2d-c398-471f-946f-7660542e3369_XNCJB-YY883-JRW64-DPXMX-JX%f%CR6_MAK___________VisioStdVolume
16_1d1c6879-39a3-47a5-9a6d-aceefa6a289d_B2HTN-JPH8C-J6Y6V-HCHKB-43%f%MGT_MAKC2R________VisioStdXVolume
16_cacaa1bf-da53-4c3b-9700-11738ef1c2a5_P8K82-NQ7GG-JKY8T-6VHVY-88%f%GGD_Retail________WordRetail
16_c3000759-551f-4f4a-bcac-a4b42cbf1de2_YHMWC-YN6V9-WJPXD-3WQKP-TM%f%VCV_MAK___________WordVolume
:: Office 2019
16_518687bd-dc55-45b9-8fa6-f918e1082e83_WRYJ6-G3NP7-7VH94-8X7KP-JB%f%7HC_Retail________Access2019Retail
16_385b91d6-9c2c-4a2e-86b5-f44d44a48c5f_6FWHX-NKYXK-BW34Q-7XC9F-Q9%f%PX7_MAK-AE________Access2019Volume
16_22e6b96c-1011-4cd5-8b35-3c8fb6366b86_FGQNJ-JWJCG-7Q8MG-RMRGJ-9T%f%QVF_PrepidBypass__AccessRuntime2019Retail
16_c201c2b7-02a1-41a8-b496-37c72910cd4a_KBPNW-64CMM-8KWCB-23F44-8B%f%7HM_Retail________Excel2019Retail
16_05cb4e1d-cc81-45d5-a769-f34b09b9b391_8NT4X-GQMCK-62X4P-TW6QP-YK%f%PYF_MAK-AE________Excel2019Volume
16_7fe09eef-5eed-4733-9a60-d7019df11cac_QBN2Y-9B284-9KW78-K48PB-R6%f%2YT_Retail________HomeBusiness2019Retail
16_6303d14a-afad-431f-8434-81052a65f575_DJTNY-4HDWM-TDWB2-8PWC2-W2%f%RRT_OEM-ARM_______HomeStudentARM2019Retail
16_215c841d-ffc1-4f03-bd11-5b27b6ab64cc_NM8WT-CFHB2-QBGXK-J8W6J-GV%f%K8F_OEM-ARM_______HomeStudentPlusARM2019Retail
16_4539aa2c-5c31-4d47-9139-543a868e5741_XNWPM-32XQC-Y7QJC-QGGBV-YY%f%7JK_Retail________HomeStudent2019Retail
16_20e359d5-927f-47c0-8a27-38adbdd27124_WR43D-NMWQQ-HCQR2-VKXDR-37%f%B7H_Retail________Outlook2019Retail
16_92a99ed8-2923-4cb7-a4c5-31da6b0b8cf3_RN3QB-GT6D7-YB3VH-F3RPB-3G%f%QYB_MAK-AE________Outlook2019Volume
16_2747b731-0f1f-413e-a92d-386ec1277dd8_NMBY8-V3CV7-BX6K6-2922Y-43%f%M7T_Retail________Personal2019Retail
16_7e63cc20-ba37-42a1-822d-d5f29f33a108_HN27K-JHJ8R-7T7KK-WJYC3-FM%f%7MM_Retail________PowerPoint2019Retail
16_13c2d7bf-f10d-42eb-9e93-abf846785434_29GNM-VM33V-WR23K-HG2DT-KT%f%QYR_MAK-AE________PowerPoint2019Volume
16_a3072b8f-adcc-4e75-8d62-fdeb9bdfae57_BN4XJ-R9DYY-96W48-YK8DM-MY%f%7PY_Retail________ProPlus2019Retail
16_6755c7a7-4dfe-46f5-bce8-427be8e9dc62_T8YBN-4YV3X-KK24Q-QXBD7-T3%f%C63_MAK-AE________ProPlus2019Volume
16_1717c1e0-47d3-4899-a6d3-1022db7415e0_9NXDK-MRY98-2VJV8-GF73J-TQ%f%9FK_Retail________Professional2019Retail
16_0d270ef7-5aaf-4370-a372-bc806b96adb7_JDTNC-PP77T-T9H2W-G4J2J-VH%f%8JK_Retail________ProjectPro2019Retail
16_d4ebadd6-401b-40d5-adf4-a5d4accd72d1_TBXBD-FNWKJ-WRHBD-KBPHH-XD%f%9F2_MAK-AE________ProjectPro2019Volume
16_bb7ffe5f-daf9-4b79-b107-453e1c8427b5_R3JNT-8PBDP-MTWCK-VD2V8-HM%f%KF9_Retail________ProjectStd2019Retail
16_fdaa3c03-dc27-4a8d-8cbf-c3d843a28ddc_RBRFX-MQNDJ-4XFHF-7QVDR-JH%f%XGC_MAK-AE________ProjectStd2019Volume
16_f053a7c7-f342-4ab8-9526-a1d6e5105823_4QC36-NW3YH-D2Y9D-RJPC7-VV%f%B9D_Retail________Publisher2019Retail
16_40055495-be00-444e-99cc-07446729b53e_K8F2D-NBM32-BF26V-YCKFJ-29%f%Y9W_MAK-AE________Publisher2019Volume
16_b639e55c-8f3e-47fe-9761-26c6a786ad6b_JBDKF-6NCD6-49K3G-2TV79-BK%f%P73_Retail________SkypeforBusiness2019Retail
16_15a430d4-5e3f-4e6d-8a0a-14bf3caee4c7_9MNQ7-YPQ3B-6WJXM-G83T3-CB%f%BDK_MAK-AE________SkypeforBusiness2019Volume
16_f88cfdec-94ce-4463-a969-037be92bc0e7_N9722-BV9H6-WTJTT-FPB93-97%f%8MK_PrepidBypass__SkypeforBusinessEntry2019Retail
16_fdfa34dd-a472-4b85-bee6-cf07bf0aaa1c_NDGVM-MD27H-2XHVC-KDDX2-YK%f%P74_Retail________Standard2019Retail
16_beb5065c-1872-409e-94e2-403bcfb6a878_NT3V6-XMBK7-Q66MF-VMKR4-FC%f%33M_MAK-AE________Standard2019Volume
16_a6f69d68-5590-4e02-80b9-e7233dff204e_2NWVW-QGF4T-9CPMB-WYDQ9-7X%f%P79_Retail________VisioPro2019Retail
16_f41abf81-f409-4b0d-889d-92b3e3d7d005_33YF4-GNCQ3-J6GDM-J67P3-FM%f%7QP_MAK-AE________VisioPro2019Volume
16_4a582021-18c2-489f-9b3d-5186de48f1cd_263WK-3N797-7R437-28BKG-3V%f%8M8_Retail________VisioStd2019Retail
16_933ed0e3-747d-48b0-9c2c-7ceb4c7e473d_BGNHX-QTPRJ-F9C9G-R8QQG-8T%f%27F_MAK-AE________VisioStd2019Volume
16_72cee1c2-3376-4377-9f25-4024b6baadf8_JXR8H-NJ3MK-X66W8-78CWD-QR%f%VR2_Retail________Word2019Retail
16_fe5fe9d5-3b06-4015-aa35-b146f85c4709_9F36R-PNVHH-3DXGQ-7CD2H-R9%f%D3V_MAK-AE________Word2019Volume
:: Office 2021
16_f634398e-af69-48c9-b256-477bea3078b5_P286B-N3XYP-36QRQ-29CMP-RV%f%X9M_Retail________Access2021Retail
16_ae17db74-16b0-430b-912f-4fe456e271db_JBH3N-P97FP-FRTJD-MGK2C-VF%f%WG6_MAK-AE________Access2021Volume
16_844c36cb-851c-49e7-9079-12e62a049e2a_MNX9D-PB834-VCGY2-K2RW2-2D%f%P3D_Bypass________AccessRuntime2021Retail
16_fb099c19-d48b-4a2f-a160-4383011060aa_V6QFB-7N7G9-PF7W9-M8FQM-MY%f%8G9_Retail________Excel2021Retail
16_9da1ecdb-3a62-4273-a234-bf6d43dc0778_WNYR4-KMR9H-KVC8W-7HJ8B-K7%f%9DQ_MAK-AE________Excel2021Volume
16_38b92b63-1dff-4be7-8483-2a839441a2bc_JM99N-4MMD8-DQCGJ-VMYFY-R6%f%3YK_Subscription__HomeBusiness2021Retail
16_2f258377-738f-48dd-9397-287e43079958_N3CWD-38XVH-KRX2Y-YRP74-6R%f%BB2_Subscription__HomeStudent2021Retail
16_279706f4-3a4b-4877-949b-f8c299cf0cc5_NB2TQ-3Y79C-77C6M-QMY7H-7Q%f%Y8P_Retail________OneNote2021Retail
16_0c7af60d-0664-49fc-9b01-41b2dea81380_THNKC-KFR6C-Y86Q9-W8CB3-GF%f%7PD_MAK-AE________OneNote2021Volume
16_778ccb9a-2f6a-44e5-853c-eb22b7609643_CNM3W-V94GB-QJQHH-BDQ3J-33%f%Y8H_Bypass________OneNoteFree2021Retail
16_ecea2cfa-d406-4a7f-be0d-c6163250d126_4NCWR-9V92Y-34VB2-RPTHR-YT%f%GR7_Retail________Outlook2021Retail
16_45bf67f9-0fc8-4335-8b09-9226cef8a576_JQ9MJ-QYN6B-67PX9-GYFVY-QJ%f%6TB_MAK-AE________Outlook2021Volume
16_8f89391e-eedb-429d-af90-9d36fbf94de6_RRRYB-DN749-GCPW4-9H6VK-HC%f%HPT_Retail________Personal2021Retail
16_c9bf5e86-f5e3-4ac6-8d52-e114a604d7bf_3KXXQ-PVN2C-8P7YY-HCV88-GV%f%M96_Retail1_______PowerPoint2021Retail
16_716f2434-41b6-4969-ab73-e61e593a3875_39G2N-3BD9C-C4XCM-BD4QG-FV%f%YDY_MAK-AE________PowerPoint2021Volume
16_c2f04adf-a5de-45c5-99a5-f5fddbda74a8_8WXTP-MN628-KY44G-VJWCK-C7%f%PCF_Retail________ProPlus2021Retail
16_3f180b30-9b05-4fe2-aa8d-0c1c4790f811_RNHJY-DTFXW-HW9F8-4982D-MD%f%2CW_MAK-AE1_______ProPlus2021Volume
16_96097a68-b5c5-4b19-8600-2e8d6841a0db_JRJNJ-33M7C-R73X3-P9XF7-R9%f%F6M_MAK-AE________ProPlusSPLA2021Volume
16_711e48a6-1a79-4b00-af10-73f4ca3aaac4_DJPHV-NCJV6-GWPT6-K26JX-C7%f%PBG_Retail________Professional2021Retail
16_3747d1d5-55a8-4bc3-b53d-19fff1913195_QKHNX-M9GGH-T3QMW-YPK4Q-QR%f%WMV_Retail________ProjectPro2021Retail
16_17739068-86c4-4924-8633-1e529abc7efc_HVC34-CVNPG-RVCMT-X2JRF-CR%f%7RK_MAK-AE1_______ProjectPro2021Volume
16_4ea64dca-227c-436b-813f-b6624be2d54c_2B96V-X9NJY-WFBRC-Q8MP2-7C%f%HRR_Retail________ProjectStd2021Retail
16_84313d1e-47c8-4e27-8ced-0476b7ee46c4_3CNQX-T34TY-99RH4-C4YD2-KW%f%6WH_MAK-AE________ProjectStd2021Volume
16_b769b746-53b1-4d89-8a68-41944dafe797_CDNFG-77T8D-VKQJX-B7KT3-KK%f%28V_Retail1_______Publisher2021Retail
16_a0234cfe-99bd-4586-a812-4f296323c760_2KXJH-3NHTW-RDBPX-QFRXJ-MT%f%GXF_MAK-AE________Publisher2021Volume
16_c3fb48b2-1fd4-4dc8-af39-819edf194288_DVBXN-HFT43-CVPRQ-J89TF-VM%f%MHG_Retail________SkypeforBusiness2021Retail
16_6029109c-ceb8-4ee5-b324-f8eb2981e99a_R3FCY-NHGC7-CBPVP-8Q934-YT%f%GXG_MAK-AE________SkypeforBusiness2021Volume
16_9e7e7b8e-a0e7-467b-9749-d0de82fb7297_HXNXB-J4JGM-TCF44-2X2CV-FJ%f%VVH_Retail________Standard2021Retail
16_223a60d8-9002-4a55-abac-593f5b66ca45_2CJN4-C9XK2-HFPQ6-YH498-82%f%TXH_MAK-AE________Standard2021Volume
16_b99ba8c4-e257-4b70-a31a-8bd308ce7073_BQWDW-NJ9YF-P7Y79-H6DCT-MK%f%Q9C_MAK-AE________StandardSPLA2021Volume
16_814014d3-c30b-4f63-a493-3708e0dc0ba8_T6P26-NJVBR-76BK8-WBCDY-TX%f%3BC_Retail________VisioPro2021Retail
16_c590605a-a08a-4cc7-8dc2-f1ffb3d06949_JNKBX-MH9P4-K8YYV-8CG2Y-VQ%f%2C8_MAK-AE________VisioPro2021Volume
16_16d43989-a5ef-47e2-9ff1-272784caee24_89NYY-KB93R-7X22F-93QDF-DJ%f%6YM_Retail________VisioStd2021Retail
16_d55f90ee-4ba2-4d02-b216-1300ee50e2af_BW43B-4PNFP-V637F-23TR2-J4%f%7TX_MAK-AE________VisioStd2021Volume
16_fb33d997-4aa3-494e-8b58-03e9ab0f181d_VNCC4-CJQVK-BKX34-77Y8H-CY%f%XMR_Retail________Word2021Retail
16_0c728382-95fb-4a55-8f12-62e605f91727_BJG97-NW3GM-8QQQ7-FH76G-68%f%6XM_MAK-AE________Word2021Volume
:: Office 2024
16_8fdb1f1e-663f-4f2e-8fdb-7c35aee7d5ea_GNXWX-DF797-B2JT3-82W27-KH%f%PXT_MAK-AE________ProPlus2024Volume-Preview
16_33b11b14-91fd-4f7b-b704-e64a055cf601_X86XX-N3QMW-B4WGQ-QCB69-V2%f%6KW_MAK-AE________ProjectPro2024Volume-Preview
16_eb074198-7384-4bdd-8e6c-c3342dac8435_DW99Y-H7NT6-6B29D-8JQ8F-R3%f%QT7_MAK-AE________VisioPro2024Volume-Preview
16_e563d108-7b0e-418a-8390-20e1d133d6bb_P6NMW-JMTRC-R6MQ6-HH3F2-BT%f%HKB_Retail________Access2024Retail
16_f748e2f7-5951-4bc2-8a06-5a1fbe42f5f4_CXNJT-98HPP-92HX7-MX6GY-2P%f%VFR_MAK-AE________Access2024Volume
16_f3a5e86a-e4f8-4d88-8220-1440c3bbcefa_82CNJ-W82TW-BY23W-BVJ6W-W4%f%8GP_Retail________Excel2024Retail
16_523fbbab-c290-460d-a6c9-48e49709cb8e_7Y287-9N2KC-8MRR3-BKY82-2D%f%QRV_MAK-AE________Excel2024Volume
16_885f83e0-5e18-4199-b8be-56697d0debfb_N69X7-73KPT-899FD-P8HQ4-QG%f%TP4_Retail________Home2024Retail
16_acd4eccb-ff89-4e6a-9350-d2d56276ec69_PRKQM-YNPQR-77QT6-328D7-BD%f%223_Retail________HomeBusiness2024Retail
16_6f5fd645-7119-44a4-91b4-eccfeeb738bf_2CFK4-N44KG-7XG89-CWDG6-P7%f%P27_Retail________Outlook2024Retail
16_9a1e1bac-2d8b-4890-832f-0a68b27c16e0_NQPXP-WVB87-H3MMB-FYBW2-9Q%f%FPB_MAK-AE________Outlook2024Volume
16_da9a57ae-81a8-4cb3-b764-5840e6b5d0bf_CT2KT-GTNWH-9HFGW-J2PWJ-XW%f%7KJ_Retail________PowerPoint2024Retail
16_eca0d8a6-e21b-4622-9a87-a7103ff14012_RRXFN-JJ26R-RVWD2-V7WMP-27%f%PWQ_MAK-AE________PowerPoint2024Volume
16_295dcc21-151a-4b4d-8f50-2b627ea197f6_GNJ6P-Y4RBM-C32WW-2VJKJ-MT%f%HKK_Retail________ProjectPro2024Retail
16_2141d341-41aa-4e45-9ca1-201e117d6495_WNFMR-HK4R7-7FJVM-VQ3JC-76%f%HF6_MAK-AE1_______ProjectPro2024Volume
16_ead42f74-817d-45b4-af6b-3beeb36ba650_C2PNM-2GQFC-CY3XR-WXCP4-GX%f%3XM_Retail________ProjectStd2024Retail
16_4b6d9b9b-c16e-429d-babe-8bb84c3c27d6_F2VNW-MW8TT-K622Q-4D96H-PW%f%J8X_MAK-AE________ProjectStd2024Volume
16_db249714-bb54-4422-8c78-2cc8d4c4a19f_VWCNX-7FKBD-FHJYG-XBR4B-88%f%KC6_Retail________ProPlus2024Retail
16_d77244dc-2b82-4f0a-b8ae-1fca00b7f3e2_4YV2J-VNG7W-YGTP3-443TK-TF%f%8CP_MAK-AE1_______ProPlus2024Volume
16_3046a03e-2277-4a51-8ccd-a6609eae8c19_XKRBW-KN2FF-G8CKY-HXVG6-FV%f%Y2V_MAK-AE________SkypeforBusiness2024Volume
16_44a07f51-8263-4b2f-b2a5-70340055c646_GVG6N-6WCHH-K2MVP-RQ78V-3J%f%7GJ_MAK-AE1_______Standard2024Volume
16_282d8f34-1111-4a6f-80fe-c17f70dec567_HGRBX-N68QF-6DY8J-CGX4W-XW%f%7KP_Retail________VisioPro2024Retail
16_4c2f32bf-9d0b-4d8c-8ab1-b4c6a0b9992d_GBNHB-B2G3Q-G42YB-3MFC2-7C%f%JCX_MAK-AE________VisioPro2024Volume
16_8504167d-887a-41ae-bd1d-f849d834352d_VBXPJ-38NR3-C4DKF-C8RT7-RG%f%HKQ_Retail________VisioStd2024Retail
16_0978336b-5611-497c-9414-96effaff4938_YNFTY-63K7P-FKHXK-28YYT-D3%f%2XB_MAK-AE________VisioStd2024Volume
16_f6b24e61-6aa7-4fd2-ab9b-4046cee4230a_XN33R-RP676-GMY2F-T3MH7-GC%f%VKR_Retail________Word2024Retail
16_06142aa2-e935-49ca-af5d-08069a3d84f3_WD8CQ-6KNQM-8W2CX-2RT63-KK%f%3TP_MAK-AE________Word2024Volume
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do (

if %1==getinfo if not defined key (
if %oVer%==%%A if /i "%2"=="%%E" (
set key=%%C
set _actid=%%B
set _allactid=!_allactid! %%B
set _lic=%%D
if %oVer%==16 (echo "%%D" | find /i "Subscription" %nul% && set _sublic=1)
)
)

if %1==getmsiprod if %oVer%==%%A (
for /f "tokens=*" %%x in ('findstr /i /c:"%%B" "%_oBranding%"') do set "prodId=%%x"
set prodId=!prodId:"/>=!
set prodId=!prodId:~-4!
reg query "%2\Registration\{%%B}" /v ProductCode %nul2% | find /i "-!prodId!-" %nul% && (
reg query "%2\Common\InstalledPackages" %nul2% | find /i "-!prodId!-" %nul% && (
if defined _oIds (set _oIds=!_oIds! %%E) else (set _oIds=%%E)
)
)
)

)
)
exit /b

::========================================================================================================================================

::  This code is used to modify the timestamp value of sppc dll file in order to change checksums
::  It's done to lower the potential false positive detection by antivirus's. On each install, it will install a unique sppc dll file

:oh_extractdll

set b=
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':%_hook%\:.*';$encoded = ($f[1]) -replace '-', 'A' -replace '_', 'a';$bytes = [Con%b%vert]::FromBas%b%e64String($encoded); $PePath='%1'; $offset='%2'; $m=[io.file]::ReadAllText('!_batp!') -split ':hexedit\:.*';iex ($m[1]);" %nul2% | find /i "Error found" %nul1% && set hasherror=1
exit /b

:hexedit:
# Use a MemoryStream to perform operations on the bytes
$MemoryStream = New-Object System.IO.MemoryStream
$Writer = New-Object System.IO.BinaryWriter($MemoryStream)
$Writer.Write($bytes)

# Define dynamic assembly, module, and type
$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TypeBuilder = $ModuleBuilder.DefineType(0)

# Define P/Invoke method
[void]$TypeBuilder.DefinePInvokeMethod('MapFileAndCheckSum', 'imagehlp.dll', 'Public, Static', [Reflection.CallingConventions]::Standard, [int], @([string], [int].MakeByRefType(), [int].MakeByRefType()), [Runtime.InteropServices.CallingConvention]::Winapi, [Runtime.InteropServices.CharSet]::Auto)

# Create the type
$Imagehlp = $TypeBuilder.CreateType()

# Offset information
$timestampOffset = 136
$exportTimestampOffset = $offset
$checkSumOffset = 216

# Calculate timestamp
$currentTimestamp = [DateTime]::UtcNow
$unixTimestamp = [int]($currentTimestamp - (Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0)).TotalSeconds

# Change timestamps
$Writer.BaseStream.Position = $timestampOffset
$Writer.Write($unixTimestamp)

$Writer.BaseStream.Position = $exportTimestampOffset
$Writer.Write($unixTimestamp)

$Writer.Flush()

# Write the current state of the MemoryStream to a temporary file
$tempFilePath = [System.IO.Path]::Combine($env:windir, "Temp", [System.IO.Path]::GetRandomFileName())
[System.IO.File]::WriteAllBytes($tempFilePath, $MemoryStream.ToArray())

# Update hash using the temporary file
[int]$HeaderSum = 0
[int]$CheckSum = 0
[void]$Imagehlp::MapFileAndCheckSum($tempFilePath, [ref]$HeaderSum, [ref]$CheckSum)

# If the checksums don't match, update the checksum in the MemoryStream
if ($HeaderSum -ne $CheckSum) {
    $Writer.BaseStream.Position = $checkSumOffset
    $Writer.Write($CheckSum)
    $Writer.Flush()
} else {
    Write-host Error found
}

# Delete the temporary file
Remove-Item -Path $tempFilePath -Force

# Get the modified bytes
$modifiedBytes = $MemoryStream.ToArray()

# Write the modified bytes to the final file
[System.IO.File]::WriteAllBytes($PePath, $modifiedBytes)

[void]$Imagehlp::MapFileAndCheckSum($PePath, [ref]$HeaderSum, [ref]$CheckSum)
if ($HeaderSum -ne $CheckSum) {
    Write-host Error found
}

$MemoryStream.Close()
:hexedit:

::========================================================================================================================================
::
::  This below blocks of text is encoded in base64 format
::  The blocks in labels "sppc32.dll" and "sppc64.dll" contains below files
::
::  09865ea5993215965e8f27a74b8a41d15fd0f60f5f404cb7a8b3c7757acdab02 *sppc32.dll
::  393a1fa26deb3663854e41f2b687c188a9eacd87b23f17ea09422c4715cb5a9f *sppc64.dll
::
::  The files are encoded in base64 to make AIO version.
::
::  mass grave[.]dev/ohook
::  Here you can find the files source code and info on how to rebuild the identical sppc.dll files
::
::  stackoverflow.com/a/35335273
::  Here you can check how to extract sppc.dll files from base64
::
::  For any further question, feel free to contact us on mass grave[.]dev/contactus
::
::========================================================================================================================================
::
::  If you want to use a different sppc.dll or without base64 format, then create a folder named "BIN" where this script is located and 
::  place these two files in that "BIN" folder. sppc32.dll, sppc64.dll
::  Script will auto pick that instead of using the below from base64 section. You can also delete the below code in that case.
::
::========================================================================================================================================
::
::  Replace "-" with "A" and "_" with "a" before base64 conversion
::  It was changed to prevent antiviruses from detecting and flagging base64 encoding

:sppc32.dll:
TVqQ--M----E----//8--Lg---------Q-----------------------------------------------g-----4fug4-t-nNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4g_W4gRE9TIG1vZGUuDQ0KJ---------BQRQ--T-EH-MDc0GQ----------O--
DiML-QIo--I----e--------RxE----Q----------C-_g-Q-----g--B-----E----G----------CQ----B---+dY---I-Q-E--C---B------E---E--------B------Q---jR----Bg---Y-Q---H---HgD-------------------------I---BQ---------
----------------------------------------------------------BsY---H------------------------------------C50ZXh0----c-E----Q-----g----Q------------------C---G-ucmRhdGE--Bg-----I-----I----G----------------
--B---B-LmVoX2ZyYW2------D-----C----C-------------------Q---QC5lZGF0YQ--jR----B-----Eg----o------------------E---E-u_WRhdGE--BgB----Y-----I----c------------------B---D-LnJzcmM---B4-w---H-----E----Hg--
----------------Q---wC5yZWxvYw--F-----C------g---CI------------------E---EI-----------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------------------------------------------------------FWJ5VZTjUXwg+wwx0Xw-----IlEJBSNRfSJ
RCQQi0UMx0QkD-----CJRCQEi0UIx0QkC--ggGqJBCTHRfQ-----6-oB--CLNXhggGqD7BiFwInDi0Xwd-qJBCQx2//WUesyi1X0x0QkB-oggGqJBCSJVCQI/xW-YIBqg+wMhcCLRfCJBCR0Cv/WuwE---BS6wP/1lCNZfiJ2FteXcNVieVXVlOD7DyLRRiLdRyJRCQQ
i0UUiXQkFIlEJ-yLRRCJRCQIi0UMiUQkBItFCIkEJOiE----McmD7BiJx4X-dVyLRRg5CHZV_9koiwYB2IN4E-B0RYlEJ-SLRQiJTeSJBCTo+/7//4tN5IX-dSwDHsdDE-E---DHQxQ-----x0MY-----MdDH-----DHQy------x0Mk-----EHrpI1l9In4W15fXcIY
-LgB----wgw-kP8lcGC-_pCQ/yVsYIBqkJD/////-----P////8-----------------------------------------------------------------------------------------------------------------------------------------------------
------------------------------------------------TgBh-G0-ZQ---Ec-cgBh-GM-ZQ------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------------------------------------U----------F6Ug-Bf-gBGwwEBIgB---k----H----ODf//+d-----EEOCIUCQg0FSIYD
gwQCj8NBxkHFD-QEK----EQ---BV4P//qg----BBDgiF-kINBU_H-4YEgwUCm8NBxkHHQcUMB-QQ----c----NPg//8I------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
------------------D-3NBk-----MZC---B----Qw---EM----oQ---NEE--EBC--DPQg--70I---VD---pQw--XUM--KFD--DpQw--F0Q--DVE--BnR---nUQ--ONE---tRQ--YUU--J9F--DTRQ--DUY--DtG--BxRg--r0Y--M9G--D7Rg--nR---FFH--BvRw--
n0c--NNH---RS---TUg--G9I--ClS---zUg---VJ--BBSQ--bUk--KdJ--C7SQ--+0k--DlK--BPSg--dUo--J1K--DTSg--B0s--D1L--BpSw--pUs--ONL---NT---OUw--IlM--DRT---EU0--FlN--CjTQ--8U0--BtO--BHTg--h04--LtO--DnTg--K08--FtP
--C1Tw--608--CdQ--BdU---4kI--P1C---_Qw--RkM--IJD--DIQw---0Q--ClE--BRR---hUQ--MNE---LRQ--SkU--INF--C8RQ--80U--CdG--BZRg--k0Y--MJG--DoRg--GUc--DFH--BjRw--ikc--LxH--D1Rw--Mkg--GFI--CNS---vEg--OxI---mSQ--
Wkk--I1J--C0SQ--3kk--B1K--BHSg--ZUo--IxK--C7Sg--8Eo--CVL--BWSw--iks--MdL--D7Sw--Jkw--GRM--CwT---9Ew--DhN--CBTQ--zU0---lO---0Tg--_k4--KRO--DUTg--DE8--EZP--CLTw--008---xQ--BFU---eF-------Q-C--M-B--F--Y-
Bw-I--k-Cg-L--w-DQ-O--8-E--R-BI-Ew-U-BU-Fg-X-Bg-GQ-_-Bs-H--d-B4-Hw-g-CE-Ig-j-CQ-JQ-m-Cc-K--p-Co-Kw-s-C0-Lg-v-D--MQ-y-DM-N--1-DY-Nw-4-Dk-Og-7-Dw-PQ-+-D8-Q-BB-EI-c3BwYy5kbGw-U1BQQ1MuU0xDYWxsU2VydmVy-FNM
Q2FsbFNlcnZlcgBTUFBDUy5TTENsb3Nl-FNMQ2xvc2U-U1BQQ1MuU0xDb25zdW1lUmln_HQ-U0xDb25zdW1lUmln_HQ-U1BQQ1MuU0xEZXBvc2l0TWlncmF0_W9uQmxvYgBTTERlcG9z_XRN_WdyYXRpb25CbG9i-FNQUENTLlNMRGVwb3NpdE9mZmxpbmVDb25m_XJt
YXRpb25JZ-BTTERlcG9z_XRPZmZs_W5lQ29uZmlybWF0_W9uSWQ-U1BQQ1MuU0xEZXBvc2l0T2ZmbGluZUNvbmZpcm1hdGlvbklkRXg-U0xEZXBvc2l0T2ZmbGluZUNvbmZpcm1hdGlvbklkRXg-U1BQQ1MuU0xEZXBvc2l0U3RvcmVUb2tlbgBTTERlcG9z_XRTdG9y
ZVRv_2Vu-FNQUENTLlNMRmlyZUV2ZW50-FNMRmlyZUV2ZW50-FNQUENTLlNMR2F0_GVyTWlncmF0_W9uQmxvYgBTTEdhdGhlck1pZ3JhdGlvbkJsb2I-U1BQQ1MuU0xHYXRoZXJN_WdyYXRpb25CbG9iRXg-U0xHYXRoZXJN_WdyYXRpb25CbG9iRXg-U1BQQ1MuU0xH
ZW5lcmF0ZU9mZmxpbmVJbnN0YWxsYXRpb25JZ-BTTEdlbmVyYXRlT2ZmbGluZUluc3RhbGxhdGlvbklk-FNQUENTLlNMR2VuZXJhdGVPZmZs_W5lSW5zdGFsbGF0_W9uSWRFe-BTTEdlbmVyYXRlT2ZmbGluZUluc3RhbGxhdGlvbklkRXg-U1BQQ1MuU0xHZXRBY3Rp
dmVM_WNlbnNlSW5mbwBTTEdldEFjdGl2ZUxpY2Vuc2VJbmZv-FNQUENTLlNMR2V0QXBwbGljYXRpb25JbmZvcm1hdGlvbgBTTEdldEFwcGxpY2F0_W9uSW5mb3JtYXRpb24-U1BQQ1MuU0xHZXRBcHBs_WNhdGlvblBvbGljeQBTTEdldEFwcGxpY2F0_W9uUG9s_WN5
-FNQUENTLlNMR2V0QXV0_GVudGljYXRpb25SZXN1bHQ-U0xHZXRBdXRoZW50_WNhdGlvblJlc3Vsd-BTUFBDUy5TTEdldEVuY3J5cHRlZFBJREV4-FNMR2V0RW5jcnlwdGVkUElERXg-U1BQQ1MuU0xHZXRHZW51_W5lSW5mb3JtYXRpb24-U0xHZXRHZW51_W5lSW5m
b3JtYXRpb24-U1BQQ1MuU0xHZXRJbnN0YWxsZWRQcm9kdWN0S2V5SWRz-FNMR2V0SW5zdGFsbGVkUHJvZHVjdEtleUlkcwBTUFBDUy5TTEdldExpY2Vuc2U-U0xHZXRM_WNlbnNl-FNQUENTLlNMR2V0TGljZW5zZUZpbGVJZ-BTTEdldExpY2Vuc2VG_WxlSWQ-U1BQ
Q1MuU0xHZXRM_WNlbnNlSW5mb3JtYXRpb24-U0xHZXRM_WNlbnNlSW5mb3JtYXRpb24-U0xHZXRM_WNlbnNpbmdTdGF0dXNJbmZvcm1hdGlvbgBTUFBDUy5TTEdldFBLZXlJZ-BTTEdldFBLZXlJZ-BTUFBDUy5TTEdldFBLZXlJbmZvcm1hdGlvbgBTTEdldFBLZXlJ
bmZvcm1hdGlvbgBTUFBDUy5TTEdldFBvbGljeUluZm9ybWF0_W9u-FNMR2V0UG9s_WN5SW5mb3JtYXRpb24-U1BQQ1MuU0xHZXRQb2xpY3lJbmZvcm1hdGlvbkRXT1JE-FNMR2V0UG9s_WN5SW5mb3JtYXRpb25EV09SR-BTUFBDUy5TTEdldFByb2R1Y3RT_3VJbmZv
cm1hdGlvbgBTTEdldFByb2R1Y3RT_3VJbmZvcm1hdGlvbgBTUFBDUy5TTEdldFNMSURM_XN0-FNMR2V0U0xJRExpc3Q-U1BQQ1MuU0xHZXRTZXJ2_WNlSW5mb3JtYXRpb24-U0xHZXRTZXJ2_WNlSW5mb3JtYXRpb24-U1BQQ1MuU0xJbnN0YWxsTGljZW5zZQBTTElu
c3RhbGxM_WNlbnNl-FNQUENTLlNMSW5zdGFsbFByb29mT2ZQdXJj_GFzZQBTTEluc3RhbGxQcm9vZk9mUHVyY2hhc2U-U1BQQ1MuU0xJbnN0YWxsUHJvb2ZPZlB1cmNoYXNlRXg-U0xJbnN0YWxsUHJvb2ZPZlB1cmNoYXNlRXg-U1BQQ1MuU0xJc0dlbnVpbmVMb2Nh
bEV4-FNMSXNHZW51_W5lTG9jYWxFe-BTUFBDUy5TTExvYWRBcHBs_WNhdGlvblBvbGlj_WVz-FNMTG9hZEFwcGxpY2F0_W9uUG9s_WNpZXM-U1BQQ1MuU0xPcGVu-FNMT3BlbgBTUFBDUy5TTFBlcnNpc3RBcHBs_WNhdGlvblBvbGlj_WVz-FNMUGVyc2lzdEFwcGxp
Y2F0_W9uUG9s_WNpZXM-U1BQQ1MuU0xQZXJz_XN0UlRTUGF5bG9hZE92ZXJy_WRl-FNMUGVyc2lzdFJUU1BheWxvYWRPdmVycmlkZQBTUFBDUy5TTFJlQXJt-FNMUmVBcm0-U1BQQ1MuU0xSZWdpc3RlckV2ZW50-FNMUmVn_XN0ZXJFdmVud-BTUFBDUy5TTFJlZ2lz
dGVyUGx1Z2lu-FNMUmVn_XN0ZXJQbHVn_W4-U1BQQ1MuU0xTZXRBdXRoZW50_WNhdGlvbkRhdGE-U0xTZXRBdXRoZW50_WNhdGlvbkRhdGE-U1BQQ1MuU0xTZXRDdXJyZW50UHJvZHVjdEtleQBTTFNldEN1cnJlbnRQcm9kdWN0S2V5-FNQUENTLlNMU2V0R2VudWlu
ZUluZm9ybWF0_W9u-FNMU2V0R2VudWluZUluZm9ybWF0_W9u-FNQUENTLlNMVW5pbnN0YWxsTGljZW5zZQBTTFVu_W5zdGFsbExpY2Vuc2U-U1BQQ1MuU0xVbmluc3RhbGxQcm9vZk9mUHVyY2hhc2U-U0xVbmluc3RhbGxQcm9vZk9mUHVyY2hhc2U-U1BQQ1MuU0xV
bmxvYWRBcHBs_WNhdGlvblBvbGlj_WVz-FNMVW5sb2FkQXBwbGljYXRpb25Qb2xpY2llcwBTUFBDUy5TTFVucmVn_XN0ZXJFdmVud-BTTFVucmVn_XN0ZXJFdmVud-BTUFBDUy5TTFVucmVn_XN0ZXJQbHVn_W4-U0xVbnJlZ2lzdGVyUGx1Z2lu-FNQUENTLlNMcEF1
dGhlbnRpY2F0ZUdlbnVpbmVU_WNrZXRSZXNwb25zZQBTTHBBdXRoZW50_WNhdGVHZW51_W5lVGlj_2V0UmVzcG9uc2U-U1BQQ1MuU0xwQmVn_W5HZW51_W5lVGlj_2V0VHJhbnNhY3Rpb24-U0xwQmVn_W5HZW51_W5lVGlj_2V0VHJhbnNhY3Rpb24-U1BQQ1MuU0xw
Q2xlYXJBY3RpdmF0_W9uSW5Qcm9ncmVzcwBTTHBDbGVhckFjdGl2YXRpb25JblByb2dyZXNz-FNQUENTLlNMcERlcG9z_XREb3dubGV2ZWxHZW51_W5lVGlj_2V0-FNMcERlcG9z_XREb3dubGV2ZWxHZW51_W5lVGlj_2V0-FNQUENTLlNMcERlcG9z_XRUb2tlbkFj
dGl2YXRpb25SZXNwb25zZQBTTHBEZXBvc2l0VG9rZW5BY3RpdmF0_W9uUmVzcG9uc2U-U1BQQ1MuU0xwR2VuZXJhdGVUb2tlbkFjdGl2YXRpb25D_GFsbGVuZ2U-U0xwR2VuZXJhdGVUb2tlbkFjdGl2YXRpb25D_GFsbGVuZ2U-U1BQQ1MuU0xwR2V0R2VudWluZUJs
b2I-U0xwR2V0R2VudWluZUJsb2I-U1BQQ1MuU0xwR2V0R2VudWluZUxvY2Fs-FNMcEdldEdlbnVpbmVMb2Nhb-BTUFBDUy5TTHBHZXRM_WNlbnNlQWNxdWlz_XRpb25JbmZv-FNMcEdldExpY2Vuc2VBY3F1_XNpdGlvbkluZm8-U1BQQ1MuU0xwR2V0TVNQ_WRJbmZv
cm1hdGlvbgBTTHBHZXRNU1BpZEluZm9ybWF0_W9u-FNQUENTLlNMcEdldE1hY2hpbmVVR1VJR-BTTHBHZXRNYWNo_W5lVUdVSUQ-U1BQQ1MuU0xwR2V0VG9rZW5BY3RpdmF0_W9uR3JhbnRJbmZv-FNMcEdldFRv_2VuQWN0_XZhdGlvbkdyYW50SW5mbwBTUFBDUy5T
THBJQUFjdGl2YXRlUHJvZHVjd-BTTHBJQUFjdGl2YXRlUHJvZHVjd-BTUFBDUy5TTHBJc0N1cnJlbnRJbnN0YWxsZWRQcm9kdWN0S2V5RGVmYXVsdEtleQBTTHBJc0N1cnJlbnRJbnN0YWxsZWRQcm9kdWN0S2V5RGVmYXVsdEtleQBTUFBDUy5TTHBQcm9jZXNzVk1Q
_XBlTWVzc2FnZQBTTHBQcm9jZXNzVk1Q_XBlTWVzc2FnZQBTUFBDUy5TTHBTZXRBY3RpdmF0_W9uSW5Qcm9ncmVzcwBTTHBTZXRBY3RpdmF0_W9uSW5Qcm9ncmVzcwBTUFBDUy5TTHBUcmlnZ2VyU2VydmljZVdvcmtlcgBTTHBUcmlnZ2VyU2VydmljZVdvcmtlcgBT
UFBDUy5TTHBWTEFjdGl2YXRlUHJvZHVjd-BTTHBWTEFjdGl2YXRlUHJvZHVjd-------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------------------------------------------------------------------------------FBg-------------Ohg--BsY---XG--------------
+G---Hhg--BkY--------------MYQ--gG------------------------------iG---Kpg--------yG--------DUY--------Ihg--CqY--------Mhg--------1G---------C-FNMR2V0TGljZW5z_W5nU3RhdHVzSW5mb3JtYXRpb24--QBTTEdldFByb2R1
Y3RT_3VJbmZvcm1hdGlvbg--3QNMb2NhbEZyZWU-RwFTdHJTdHJOSVc--G----Bg--BzcHBjcy5kbGw----UY---S0VSTkVMMzIuZGxs-----Chg--BTSExXQVBJLmRsb-----------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------------------B-B-----Y--C--------------------B--E----w--C--------------------B--kE--BI----WH---BwD-------------BwDN----FY-UwBf-FY-RQBS-FM-SQBP-E4-XwBJ-E4-
RgBP------C9BO/+---B--U---------BQ--------------------Q-B--C--------------------f-I---E-UwB0-HI-_QBu-Gc-RgBp-Gw-ZQBJ-G4-ZgBv----W-I---E-M--0-D--OQ-w-DQ-RQ-0----eg-t--E-QwBv-G0-c-Bh-G4-eQBO-GE-bQBl----
--BB-G4-bwBt-GE-b-Bv-HU-cw-g-FM-bwBm-HQ-dwBh-HI-ZQ-g-EQ-ZQB0-GU-cgBp-G8-cgBh-HQ-_QBv-G4-I-BD-G8-cgBw-G8-cgBh-HQ-_QBv-G4------D4-Cw-B-EY-_QBs-GU-R-Bl-HM-YwBy-Gk-c-B0-Gk-bwBu------Bv-Gg-bwBv-Gs-I-BT-F--
U-BD-------w--g--QBG-Gk-b-Bl-FY-ZQBy-HM-_QBv-G4------D--Lg-1-C4-M--u-D-----q--U--QBJ-G4-d-Bl-HI-bgBh-Gw-TgBh-G0-ZQ---HM-c-Bw-GM------Iw-N--B-Ew-ZQBn-GE-b-BD-G8-c-B5-HI-_QBn-Gg-d----Kk-I--y-D--Mg-0-C--
QQBu-G8-bQBh-Gw-bwB1-HM-I-BT-G8-ZgB0-Hc-YQBy-GU-I-BE-GU-d-Bl-HI-_QBv-HI-YQB0-Gk-bwBu-C--QwBv-HI-c-Bv-HI-YQB0-Gk-bwBu----Og-J--E-TwBy-Gk-ZwBp-G4-YQBs-EY-_QBs-GU-bgBh-G0-ZQ---HM-c-Bw-GM-LgBk-Gw-b-------
L--G--E-U-By-G8-Z-B1-GM-d-BO-GE-bQBl------Bv-Gg-bwBv-Gs----0--g--QBQ-HI-bwBk-HU-YwB0-FY-ZQBy-HM-_QBv-G4----w-C4-NQ-u-D--Lg-w----R-----E-VgBh-HI-RgBp-Gw-ZQBJ-G4-ZgBv-------k--Q---BU-HI-YQBu-HM-b-Bh-HQ-
_QBv-G4-------kE5-Q-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-------Q---U----MzBIMGkwdjBSMVox------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------
:sppc32.dll:

:========================================================================================================================================

::  Replace "-" with "A" and "_" with "a" before base64 conversion
::  It was changed to prevent antiviruses from detecting and flagging base64 encoding

:sppc64.dll:
TVqQ--M----E----//8--Lg---------Q-----------------------------------------------g-----4fug4-t-nNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4g_W4gRE9TIG1vZGUuDQ0KJ---------BQRQ--ZIYH-MDc0GQ----------P--
LiIL-gIo--I----e--------ExE----Q-----JIx-g-----Q-----g--B----------G----------CQ----B---LeY---I-Y-E--C---------Q-----------Q--------E--------------Q-----F---I0Q----c---U-E---C---B4-w---D---CQ---------
--------------------------------------------------------------------------------iH---Dg------------------------------------udGV4d----H-B----E-----I----E-------------------g--BgLnJkYXRh---g-----C-----C
----Bg------------------Q---QC5wZGF0YQ--J------w-----g----g------------------E---E-ueGRhdGE--CQ-----Q-----I----K------------------B---B-LmVkYXRh--CNE----F-----S----D-------------------Q---QC5pZGF0YQ--
U-E---Bw-----g---B4------------------E---M-ucnNyYw---HgD----g-----Q----g------------------B---D---------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------------------------------------------------------EFUU0iD7EhFMclMjQXvDw--SI1EJDjHRCQ0
-----EiJRCQoSI1EJDRIiUQkIEjHRCQ4-----OgF-Q--SItMJDhIix1ZY---hcBBicR0B//TRTHk6yhEi0QkNEiNF_kP--D/FUlg--BIi0wkOEiFwHQK/9NBv-E---Dr-v/TRIngSIPESFtBXMNBVUFUVVdWU0iD7Dgx9kyLrCSQ----SIusJJg---BMiWwkIEiJz0iJ
bCQo6J----BBicSFwHVEQTl1-HY+SGveKEiLVQBI-dqDeh--dChIifnoIv///4X-dRxI-10-SMdDE-E---BIx0MY-----EjHQy------SP/G67xEieBIg8Q4W15fXUFcQV3Du-E---DDkJCQkJCQkP8lel8--JCQDx+E------D/JXpf--CQk-8fh-------/yVKXw--
kJD/JTpf--CQkP//////////----------D//////////w----------------------------------------------------------------------------------------------------------------------------------------------------------
------------------------------------------------TgBh-G0-ZQ---Ec-cgBh-GM-ZQ------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------------------E---iB----B---CIE---ExE---x----TEQ--GRE--CB-------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------EH-w-HggMw-s----EMBw-MYggwB2-Gc-VQBM-C0----Q----------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------MDc0GQ-----xlI---E---BD----Qw---ChQ---0UQ--QFI--M9S--DvUg--BVM--ClT--BdUw--oVM--OlT---XV---NVQ--GdU
--CdV---41Q--C1V--BhVQ--n1U--NNV---NVg--O1Y--HFW--CvVg--z1Y--PtW--CIE---UVc--G9X--CfVw--01c--BFY--BNW---b1g--KVY--DNW---BVk--EFZ--BtWQ--p1k--LtZ--D7WQ--OVo--E9_--B1Wg--nVo--NN_---HWw--PVs--Glb--ClWw--
41s---1c---5X---iVw--NFc---RXQ--WV0--KNd--DxXQ--G14--Ede--CHXg--u14--Ode---rXw--W18--LVf--DrXw--J2---F1g--DiUg--/VI--BpT--BGUw--glM--MhT---DV---KVQ--FFU--CFV---w1Q---tV--BKVQ--g1U--LxV--DzVQ--J1Y--FlW
--CTVg--wlY--OhW---ZVw--MVc--GNX--CKVw--vFc--PVX---yW---YVg--I1Y--C8W---7Fg--CZZ--B_WQ--jVk--LRZ--DeWQ--HVo--Ed_--BlWg--jFo--Lt_--DwWg--JVs--FZb--CKWw--x1s--Ptb---mX---ZFw--LBc--D0X---OF0--IFd--DNXQ--
CV4--DRe--BqXg--pF4--NRe---MXw--Rl8--Itf--DTXw--DG---EVg--B4Y------B--I--w-E--U-Bg-H--g-CQ-K--s-D--N--4-Dw-Q-BE-Eg-T-BQ-FQ-W-Bc-G--Z-Bo-Gw-c-B0-Hg-f-C--IQ-i-CM-J--l-CY-Jw-o-Ck-Kg-r-Cw-LQ-u-C8-M--x-DI-
Mw-0-DU-Ng-3-Dg-OQ-6-Ds-P--9-D4-PwB--EE-QgBzcHBjLmRsb-BTUFBDUy5TTENhbGxTZXJ2ZXI-U0xDYWxsU2VydmVy-FNQUENTLlNMQ2xvc2U-U0xDbG9zZQBTUFBDUy5TTENvbnN1bWVS_Wdod-BTTENvbnN1bWVS_Wdod-BTUFBDUy5TTERlcG9z_XRN_Wdy
YXRpb25CbG9i-FNMRGVwb3NpdE1pZ3JhdGlvbkJsb2I-U1BQQ1MuU0xEZXBvc2l0T2ZmbGluZUNvbmZpcm1hdGlvbklk-FNMRGVwb3NpdE9mZmxpbmVDb25m_XJtYXRpb25JZ-BTUFBDUy5TTERlcG9z_XRPZmZs_W5lQ29uZmlybWF0_W9uSWRFe-BTTERlcG9z_XRP
ZmZs_W5lQ29uZmlybWF0_W9uSWRFe-BTUFBDUy5TTERlcG9z_XRTdG9yZVRv_2Vu-FNMRGVwb3NpdFN0b3JlVG9rZW4-U1BQQ1MuU0xG_XJlRXZlbnQ-U0xG_XJlRXZlbnQ-U1BQQ1MuU0xHYXRoZXJN_WdyYXRpb25CbG9i-FNMR2F0_GVyTWlncmF0_W9uQmxvYgBT
UFBDUy5TTEdhdGhlck1pZ3JhdGlvbkJsb2JFe-BTTEdhdGhlck1pZ3JhdGlvbkJsb2JFe-BTUFBDUy5TTEdlbmVyYXRlT2ZmbGluZUluc3RhbGxhdGlvbklk-FNMR2VuZXJhdGVPZmZs_W5lSW5zdGFsbGF0_W9uSWQ-U1BQQ1MuU0xHZW5lcmF0ZU9mZmxpbmVJbnN0
YWxsYXRpb25JZEV4-FNMR2VuZXJhdGVPZmZs_W5lSW5zdGFsbGF0_W9uSWRFe-BTUFBDUy5TTEdldEFjdGl2ZUxpY2Vuc2VJbmZv-FNMR2V0QWN0_XZlTGljZW5zZUluZm8-U1BQQ1MuU0xHZXRBcHBs_WNhdGlvbkluZm9ybWF0_W9u-FNMR2V0QXBwbGljYXRpb25J
bmZvcm1hdGlvbgBTUFBDUy5TTEdldEFwcGxpY2F0_W9uUG9s_WN5-FNMR2V0QXBwbGljYXRpb25Qb2xpY3k-U1BQQ1MuU0xHZXRBdXRoZW50_WNhdGlvblJlc3Vsd-BTTEdldEF1dGhlbnRpY2F0_W9uUmVzdWx0-FNQUENTLlNMR2V0RW5jcnlwdGVkUElERXg-U0xH
ZXRFbmNyeXB0ZWRQSURFe-BTUFBDUy5TTEdldEdlbnVpbmVJbmZvcm1hdGlvbgBTTEdldEdlbnVpbmVJbmZvcm1hdGlvbgBTUFBDUy5TTEdldEluc3RhbGxlZFByb2R1Y3RLZXlJZHM-U0xHZXRJbnN0YWxsZWRQcm9kdWN0S2V5SWRz-FNQUENTLlNMR2V0TGljZW5z
ZQBTTEdldExpY2Vuc2U-U1BQQ1MuU0xHZXRM_WNlbnNlRmlsZUlk-FNMR2V0TGljZW5zZUZpbGVJZ-BTUFBDUy5TTEdldExpY2Vuc2VJbmZvcm1hdGlvbgBTTEdldExpY2Vuc2VJbmZvcm1hdGlvbgBTTEdldExpY2Vuc2luZ1N0YXR1c0luZm9ybWF0_W9u-FNQUENT
LlNMR2V0UEtleUlk-FNMR2V0UEtleUlk-FNQUENTLlNMR2V0UEtleUluZm9ybWF0_W9u-FNMR2V0UEtleUluZm9ybWF0_W9u-FNQUENTLlNMR2V0UG9s_WN5SW5mb3JtYXRpb24-U0xHZXRQb2xpY3lJbmZvcm1hdGlvbgBTUFBDUy5TTEdldFBvbGljeUluZm9ybWF0
_W9uRFdPUkQ-U0xHZXRQb2xpY3lJbmZvcm1hdGlvbkRXT1JE-FNQUENTLlNMR2V0UHJvZHVjdFNrdUluZm9ybWF0_W9u-FNMR2V0UHJvZHVjdFNrdUluZm9ybWF0_W9u-FNQUENTLlNMR2V0U0xJRExpc3Q-U0xHZXRTTElETGlzd-BTUFBDUy5TTEdldFNlcnZpY2VJ
bmZvcm1hdGlvbgBTTEdldFNlcnZpY2VJbmZvcm1hdGlvbgBTUFBDUy5TTEluc3RhbGxM_WNlbnNl-FNMSW5zdGFsbExpY2Vuc2U-U1BQQ1MuU0xJbnN0YWxsUHJvb2ZPZlB1cmNoYXNl-FNMSW5zdGFsbFByb29mT2ZQdXJj_GFzZQBTUFBDUy5TTEluc3RhbGxQcm9v
Zk9mUHVyY2hhc2VFe-BTTEluc3RhbGxQcm9vZk9mUHVyY2hhc2VFe-BTUFBDUy5TTElzR2VudWluZUxvY2FsRXg-U0xJc0dlbnVpbmVMb2NhbEV4-FNQUENTLlNMTG9hZEFwcGxpY2F0_W9uUG9s_WNpZXM-U0xMb2FkQXBwbGljYXRpb25Qb2xpY2llcwBTUFBDUy5T
TE9wZW4-U0xPcGVu-FNQUENTLlNMUGVyc2lzdEFwcGxpY2F0_W9uUG9s_WNpZXM-U0xQZXJz_XN0QXBwbGljYXRpb25Qb2xpY2llcwBTUFBDUy5TTFBlcnNpc3RSVFNQYXlsb2FkT3ZlcnJpZGU-U0xQZXJz_XN0UlRTUGF5bG9hZE92ZXJy_WRl-FNQUENTLlNMUmVB
cm0-U0xSZUFybQBTUFBDUy5TTFJlZ2lzdGVyRXZlbnQ-U0xSZWdpc3RlckV2ZW50-FNQUENTLlNMUmVn_XN0ZXJQbHVn_W4-U0xSZWdpc3RlclBsdWdpbgBTUFBDUy5TTFNldEF1dGhlbnRpY2F0_W9uRGF0YQBTTFNldEF1dGhlbnRpY2F0_W9uRGF0YQBTUFBDUy5T
TFNldEN1cnJlbnRQcm9kdWN0S2V5-FNMU2V0Q3VycmVudFByb2R1Y3RLZXk-U1BQQ1MuU0xTZXRHZW51_W5lSW5mb3JtYXRpb24-U0xTZXRHZW51_W5lSW5mb3JtYXRpb24-U1BQQ1MuU0xVbmluc3RhbGxM_WNlbnNl-FNMVW5pbnN0YWxsTGljZW5zZQBTUFBDUy5T
TFVu_W5zdGFsbFByb29mT2ZQdXJj_GFzZQBTTFVu_W5zdGFsbFByb29mT2ZQdXJj_GFzZQBTUFBDUy5TTFVubG9hZEFwcGxpY2F0_W9uUG9s_WNpZXM-U0xVbmxvYWRBcHBs_WNhdGlvblBvbGlj_WVz-FNQUENTLlNMVW5yZWdpc3RlckV2ZW50-FNMVW5yZWdpc3Rl
ckV2ZW50-FNQUENTLlNMVW5yZWdpc3RlclBsdWdpbgBTTFVucmVn_XN0ZXJQbHVn_W4-U1BQQ1MuU0xwQXV0_GVudGljYXRlR2VudWluZVRpY2tldFJlc3BvbnNl-FNMcEF1dGhlbnRpY2F0ZUdlbnVpbmVU_WNrZXRSZXNwb25zZQBTUFBDUy5TTHBCZWdpbkdlbnVp
bmVU_WNrZXRUcmFuc2FjdGlvbgBTTHBCZWdpbkdlbnVpbmVU_WNrZXRUcmFuc2FjdGlvbgBTUFBDUy5TTHBDbGVhckFjdGl2YXRpb25JblByb2dyZXNz-FNMcENsZWFyQWN0_XZhdGlvbkluUHJvZ3Jlc3M-U1BQQ1MuU0xwRGVwb3NpdERvd25sZXZlbEdlbnVpbmVU
_WNrZXQ-U0xwRGVwb3NpdERvd25sZXZlbEdlbnVpbmVU_WNrZXQ-U1BQQ1MuU0xwRGVwb3NpdFRv_2VuQWN0_XZhdGlvblJlc3BvbnNl-FNMcERlcG9z_XRUb2tlbkFjdGl2YXRpb25SZXNwb25zZQBTUFBDUy5TTHBHZW5lcmF0ZVRv_2VuQWN0_XZhdGlvbkNoYWxs
ZW5nZQBTTHBHZW5lcmF0ZVRv_2VuQWN0_XZhdGlvbkNoYWxsZW5nZQBTUFBDUy5TTHBHZXRHZW51_W5lQmxvYgBTTHBHZXRHZW51_W5lQmxvYgBTUFBDUy5TTHBHZXRHZW51_W5lTG9jYWw-U0xwR2V0R2VudWluZUxvY2Fs-FNQUENTLlNMcEdldExpY2Vuc2VBY3F1
_XNpdGlvbkluZm8-U0xwR2V0TGljZW5zZUFjcXVpc2l0_W9uSW5mbwBTUFBDUy5TTHBHZXRNU1BpZEluZm9ybWF0_W9u-FNMcEdldE1TUGlkSW5mb3JtYXRpb24-U1BQQ1MuU0xwR2V0TWFj_GluZVVHVUlE-FNMcEdldE1hY2hpbmVVR1VJR-BTUFBDUy5TTHBHZXRU
b2tlbkFjdGl2YXRpb25HcmFudEluZm8-U0xwR2V0VG9rZW5BY3RpdmF0_W9uR3JhbnRJbmZv-FNQUENTLlNMcElBQWN0_XZhdGVQcm9kdWN0-FNMcElBQWN0_XZhdGVQcm9kdWN0-FNQUENTLlNMcElzQ3VycmVudEluc3RhbGxlZFByb2R1Y3RLZXlEZWZhdWx0S2V5
-FNMcElzQ3VycmVudEluc3RhbGxlZFByb2R1Y3RLZXlEZWZhdWx0S2V5-FNQUENTLlNMcFByb2Nlc3NWTVBpcGVNZXNzYWdl-FNMcFByb2Nlc3NWTVBpcGVNZXNzYWdl-FNQUENTLlNMcFNldEFjdGl2YXRpb25JblByb2dyZXNz-FNMcFNldEFjdGl2YXRpb25JblBy
b2dyZXNz-FNQUENTLlNMcFRy_WdnZXJTZXJ2_WNlV29y_2Vy-FNMcFRy_WdnZXJTZXJ2_WNlV29y_2Vy-FNQUENTLlNMcFZMQWN0_XZhdGVQcm9kdWN0-FNMcFZMQWN0_XZhdGVQcm9kdWN0--------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
----------------------------------------UH--------------IHE--Ihw--Boc--------------wcQ--oH---Hhw-------------ERx--Cwc-----------------------------D-c--------OJw--------------------cQ------------------
DHE------------------MBw--------4n--------------------Bx-------------------McQ-------------------gBTTEdldExpY2Vuc2luZ1N0YXR1c0luZm9ybWF0_W9u--E-U0xHZXRQcm9kdWN0U2t1SW5mb3JtYXRpb24--OgDTG9jYWxGcmVl-FEB
U3RyU3RyTklX--Bw----c---c3BwY3MuZGxs----FH---EtFUk5FTDMyLmRsb------oc---U0hMV0FQSS5kbGw-----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------------------------------------E-E----Bg--I--------------------E--Q---D---I--------------
------E-CQQ--Eg---BYg---H-M-------------H-M0----VgBT-F8-VgBF-FI-UwBJ-E8-TgBf-Ek-TgBG-E8------L0E7/4---E-BQ---------F--------------------B--E--I-------------------B8-g---QBT-HQ-cgBp-G4-ZwBG-Gk-b-Bl-Ek-
bgBm-G8---BY-g---Q-w-DQ-M--5-D--N-BF-DQ---B6-C0--QBD-G8-bQBw-GE-bgB5-E4-YQBt-GU------EE-bgBv-G0-YQBs-G8-dQBz-C--UwBv-GY-d-B3-GE-cgBl-C--R-Bl-HQ-ZQBy-Gk-bwBy-GE-d-Bp-G8-bg-g-EM-bwBy-H--bwBy-GE-d-Bp-G8-
bg------Pg-L--E-RgBp-Gw-ZQBE-GU-cwBj-HI-_QBw-HQ-_QBv-G4------G8-_-Bv-G8-_w-g-FM-U-BQ-EM------D--C--B-EY-_QBs-GU-VgBl-HI-cwBp-G8-bg------M--u-DU-Lg-w-C4-M----Co-BQ-B-Ek-bgB0-GU-cgBu-GE-b-BO-GE-bQBl----
cwBw-H--Yw------j--0--E-T-Bl-Gc-YQBs-EM-bwBw-Hk-cgBp-Gc-_-B0----qQ-g-DI-M--y-DQ-I-BB-G4-bwBt-GE-b-Bv-HU-cw-g-FM-bwBm-HQ-dwBh-HI-ZQ-g-EQ-ZQB0-GU-cgBp-G8-cgBh-HQ-_QBv-G4-I-BD-G8-cgBw-G8-cgBh-HQ-_QBv-G4-
---6--k--QBP-HI-_QBn-Gk-bgBh-Gw-RgBp-Gw-ZQBu-GE-bQBl----cwBw-H--Yw-u-GQ-b-Bs-------s--Y--QBQ-HI-bwBk-HU-YwB0-E4-YQBt-GU------G8-_-Bv-G8-_w---DQ-C--B-F--cgBv-GQ-dQBj-HQ-VgBl-HI-cwBp-G8-bg---D--Lg-1-C4-
M--u-D----BE-----QBW-GE-cgBG-Gk-b-Bl-Ek-bgBm-G8------CQ-B----FQ-cgBh-G4-cwBs-GE-d-Bp-G8-bg------CQTkB---------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------
:sppc64.dll:

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:TSforgeActivation

::  To activate Windows, run the script with "/Z-Windows" parameter or change 0 to 1 in below line
set _actwin=0

::  To activate Windows ESU, run the script with "/Z-ESU" parameter or change 0 to 1 in below line
set _actesu=0

::  To activate all Office apps (including Project/Visio), run the script with "/Z-Office" parameter or change 0 to 1 in below line
set _actoff=0

::  To activate only Project/Visio, run the script with "/Z-ProjectVisio" parameter or change 0 to 1 in below line
set _actprojvis=0

::  To activate all Windows/ESU/Office, run the script with "/Z-WindowsESUOffice" parameter or change 0 to 1 in below line
set _actwinesuoff=0

::  Advanced options:
::  To activate Windows K-M-S host (csvlk), run the script with "/Z-WinHost" parameter or change 0 to 1 in below line
set _actwinhost=0

::  To activate Office K-M-S host (csvlk), run the script with "/Z-OffHost" parameter or change 0 to 1 in below line
set _actoffhost=0

::  To activate Windows 8/8.1 APPX Sideloading (APPXLOB), run the script with "/Z-APPX" parameter or change 0 to 1 in below line
set _actappx=0

::  To activate certain activation IDs, change 0 to 1 in below line and set activation IDs in "tsids" variable, you can enter multiple by adding a space after each of them
::  or run the script with "/Z-ID-ActivationIdGoesHere" parameter. If you want to add multiple through parameter, pass each of them in separate parameters
set _actman=
set tsids=

::  To reset rearm counter, evaluation period and clear the tamper state, key lock, run the script with "/Z-Reset" parameter or change 0 to 1 in below line
set _resall=0

::  Debug Mode:
::  To run the script in debug mode, change 0 to any parameter above that you want to run, in below line
set "_debug=0"

::  Script will run in unattended mode if parameters are used OR value is changed in above lines.
::  If multiple options are selected then script will only pick one from the advanced option.



::========================================================================================================================================

cls
color 07
set KS=K%blank%MS
title  TSforge Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args set _args=%_args:re1=%
if defined _args set _args=%_args:re2=%
if defined _args for %%A in (%_args%) do (
if /i "%%A"=="-el"                     (set _elev=1)
if /i "%%A"=="/Z-Windows"              (set _actwin=1)
if /i "%%A"=="/Z-ESU"                  (set _actesu=1)
if /i "%%A"=="/Z-Office"               (set _actoff=1)
if /i "%%A"=="/Z-ProjectVisio"         (set _actprojvis=1)
if /i "%%A"=="/Z-WindowsESUOffice"     (set _actwinesuoff=1)
if /i "%%A"=="/Z-WinHost"              (set _actwinhost=1)
if /i "%%A"=="/Z-OffHost"              (set _actoffhost=1)
if /i "%%A"=="/Z-APPX"                 (set _actappx=1)
echo "%%A" | find /i "/Z-ID-"  >nul && (set _actman=1& set "filtsids=%%A" & call set "filtsids=%%filtsids:~6%%" & if defined filtsids call set tsids=%%filtsids%% %%tsids%%)
if /i "%%A"=="/Z-Reset"                (set _resall=1)
)

if not defined tsids set _actman=0
for %%A in (%_actwin% %_actesu% %_actoff% %_actprojvis% %_actwinesuoff% %_actwinhost% %_actoffhost% %_actappx% %_actman% %_resall%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

:ts_menu

if %_unattended%==0 (
cls
if not defined terminal mode 76, 33
title  TSforge Activation %masver%

echo:
echo:
echo:
echo        ______________________________________________________________
echo: 
echo               [1] Activate - Windows
echo               [2] Activate - Windows [ESU]
echo               [3] Activate - Office  [All]
echo               [4] Activate - Office  [Project/Visio]
echo               [5] Activate - All
echo               _______________________________________________  
echo: 
echo                   Advanced Options:
echo:
echo               [A] Activate - Windows %KS% Host
echo               [B] Activate - Office %KS% Host
echo               [C] Activate - Windows 8/8.1 APPX Sideloading
echo               [D] Activate - Manually Select Products
echo               [E] Reset    - Rearm/Timers/Tamper/Lock
echo               _______________________________________________       
echo:
echo               [6] Remove TSforge Activation
echo               [7] Download Office
echo               [0] %_exitmsg%
echo        ______________________________________________________________
echo:
call :dk_color2 %_White% "            " %_Green% "Choose a menu option using your keyboard..."
choice /C:12345ABCDE670 /N
set _el=!errorlevel!

if !_el!==13 exit /b
if !_el!==12 start %mas%genuine-installation-media & goto :ts_menu
if !_el!==11 call :ts_remove & cls & goto :ts_menu
if !_el!==10 cls & setlocal & set "_resall=1"       & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==9  cls & setlocal & set "_actman=1"       & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==8  cls & setlocal & set "_actappx=1"      & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==7  cls & setlocal & set "_actoffhost=1"   & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==6  cls & setlocal & set "_actwinhost=1"   & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==5  cls & setlocal & set "_actwinesuoff=1" & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==4  cls & setlocal & set "_actprojvis=1"   & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==3  cls & setlocal & set "_actoff=1"       & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==2  cls & setlocal & set "_actesu=1"       & call :ts_start & endlocal & cls & goto :ts_menu
if !_el!==1  cls & setlocal & set "_actwin=1"       & call :ts_start & endlocal & cls & goto :ts_menu
goto :ts_menu
)

::========================================================================================================================================

:ts_start

cls

if %_actwinesuoff%==1 (set height=38) else (set height=32)
if not defined terminal (
mode 125, %height%
if exist "%SysPath%\spp\store_test\" mode 134, %height%
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=%height%;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)
title  TSforge Activation %masver%

echo:
echo Initializing...
call :dk_chkmal

if not exist %SysPath%\sppsvc.exe (
%eline%
echo [%SysPath%\sppsvc.exe] file is missing, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

for /f "delims=" %%a in ('%psc% "[System.Environment]::Version.Major" %nul6%') do if "%%a"=="2" (
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install %nul2% | find /i "0x1" %nul1% || (
%eline%
echo .NET 3.5 Framework is corrupt or missing. Aborting...
if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-Embedded" (
echo Install .NET Framework 4.8 and Windows Management Framework 5.1
)
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)
)

if %winbuild% LSS 9200 if exist "%SysPath%\wlms\wlms.exe" (
sc query wlms | find /i "RUNNING" %nul% && (
sc stop sppsvc %nul%
if !errorlevel! EQU 1051 (
%eline%
echo Evaluation WLMS service is running, sppsvc service can not be stopped. Aborting...
echo Install Non-Eval version for Windows build %winbuild%.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)
)
)

::========================================================================================================================================

if %_actwinesuoff%==1 (set "_actwin=1" & set "_actesu=1" & set "_actoff=1")
if %_actprojvis%==1   (set "_actoff=1")

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_ckeckwmic
call :dk_checksku
call :dk_product
call :dk_sppissue

::========================================================================================================================================

set error=

cls
echo:
call :dk_showosinfo

echo Initiating Diagnostic Tests...

set "_serv=sppsvc Winmgmt"

::  Software Protection
::  Windows Management Instrumentation

call :dk_errorcheck

if defined error (
call :dk_color %Red% "Some errors were detected. Aborting the operation..."
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto :dk_done
)

call :ts_getedition
if not defined tsedition (
call :dk_color %Red% "Checking Windows Edition ID             [Not found in installed licenses, aborting...]"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto :dk_done
)

::========================================================================================================================================

if %_resall%==1 goto :ts_resetall
if %_actman%==1 goto :ts_actman
if %_actappx%==1 goto :ts_appxlob
if %_actwinhost%==1 goto :ts_whost
if %_actoffhost%==1 goto :ts_ohost
if not %_actwin%==1 goto :ts_esu

::========================================================================================================================================

::  Process Windows
::  Check if system is permanently activated or not

echo:
echo Processing Windows...

echo %tsedition% | find /i "Eval" %nul1% && (
goto :ts_wineval
)

call :ts_checkwinperm
if defined _perm (
call :dk_color %Gray% "Checking OS Activation                  [Windows is already permanently activated]"
goto :ts_esu
)

set tempid=
set keytype=zero
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':wintsid\:.*';iex ($f[1])" %nul6%') do (
echo "%%a" | findstr /r ".*-.*-.*-.*-.*" %nul1% && (set tsids=!tsids! %%a& set tempid=%%a)
)

if defined tempid (
echo Checking Activation ID                  [%tempid%] [%tsedition%]
) else (
call :dk_color %Red% "Checking Activation ID                  [Not Found] [%tsedition%] [%osSKU%]"
set error=1
goto :ts_esu
)

if defined winsub (
call :dk_color %Blue% "Windows Subscription [SKU ID-%slcSKU%] found. Script will activate base edition [SKU ID-%regSKU%]."
echo:
)
goto :ts_esu

::========================================================================================================================================

:ts_wineval

call :dk_color %Gray% "Checking OS Edition                     [%tsedition%] [Evaluation edition found]"
call :dk_color %Blue% "Evaluation editions cannot be activated outside of evaluation period."

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
call :dk_color %Blue% "Script will reset evaluation period, but to permanently activate Windows,"
call :dk_color %Blue% "Go back to main menu and use [Change Edition] option and change to Non-eval edition."
) else (
call :dk_color %Blue% "Script will reset evaluation period, but to permanently activate Windows, install Non-eval edition."
call :dk_color %_Yellow% "%mas%evaluation_editions"
)

::  Check Internet connection

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not "%%#"=="" set _int=1)
)

if not defined _int (
%psc% "If([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet){Exit 0}Else{Exit 1}"
if !errorlevel!==0 (set _int=1&set ping_f= But Ping Failed)
)

if defined _int (
echo Checking Internet Connection            [Connected%ping_f%]
) else (
set error=1
call :dk_color %Red% "Checking Internet Connection            [Not Connected]"
call :dk_color %Blue% "Internet is required for Windows Evaluation activation."
)

::  List of products lacking activable evaluation keys and ISOs

::  c4b908d2-c4b9-439d-8ff0-48b656a24da4_EmbeddedIndustryEEval_8.1
::  9b74255b-afe1-4da7-a143-98d1874b2a6c_EnterpriseNEval_8
::  7fd0a88b-fb89-415f-9b79-84adc6a7cd56_EnterpriseNEval_8.1
::  994578eb-193c-4c99-bea0-2483274c9afd_EnterpriseSNEval_2015
::  b9f3109c-bfa9-4f37-9824-6dba9ee62056_ServerStorageStandardEval_2012R2
::  2d3b7269-65f4-467d-9d51-dbe0e5a4e668_ServerStorageWorkgroupEval_2012R2

:: --------

::  1st column = Activation ID
::  2nd column = Activable evaluation key
::  3rd column = Edition ID
::  4th column = Windows version (for reference only)
::  5th column = NoAct = activation is not working
::  Separator  = _

set f=
set key=
set eval=
if not defined allapps call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f

for %%# in (
d9eea459-1e6b-499d-8486-e68163f2a8be_N3QJR-YCWKK-RVJGK-GQFMX-T8%f%2BF_EmbeddedIndustryEval_8.1
fbd4c5c6-adc6-4740-bc65-b2dc6dc249c1_MJ8TN-42JH8-886MT-8THCF-36%f%67B_EnterpriseEval_8_NoAct_          REM New time based activation not available
0eebbb45-29d4-49cb-ba87-a23db0cce40a_76FKW-8NR3K-QDH4P-3C87F-JH%f%TTW_EnterpriseEval_8.1
3f4c0546-36c6-46a8-a37f-be13cdd0cf25_7HBDQ-QNKVG-K4RBF-HMBY6-YG%f%9R6_EnterpriseEval_10
1f8dbfe8-defa-4676-b5a6-f76949a01540_4N8VT-7Y686-43DGV-THTW9-M9%f%8W7_EnterpriseNEval_10
57a4ebb6-8e0c-41f8-b79e-8872ddc971ef_W63GF-7N4D9-GQH3K-K4FP7-9B%f%T6C_EnterpriseSEval_2015
b47dd250-fd6a-44c8-9217-03aca6e4812e_N4DMT-RJKDQ-XR6H7-3DKKP-3Y%f%JWT_EnterpriseSEval_2016
267bf82d-08e8-4046-b061-9ef3f8ac2b5a_N7HMH-MK36Q-M4X93-76KQ2-6J%f%HWR_EnterpriseSEval_2019
aff25f1f-fb53-4e27-95ef-b8e5aca10ac6_9V4NK-624Y3-VK47R-Q27GP-27%f%PGF_EnterpriseSEval_2021
399f0697-886b-4881-894c-4ff6c52e7d8f_CYPB3-XNV9V-QR4G4-Q3B8K-KQ%f%FGJ_EnterpriseSEval_2024
6162e8c2-3c30-46e1-b964-0de603498e2d_R34N9-HJ6Q3-GBX4F-Q24KQ-49%f%DF7_EnterpriseSNEval_2016
aed14fc8-907d-44fb-a3a1-d5d8e638acb3_MHN9Q-RD9PW-BFHDQ-9FTWQ-WQ%f%PF8_EnterpriseSNEval_2019
5dd0c869-eae9-40ce-af48-736692cd8e43_XCN62-29X92-C4T8X-WP82X-DY%f%MJ8_EnterpriseSNEval_2021
522cc0dc-3c7b-4258-ae68-f297ca63b64e_Y8DJM-NPXF3-QG4MH-W7WJK-KQ%f%FGM_EnterpriseSNEval_2024
aa708397-8618-42de-b120-a44190ef456d_R63DV-9NPDX-QVWJF-HMR8V-M4%f%K7D_IoTEnterpriseSEval_2024
cd25b1e8-5839-4a96-a769-b6abe3aa5dee_73BMN-332G9-DX6B8-FGDT3-GF%f%YT6_ServerDatacenterEval_2012
e628c5e8-2300-4429-8b80-a8b21bd7ce0a_WPR94-KN3J7-MRB7X-JPJV8-RX%f%7J2_ServerDatacenterEval_2012R2
01398239-85ff-487f-9e90-0e3cc5bcc92e_QVTQ9-GNRBH-JQ9G7-W7FBW-RX%f%9QR_ServerDatacenterEval_2016
5ea4af9e-fd59-4691-b61c-1fc1ff3e309e_KNW3G-22YD2-7QKQJ-2RF2X-H6%f%F8M_ServerDatacenterEval_2019
1d02774d-66ab-4c57-8b14-e254fdce09d4_PK7JN-24236-FH7JP-V792F-37%f%CYR_ServerDatacenterEval_2021
96794a98-097f-42fe-8f28-2c38ea115229_M4RNW-CRTHF-TY7BG-DDHG6-J2%f%T92_ServerDatacenterEval_2025
38d172c7-36b3-4e4b-b435-fd0b06b95c6e_RNFGD-WFFQR-XQ8BG-K7QQK-GJ%f%CP9_ServerStandardEval_2012
4fc45a88-26b5-4cf9-9eef-769ee3f0a016_79M8M-N36BX-8YGJY-2G9KP-3Y%f%GPC_ServerStandardEval_2012R2
9dfa8ec0-7665-4b9d-b2cb-bfc2dc37c9f4_9PBKX-4NHGT-QWV4C-4JD94-TV%f%KQ6_ServerStandardEval_2016
7783a126-c108-4cf7-b59f-13c78c7a7337_J4WNC-H9BG3-6XRX4-3XD8K-Y7%f%XRX_ServerStandardEval_2019
c1a197b6-ba5e-4394-b9bf-b659a6c1b873_7PBJM-MNVPD-MBQD7-TYTY4-W8%f%JDY_ServerStandardEval_2021
753c53a2-4274-4339-8c2e-f66c0b9646c5_YPBVM-HFNWQ-CTF9M-FR4RR-7H%f%9YG_ServerStandardEval_2025
0de5ff31-2d62-4912-b1a8-3ea01d2461fd_3CKBN-3GJ8X-7YT4X-D8DDC-D6%f%69B_ServerStorageStandardEval_2012
fb08f53a-e597-40dc-9f08-8bbf99f19b92_NCJ6J-J23VR-DBYB3-QQBJF-W8%f%CP7_ServerStorageWorkgroupEval_2012
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do if %tsedition%==%%C if not defined key (
echo "%allapps%" | find /i "%%A" %nul1% && (
set key=%%B
set eval=1
if /i "%%E"=="NoAct" set noact=1
echo Checking Activation ID                  [%%A] [%%C]
)
)
)

if not defined key (
set error=1
call :dk_color %Red% "Checking Activation ID                  [%tsedition% not found in the script]"
call :dk_color %Blue% "Make sure you are using the updated version of the script."
goto :ts_esu
)

set resetstuff=1
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':tsforge\:.*';iex ($f[1])"
set resetstuff=
if !errorlevel!==3 (
set error=1
call :dk_color %Red% "Resetting Rearm / GracePeriod           [Failed]"
call :dk_color %Blue% "%_fixmsg%"
goto :ts_esu
) else (
echo Resetting Rearm / GracePeriod           [Successful]
)

%psc% "try { $null=(([WMISEARCHER]'SELECT Version FROM %sps%').Get()).InstallProductKey('%key%'); exit 0 } catch { exit $_.Exception.InnerException.HResult }" %nul%
set keyerror=%errorlevel%
cmd /c exit /b %keyerror%
if %keyerror% NEQ 0 set "keyerror=[0x%=ExitCode%]"

if %keyerror% EQU 0 (
call :dk_refresh
echo Installing Activable Evaluation Key     [%key%] [Successful]
) else (
set error=1
call :dk_color %Red% "Installing Activable Evaluation Key     [%key%] [Failed] %keyerror%"
call :dk_color %Blue% "%_fixmsg%"
)

::========================================================================================================================================

:ts_esu

if not %_actesu%==1 goto :ts_off

::  Process Windows ESU

echo:
echo Processing Windows ESU...

set esuexist=
set esuexistsup=
set esueditionlist=
set esuexistbutnosup=

for %%# in (EnterpriseS IoTEnterpriseS IoTEnterpriseSK) do (if /i %tsedition%==%%# set isltsc=1)
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" set isServer=1

if /i %tsedition%==Embedded (
if exist "%SystemRoot%\Servicing\Packages\WinEmb-Branding-Embedded-ThinPC-Package*.mum" set isThinpc=1
if exist "%SystemRoot%\Servicing\Packages\WinEmb-Branding-Embedded-POSReady7-Package*.mum" set subEdition=[POS]
if exist "%SystemRoot%\Servicing\Packages\WinEmb-Branding-Embedded-Standard-Package*.mum" set subEdition=[Standard]
)
if not defined allapps call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f

if not defined isThinpc if not defined isltsc for %%# in (
REM Windows7
4220f546-f522-46df-8202-4d07afd26454_Client-ESU-Year3[1-3y]_-Enterprise-EnterpriseE-EnterpriseN-Professional-ProfessionalE-ProfessionalN-Ultimate-UltimateE-UltimateN-
7e94be23-b161-4956-a682-146ab291774c_Client-ESU-Year6[4-6y]_-Enterprise-EnterpriseE-EnterpriseN-Professional-ProfessionalE-ProfessionalN-Ultimate-UltimateE-UltimateN-
REM Windows7EmbeddedPOSReady7
4f1f646c-1e66-4908-acc7-d1606229b29e_POS-ESU-Year3[1-3y]_-Embedded[POS]-
REM Windows7EmbeddedStandard
6aaf1c7d-527f-4ed5-b908-9fc039dfc654_WES-ESU-Year3[1-3y]_-Embedded[Standard]-
REM WindowsServer2008R2
8e7bfb1e-acc1-4f56-abae-b80fce56cd4b_Server-ESU-PA[1-6y]_-ServerDatacenter-ServerDatacenterCore-ServerDatacenterV-ServerDatacenterVCore-ServerStandard-ServerStandardCore-ServerStandardV-ServerStandardVCore-ServerEnterprise-ServerEnterpriseCore-ServerEnterpriseV-ServerEnterpriseVCore-
REM Windows8.1
4afc620f-12a4-48ad-8015-2aebfbd6e47c_Client-ESU-Year3[1-3y]_-Enterprise-EnterpriseN-Professional-ProfessionalN-
11be7019-a309-4763-9a09-091d1722ffe3_Client-FES-ESU-Year3[1-3y]_-EmbeddedIndustry-EmbeddedIndustryE-
REM WindowsServer2012/2012R2
55b1dd2d-2209-4ea0-a805-06298bad25b3_Server-ESU-Year3[1-3y]_-ServerDatacenter-ServerDatacenterCore-ServerDatacenterV-ServerDatacenterVCore-ServerStandard-ServerStandardCore-ServerStandardV-ServerStandardVCore-
REM Windows10
83d49986-add3-41d7-ba33-87c7bfb5c0fb_Client-ESU-Year3[1-3y]_-Education-EducationN-Enterprise-EnterpriseN-Professional-ProfessionalEducation-ProfessionalEducationN-ProfessionalN-ProfessionalWorkstation-ProfessionalWorkstationN-
0b533b5e-08b6-44f9-b885-c2de291ba456_Client-ESU-Year6[4-6y]_-Education-EducationN-Enterprise-EnterpriseN-Professional-ProfessionalEducation-ProfessionalEducationN-ProfessionalN-ProfessionalWorkstation-ProfessionalWorkstationN-
4dac5a0c-5709-4595-a32c-14a56a4a6b31_Client-IoT-ESU-Year3[1-3y]_-IoTEnterprise- REM Removed IoTEnterpriseS because it already has longer support
f69e2d51-3bbd-4ddf-8da7-a145e9dca597_Client-IoT-ESU-Year6[4-6y]_-IoTEnterprise- REM Removed IoTEnterpriseS because it already has longer support
) do (
for /f "tokens=1-3 delims=_" %%A in ("%%#") do (
echo "%allapps%" | find /i "%%A" %nul1% && (
set esuexist=1
echo "%%C" | find /i "-%tsedition%%subEdition%-" %nul1% && (
set esuexistsup=1
set esueditionlist=
set esuexistbutnosup=
set tsids=!tsids! %%A
echo Checking Activation ID                  [%%A] [%%B]
) || (
if not defined esueditionlist set esueditionlist=%%C
set esuexistbutnosup=1
)
)
)
)

if defined esuexistsup (
echo "%tsids%" | find /i "4220f546-f522-46df-8202-4d07afd26454" %nul1% && (
echo "%tsids%" | find /i "7e94be23-b161-4956-a682-146ab291774c" %nul1% || (
call :dk_color %Gray% "Now update Windows to get Client-ESU-Year6[4-6y] license and activate that using this script."
)
)
goto :ts_off
)

if defined isltsc (
call :dk_color %Gray% "Checking Activation ID                  [%tsedition% LTSC already has longer support, ESU is not applicable]"
goto :ts_off
)

if defined esuexistbutnosup (
call :dk_color %Red% "Checking Activation ID                  [Commercial ESU is not supported for %tsedition%]"
call :dk_color %Blue% "Go back to Main Menu, select Change Windows Edition option and change to any of the below listed editions."
echo [%esueditionlist%]
goto :ts_off
)

set esuavail=
if %winbuild% LEQ 7602 if not defined isThinpc set esuavail=1
if %winbuild% GTR 7602 if %winbuild% LSS 10240 if defined isServer set esuavail=1
if %winbuild% GEQ 10240 if %winbuild% LEQ 19045 if not defined isServer set esuavail=1
if %winbuild% EQU 9600 set esuavail=1

if defined esuavail (
call :dk_color %Red% "Checking Activation ID                  [ESU license is not found, make sure Windows is fully updated]"
set fixes=%fixes% %mas%tsforge#windows-esu
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%tsforge#windows-esu"
) else (
call :dk_color %Gray% "Checking Activation ID                  [ESU is not available for %winos%]"
)

::========================================================================================================================================

:ts_off

if not %_actoff%==1 goto :ts_act

if %winbuild% LSS 9200 (
echo:
call :dk_color %Gray% "Checking Supported Office               [TSforge for Office is supported on Windows 8 and later versions]"
call :dk_color %Blue% "On Windows 7 build, use Online %KS% activation option for Office instead."
goto :ts_act
)

::  Check ohook install

set ohook=
for %%# in (15 16) do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" set ohook=1
)
)

for %%# in (System SystemX86) do (
for %%G in ("Office 15" "Office") do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" set ohook=1
)
)
)

if defined ohook (
echo:
call :dk_color %Gray% "Checking Ohook                          [Ohook activation is already installed for Office]"
)

::  Check unsupported office versions

set o14msi=
set o14c2r=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o14msi=Office 2010 MSI )
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o14msi=Office 2010 MSI )
%nul% reg query %_68%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 
%nul% reg query %_86%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 

if not "%o14msi%%o14c2r%"=="" (
echo:
call :dk_color %Red% "Checking Unsupported Office Install     [ %o14msi%%o14c2r%]"
)

if %winbuild% GEQ 10240 %psc% "Get-AppxPackage -name "Microsoft.MicrosoftOfficeHub"" | find /i "Office" %nul1% && (
set ohub=1
)

::========================================================================================================================================

::  Check supported office versions

call :ts_getpath

set o16uwp=
set o16uwp_path=

if %winbuild% GEQ 10240 (
for /f "delims=" %%a in ('%psc% "(Get-AppxPackage -name 'Microsoft.Office.Desktop' | Select-Object -ExpandProperty InstallLocation)" %nul6%') do (if exist "%%a\Integration\Integrator.exe" (set o16uwp=1&set "o16uwp_path=%%a"))
)

sc query ClickToRunSvc %nul%
set error1=%errorlevel%

if defined o16c2r if %error1% EQU 1060 (
echo:
call :dk_color %Red% "Checking ClickToRun Service             [Not found, Office 16.0 files found]"
set o16c2r=
set error=1
)

sc query OfficeSvc %nul%
set error2=%errorlevel%

if defined o15c2r if %error1% EQU 1060 if %error2% EQU 1060 (
echo:
call :dk_color %Red% "Checking ClickToRun Service             [Not found, Office 15.0 files found]"
set o15c2r=
set error=1
)

if "%o16uwp%%o16c2r%%o15c2r%%o16msi%%o15msi%"=="" (
set error=1
set showfix=1
echo:
if not "%o14msi%%o14c2r%"=="" (
call :dk_color %Red% "Checking Supported Office Install       [Not Found]"
) else (
if %_actwin%==0 (
call :dk_color %Red% "Checking Installed Office               [Not Found]"
) else (
call :dk_color %Gray% "Checking Installed Office               [Not Found]"
)
)

if defined ohub (
echo:
echo You have only Office dashboard app installed, you need to install full Office version.
)
call :dk_color %Blue% "Download and install Office from below URL and try again."
if %_actwin%==0 set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto :ts_act
)

set multioffice=
if not "%o16uwp%%o16c2r%%o15c2r%%o16msi%%o15msi%"=="1" set multioffice=1
if not "%o14c2r%%o14msi%"=="" set multioffice=1

if defined multioffice (
echo:
call :dk_color %Gray% "Checking Multiple Office Install        [Found. Recommended to install one version only]"
)

::========================================================================================================================================

::  Check Windows Server

set winserver=
reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v ProductType %nul2% | find /i "WinNT" %nul1% || set winserver=1
if not defined winserver (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Server" %nul1% && set winserver=1
)

::========================================================================================================================================

::  Process Office UWP

if not defined o16uwp goto :ts_starto15c2r

call :ts_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=16
set "_oLPath=%o16uwp_path%\Licenses16"
set "pkeypath=%o16uwp_path%\Office16\pkeyconfig-office.xrm-ms"
for /f "delims=" %%a in ('%psc% "(Get-AppxPackage -name 'Microsoft.Office.Desktop' | Select-Object -ExpandProperty Dependencies) | Select-Object PackageFullName" %nul6%') do (set "o16uwpapplist=!o16uwpapplist! %%a")

echo "%o16uwpapplist%" | findstr /i "Access Excel OneNote Outlook PowerPoint Publisher SkypeForBusiness Word" %nul% && set "_oIds=O365HomePremRetail"

for %%# in (Project Visio) do (
echo "%o16uwpapplist%" | findstr /i "%%#" %nul% && (
set _lat=
if exist "%_oLPath%\%%#Pro2024VL*.xrm-ms" set "_oIds= !_oIds! %%#Pro2024Retail " & set _lat=1
if not defined _lat if exist "%_oLPath%\%%#Pro2021VL*.xrm-ms" set "_oIds= !_oIds! %%#Pro2021Retail " & set _lat=1
if not defined _lat if exist "%_oLPath%\%%#Pro2019VL*.xrm-ms" set "_oIds= !_oIds! %%#Pro2019Retail " & set _lat=1
if not defined _lat set "_oIds= !_oIds! %%#ProRetail "
)
)

set uwpinfo=%o16uwp_path:C:\Program Files\WindowsApps\Microsoft.Office.Desktop_=%

echo:
echo Processing Office...                    [UWP ^| %uwpinfo%]

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :ts_starto15c2r
)

call :ts_process

::========================================================================================================================================

:ts_starto15c2r

::  Process Office 15.0 C2R

if not defined o15c2r goto :ts_starto16c2r

call :ts_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=15
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v ProductReleaseIds" %nul6%') do (set "_prids=%o15c2r_reg%\Configuration /v ProductReleaseIds" & set "_config=%o15c2r_reg%\Configuration")
if not defined _oArch   for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v Platform" %nul6%') do (set "_oArch=%%b")
if not defined _version for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v version" %nul6%') do (set "_version=%%b")
if not defined _prids   for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v ProductReleaseId" %nul6%') do (set "_prids=%o15c2r_reg%\propertyBag /v ProductReleaseId" & set "_config=%o15c2r_reg%\propertyBag")

echo "%o15c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=10) || (set _tok=9)
for /f "tokens=%_tok% delims=\" %%a in ('reg query %o15c2r_reg%\ProductReleaseIDs\Active %nul6% ^| findstr /i "Retail Volume"') do (
echo "!_oIds!" | find /i " %%a " %nul1% || (set "_oIds= !_oIds! %%a ")
)

set "_oLPath=%_oRoot%\Licenses"
set "pkeypath=%_oRoot%\Office15\pkeyconfig-office.xrm-ms"
set "_oIntegrator=%_oRoot%\integration\integrator.exe"

echo:
echo Processing Office...                    [C2R ^| %_version% ^| %_oArch%]

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :ts_starto16c2r
)

if "%_actprojvis%"=="0" call :oh_fixprids
call :ts_process

::========================================================================================================================================

:ts_starto16c2r

::  Process Office 16.0 C2R

if not defined o16c2r goto :ts_startmsi

call :ts_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=16
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v AudienceData" %nul6%') do (set "_AudienceData=^| %%b ")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ProductReleaseIds" %nul6%') do (set "_prids=%o16c2r_reg%\Configuration /v ProductReleaseIds" & set "_config=%o16c2r_reg%\Configuration")

echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=9) || (set _tok=8)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs" /s /f ".16" /k %nul6% ^| findstr /i "Retail Volume"') do (
echo "!_oIds!" | find /i " %%a " %nul1% || (set "_oIds= !_oIds! %%a ")
)
set _oIds=%_oIds:.16=%
set _o16c2rIds=%_oIds%

set "_oLPath=%_oRoot%\Licenses16"
set "pkeypath=%_oRoot%\Office16\pkeyconfig-office.xrm-ms"
set "_oIntegrator=%_oRoot%\integration\integrator.exe"

echo:
echo Processing Office...                    [C2R ^| %_version% %_AudienceData%^| %_oArch%]

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :ts_startmsi
)

if "%_actprojvis%"=="0" call :oh_fixprids
call :ts_process

::========================================================================================================================================

:ts_startmsi

if defined o15msi call :ts_processmsi 15 %o15msi_reg%
if defined o16msi call :ts_processmsi 16 %o16msi_reg%

::========================================================================================================================================

echo:
call :oh_clearblock
if "%o16msi%%o15msi%"=="" if not "%o16uwp%%o16c2r%%o15c2r%"=="" call :oh_uninstkey
call :oh_licrefresh

goto :ts_act

::========================================================================================================================================

:ts_whost

::  Process Windows K-M-S host

echo:
echo Processing Windows %KS% Host...

echo:
if %winbuild% GEQ 10586 (
call :dk_color %Gray% "With %KS% Host license, system may randomly change Windows Edition later. It is a Windows issue and can be safely ignored."
)
call :dk_color %Gray% "%KS% Host [Not to be confused with %KS% Client] license causes the sppsvc service to run continuously."
call :dk_color %Blue% "Only use this activation when necessary, you can revert to normal activation from the previous menu."

if %_unattended%==0 (
echo:
choice /C:0F /N /M "> [0] Go back  [F] Continue : "
if !errorlevel!==1 exit /b
echo:
)

set _arr=
set tempid=
set keytype=kmshost

::  Install current edition csvlk license so that correct edition can reflect for csvlk

if %winbuild% GEQ 10586 (
for %%# in ("%SysPath%\spp\tokens\skus\%tsedition%\*CSVLK*.xrm-ms") do (
if defined _arr (set "_arr=!_arr!;"%SysPath%\spp\tokens\skus\%tsedition%\%%~nx#"") else (set "_arr="%SysPath%\spp\tokens\skus\%tsedition%\%%~nx#"")
)
if defined _arr %psc% "$sls = Get-WmiObject %sps%; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); InstallLicenseArr '!_arr!'" %nul%
)

for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':wintsid\:.*';iex ($f[1])" %nul6%') do (
echo "%%a" | findstr /r ".*-.*-.*-.*-.*" %nul1% && (set tsids=!tsids! %%a& set tempid=%%a)
)

if defined tempid (
echo Checking Activation ID                  [%tempid%] [%tsedition%]
) else (
call :dk_color %Red% "Checking Activation ID                  [Not Found] [%tsedition%] [%osSKU%]"
call :dk_color %Blue% "%KS% Host license is not found on your system. It is available for the below editions."
call :dk_color %Blue% "Professional, Education, ProfessionalWorkstation, Enterprise, EnterpriseS, and Server editions, etc."
goto :ts_act
)

if defined winsub (
echo:
call :dk_color %Blue% "Windows Subscription [SKU ID-%slcSKU%] found. Script will activate base edition [SKU ID-%regSKU%]."
)

goto :ts_act

::========================================================================================================================================

:ts_ohost

::  Process Office K-M-S host

echo:
echo Processing Office %KS% Host...

set ohostexist=
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663
set ohostids=%allapps%
call :dk_actids 59a52881-a989-479d-af46-f275c6370663
set ohostids=%ohostids% %allapps%

for %%# in (
bfe7a195-4f8f-4f0b-a622-cf13c7d16864_KMSHost2010-ProPlusVL
f3d89bbf-c0ec-47ce-a8fa-e5a5f97e447f_KMSHost2024Volume
47f3b983-7c53-4d45-abc6-bcd91e2dd90a_KMSHost2021Volume
70512334-47b4-44db-a233-be5ea33b914c_KMSHost2019Volume
98ebfe73-2084-4c97-932c-c0cd1643bea7_KMSHost2016Volume
2e28138a-847f-42bc-9752-61b03fff33cd_KMSHost2013Volume
) do (
for /f "tokens=1-2 delims=_" %%A in ("%%#") do (
echo "%ohostids%" | find /i "%%A" %nul1% && (
set ohostexist=1
set tsids=!tsids! %%A
echo Checking Activation ID                  [%%A] [%%B]
)
)
)

if not defined ohostexist (
call :dk_color %Gray% "Checking Activation ID                  [Not found for Office %KS% Host]"
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%tsforge#office-kms-host"
)

echo:
call :dk_color %Gray% "%KS% Host [Not to be confused with %KS% Client] license causes the sppsvc service to run continuously."
call :dk_color %Gray% "Only use this activation when necessary."

goto :ts_act

::========================================================================================================================================

:ts_appxlob

::  Process Windows 8/8.1 APPX Sideloading

echo:
echo Processing Windows 8/8.1 APPX Sideloading...

if %winbuild% LSS 9200 set noappx=1
if %winbuild% GTR 9600 set noappx=1

echo:
if defined noappx (
call :dk_color %Gray% "Checking Activation ID                  [APPX Sideloading feature is available only on Windows 8/8.1]"
goto :dk_done
)

set appxexist=
if not defined allapps call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f

for %%# in (
ec67814b-30e6-4a50-bf7b-d55daf729d1e_APPXLOB-Client
251ef9bf-2005-442f-94c4-86307de7bb32_APPXLOB-Embedded-Industry
1e58c9d7-e3f1-4f69-9039-1f162463ac2c_APPXLOB-Embedded-Standard
3502d53e-5d43-436a-84af-714e8d334f8d_APPXLOB-Server
) do (
for /f "tokens=1-2 delims=_" %%A in ("%%#") do (
echo "%allapps%" | find /i "%%A" %nul1% && (
set appxexist=1
set tsids=!tsids! %%A
echo Checking Activation ID                  [%%A] [%%B]
)
)
)

if not defined appxexist (
call :dk_color %Red% "Checking Activation ID                  [Not found]"
call :dk_color %Blue% "APPX Sideloading feature is available only on Pro and higher level editions."
)

goto :ts_act

::========================================================================================================================================

:ts_resetall

echo:
echo Processing Reset of Rearm / Timers / Tamper / Lock...
echo:

set resetstuff=1
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':tsforge\:.*';iex ($f[1])"

if %errorlevel%==3 (
call :dk_color %Red% "Reset Failed."
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
) else (
call :dk_color %Green% "Reset process has been successfully done."
)

goto :dk_done

::========================================================================================================================================

:ts_actman

echo:
echo Processing Manual Activation...
echo:

call :dk_color %Gray% "This option is for advanced users, those who already know what they are doing."
call :dk_color %Blue% "Some activation IDs may cause system crash [MUI mismatch], or irreversible changes [CloudEdition etc]."

if %_unattended%==1 (
echo:
for %%# in (%tsids%) do (echo Activation ID - %%#)
goto :ts_act
)

call :dk_color %Blue% "Although the script will try to remove those IDs from the list, it is not fully guaranteed."
echo:
choice /C:0F /N /M "> [0] Go back  [F] Continue : "
if %errorlevel%==1 exit /b

echo:
echo Fetching Supported Activation IDs list. Please wait...

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':listactids\:.*';iex ($f[1])"
if %errorlevel%==3 (
call :dk_color %Gray% "No supported activation ID found, aborting..."
goto :dk_done
)

for /f "delims=" %%a in ('%psc% "$ids = Get-WmiObject -Query 'SELECT ID FROM SoftwareLicensingProduct' | Select-Object -ExpandProperty ID; $ids" %nul6%') do call set "allactids= %%a !allactids! "

echo:
call :dk_color %Gray% "Enter / Paste the Activation ID shown in first column in the opened text file, or just press Enter to return:"
echo Add space after each Activation ID if you are adding multiple:
echo:
set /p tsids=

del /f /q "%SystemRoot%\Temp\actids_159_*" %nul%
if not defined tsids goto :dk_done

for %%# in (%tsids%) do (
echo "%allactids%" | find /i " %%# " %nul1% || (
call :dk_color %Red% "[%%#] Incorrect Activation ID entered, aborting..."
goto :dk_done
)
)

goto :ts_act

:listactids:
$t = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0)
$t.DefinePInvokeMethod('SLOpen', 'slc.dll', 22, 1, [Int32], @([IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
$t.DefinePInvokeMethod('SLClose', 'slc.dll', 22, 1, [IntPtr], @([IntPtr]), 1, 3).SetImplementationFlags(128)
$t.DefinePInvokeMethod('SLGetProductSkuInformation', 'slc.dll', 22, 1, [Int32], @([IntPtr], [Guid].MakeByRefType(), [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
$t.DefinePInvokeMethod('SLGetLicense', 'slc.dll', 22, 1, [Int32], @([IntPtr], [Guid].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
$w = $t.CreateType()
$m = [Runtime.InteropServices.Marshal]

function slGetSkuInfo($SkuId) {
    $c = 0; $b = 0
    $r = $w::SLGetProductSkuInformation($hSLC, [ref][Guid]$SkuId, "msft:sl/EUL/PHONE/PUBLIC", [ref]$null, [ref]$c, [ref]$b)
    return ($r -eq 0)
}

function IsMuiNotLocked($SkuId) {
    $r = $true; $c = 0; $b = 0

    $LicId = [Guid]::Empty
    [void]$w::SLGetProductSkuInformation($hSLC, [ref][Guid]$SkuId, "fileId", [ref]$null, [ref]$c, [ref]$b)
    $FileId = $m::PtrToStringUni($b)

    $c = 0; $b = 0
    [void]$w::SLGetLicense($hSLC, [ref][Guid]$FileId, [ref]$c, [ref]$b)
    $blob = New-Object byte[] $c; $m::Copy($b, $blob, 0, $c)
    $cont = [Text.Encoding]::UTF8.GetString($blob)
    $xml = [xml]$cont.SubString($cont.IndexOf('<r'))

    $xml.licenseGroup.license[0].grant | foreach {
        $_.allConditions.allConditions.productPolicies.policyStr | where { $_.name -eq 'Kernel-MUI-Language-Allowed' } | foreach {
            if ($_.InnerText -ne 'EMPTY') { $r = $false }
        }
    }
    return $r
}

$hSLC = 0; [void]$w::SLOpen([ref]$hSLC)
$results = Get-WmiObject -Query "SELECT ID, Name, Description FROM SoftwareLicensingProduct"
$maxNameWidth = 60

$filteredResults = $results | Where-Object {
    if ($env:tsedition -like "*CountrySpecific*") {
        $true
    }
    else {
        $_.Name -notlike "*CountrySpecific*"
    }
} | Where-Object {
    if ($env:tsedition -like "*CloudEdition*") {
        $true
    }
    else {
        $_.Name -notlike "*CloudEdition*"
    }
} | Where-Object {
    $_.Name -like "*CountrySpecific*" -or (IsMuiNotLocked $_.ID)
} | Where-Object {
    slGetSkuInfo $_.ID
} | ForEach-Object {
    "$($_.ID)`t$($_.Name.PadRight($maxNameWidth))`t$($_.Description)"
}

[void]$w::SLClose($hSLC)
if (-not $filteredResults) {
    Exit 3
}

$sortedResults = $filteredResults | Sort-Object { $_.Split("`t")[1].Trim() }
$output = $sortedResults -join "`r`n"
$newGuid = [Guid]::NewGuid().Guid
$filename = "$env:SystemRoot\Temp\actids_159_$newGuid.txt"
$output | Set-Content -Path $filename -Encoding ASCII
Start-Process notepad.exe $filename
:listactids:

::========================================================================================================================================

:ts_act

if defined eval (
echo:
echo Activating...
echo:
call :dk_act

set gpr=0
set gprdays=0
set actdone=
for /f "delims=" %%a in ('%psc% "(Get-WmiObject -Query 'SELECT GracePeriodRemaining FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').GracePeriodRemaining" %nul6%') do set "gpr=%%a"
set /a "gprdays=(!gpr!+1440-1)/1440"

if !gprdays! EQU 90 set actdone=1
if !gprdays! EQU 180 set actdone=1

if defined actdone (
call :dk_color %Green% "[%winos%] has been reset and activated successfully for !gprdays! days."
) else (
set error=1
set showfix=1
call :dk_color %Red% "[%winos%] Activation Failed %error_code%. Remaining Period: !gprdays! days [!gpr! minutes]."
if not defined noact (
call :dk_color %Gray% "To activate, check your internet connection and ensure the date and time are correct."
) else (
call :dk_color %Blue% "This Windows version is known to not activate due to MS Windows/Server issues."
)
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)
)

if defined tsids (
echo:
echo Installing Forged Product Key Data...
echo Depositing Zero Confirmation ID...
echo:
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':tsforge\:.*';& ([ScriptBlock]::Create($f[1])) %tsids%"
if !errorlevel!==3 (
if %_actman%==0 call :dk_color %Blue% "%_fixmsg%"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
) else (
echo "%tsids%" | find /i "7e94be23-b161-4956-a682-146ab291774c" %nul1% && (
call :dk_color %Gray% "Windows Update can receive 1-3 years of ESU. 4-6 year ESU is not officially supported, but you can manually install updates."
)
echo "%tsids%" | findstr /i "4afc620f-12a4-48ad-8015-2aebfbd6e47c 11be7019-a309-4763-9a09-091d1722ffe3" %nul1% && (
call :dk_color %Gray% "ESU is not officially supported on Windows 8.1, but you can manually install updates until Jan-2024."
)
echo "%tsids%" | findstr /i "0b533b5e-08b6-44f9-b885-c2de291ba456 f69e2d51-3bbd-4ddf-8da7-a145e9dca597" %nul1% && (
call :dk_color %Gray% "Windows Update can receive 1-3 years of ESU. 4-6 year ESU license is added just as a placeholder."
)
)

if %_actwin%==1 for %%# in (407) do if %osSKU%==%%# (
call :dk_color %Red% "%winos% does not support activation on non-azure platforms."
)

if %_actoff%==1 if not defined error if defined ohub (
echo:
call :dk_color %Gray% "Office apps such as Word, Excel are activated, use them directly. Ignore 'Buy' button in Office dashboard app."
)

REM Trigger reevaluation of SPP's Scheduled Tasks
call :dk_reeval %nul%
)

if not defined tsids if defined error if not defined showfix (
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

goto :dk_done

::========================================================================================================================================

:ts_remove

cls
if not defined terminal (
mode 100, 30
)
title  Remove TSforge Activation %masver%

echo:
echo TSforge activation doesn't modify any Windows component.
echo TSforge activation doesn't install any new file in the system.
echo:
echo Instead, it appends data to one of data files used by Software Protection Platform.
echo:
call :dk_color %Gray% "If you want to reset the activation status,"
call :dk_color %Blue% "%_fixmsg%"
echo:

goto :dk_done

::========================================================================================================================================

:ts_reset

set key=
set _oRoot=
set _oArch=
set _oIds=
set _oLPath=
set _actid=
set _prod=
set _lic=
set _arr=
set _prids=
set _config=
set _version=
set _License=
set _oBranding=
exit /b

::========================================================================================================================================

:ts_getpath

set o16c2r=
set o15c2r=
set o16msi=
set o15msi=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set o16c2r=1&set o16c2r_reg=%_86%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set o16c2r=1&set o16c2r_reg=%_68%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set o15c2r=1&set o15c2r_reg=%_86%\15.0\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set o15c2r=1&set o15c2r_reg=%_68%\15.0\ClickToRun)

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o16msi=1&set o16msi_reg=%_86%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o16msi=1&set o16msi_reg=%_68%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o15msi=1&set o15msi_reg=%_86%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o15msi=1&set o15msi_reg=%_68%\15.0)

exit /b

::========================================================================================================================================

:ts_process

if not exist "%pkeypath%" (
call :dk_color %Red% "Checking pkeyconfig-office.xrm-ms       [Not found. Aborting activation...]"
set error=1
exit /b
)

for %%# in (%_oIds%) do (
set _actid=
set _preview=
set _License=%%#

set skipprocess=
if "%_actprojvis%"=="1" (
echo %%# | findstr /i "Project Visio" %nul% || (
set skipprocess=1
call :dk_color %Gray% "Skipping Because Project/Visio Mode     [%%#]"
)
)

if not defined skipprocess (

echo %%# | findstr /i "O365" %nul% && (
set _License=MondoRetail
set _altoffid=MondoRetail
call :ks_osppready
echo Converting Unsupported O365 Office      [%%# To MondoRetail]
)

set keytype=zero
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':offtsid\:.*';iex ($f[1])" %nul6%') do (
echo "%%a" | findstr /r ".*-.*-.*-.*-.*" %nul1% && (set tsids=!tsids! %%a& set _actid=%%a)
)
set "_allactid=!tsids!"

if defined _actid (
echo Checking Activation ID                  [!_actid!] [!_License!]
) else (
call :dk_color %Red% "Checking Activation ID                  [Office %oVer%.0 !_License! not found]"
set error=1
set showfix=1
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

echo %%# | find /i "2024" %nul% && (
if exist "!_oLPath!\ProPlus2024PreviewVL_*.xrm-ms" if not exist "!_oLPath!\ProPlus2024VL_*.xrm-ms" set _preview=1
)

if defined _actid (
echo "!allapps!" | find /i "!_actid!" %nul1% || call :oh_installlic
)
)
)

::  Add SharedComputerLicensing registry key if Retail Office C2R is installed on Windows Server
::  https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/click-to-run-office-on-terminal-server

if defined winserver if defined _config (
echo %_oIds% | find /i "Retail" %nul1% && (
set scaIsNeeded=1
reg add %_config% /v SharedComputerLicensing /t REG_SZ /d "1" /f %nul1%
echo Adding SharedComputerLicensing Reg      [Successful] [Needed on Server With Retail Office]"
)
)

exit /b

::========================================================================================================================================

:ts_processmsi

::  Process Office MSI Version

call :ts_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=%1
for /f "skip=2 tokens=2*" %%a in ('"reg query %2\Common\InstallRoot /v Path" %nul6%') do (set "_oRoot=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %2\Common\ProductVersion /v LastProduct" %nul6%') do (set "_version=%%b")
if "%_oRoot:~-1%"=="\" set "_oRoot=%_oRoot:~0,-1%"

echo "%2" | find /i "Wow6432Node" %nul1% && set _oArch=x86
if not "%osarch%"=="x86" if not defined _oArch set _oArch=x64
if "%osarch%"=="x86" set _oArch=x86

set "_common=%CommonProgramFiles%"
if defined PROCESSOR_ARCHITEW6432 set "_common=%CommonProgramW6432%"
set "_common2=%CommonProgramFiles(x86)%"

for /r "%_common%\Microsoft Shared\OFFICE%oVer%\" %%f in (BRANDING.XML) do if exist "%%f" set "_oBranding=%%f"
if not defined _oBranding for /r "%_common2%\Microsoft Shared\OFFICE%oVer%\" %%f in (BRANDING.XML) do if exist "%%f" set "_oBranding=%%f"

if exist "%_common%\Microsoft Shared\OFFICE%oVer%\Office Setup Controller\pkeyconfig-office.xrm-ms" (
set "pkeypath=%_common%\Microsoft Shared\OFFICE%oVer%\Office Setup Controller\pkeyconfig-office.xrm-ms"
) else if exist "%_common2%\Microsoft Shared\OFFICE%oVer%\Office Setup Controller\pkeyconfig-office.xrm-ms" (
set "pkeypath=%_common2%\Microsoft Shared\OFFICE%oVer%\Office Setup Controller\pkeyconfig-office.xrm-ms"
)

call :ts_msiofficedata %2

echo:
echo Processing Office...                    [MSI ^| %_version% ^| %_oArch%]

if not defined _oBranding (
set error=1
call :dk_color %Red% "Checking BRANDING.XML                   [Not Found. Aborting activation...]"
exit /b
)

if not defined _oIds (
set error=1
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
exit /b
)

call :ts_process
exit /b

::========================================================================================================================================

::  Get Windows permanent activation status (not counting csvlk)

:ts_checkwinperm

%psc% "Get-WmiObject -Query 'SELECT Name, Description FROM SoftwareLicensingProduct WHERE LicenseStatus=''1'' AND GracePeriodRemaining=''0'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn IS NULL' | Where-Object { $_.Description -notmatch 'KMS_' } | Select-Object -Property Name" %nul2% | findstr /i "Windows" %nul1% && set _perm=1||set _perm=
exit /b

::========================================================================================================================================

:tsforge:
$src = @'
// Common.cs
namespace LibTSforge
{
    using Microsoft.Win32;
    using System;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.ServiceProcess;
    using System.Text;
    using LibTSforge.Crypto;
    using LibTSforge.PhysicalStore;
    using LibTSforge.SPP;
    using LibTSforge.TokenStore;

    public enum PSVersion
    {
        Vista,
        Win7,
        Win8Early,
        Win8,
        WinBlue,
        WinModern
    }

    public static class Constants
    {
        public static readonly byte[] UniversalHWIDBlock =
        {
            0x26, 0x00, 0x00, 0x00, 0x00, 0x00, 0x1c, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x01, 0x0c, 0x01, 0x00
        };

        public static readonly byte[] KMSv4Response =
        {
            0x00, 0x00, 0x04, 0x00, 0x62, 0x00, 0x00, 0x00, 0x30, 0x00, 0x35, 0x00, 0x34, 0x00, 0x32, 0x00,
            0x36, 0x00, 0x2D, 0x00, 0x30, 0x00, 0x30, 0x00, 0x32, 0x00, 0x30, 0x00, 0x36, 0x00, 0x2D, 0x00,
            0x31, 0x00, 0x36, 0x00, 0x31, 0x00, 0x2D, 0x00, 0x36, 0x00, 0x35, 0x00, 0x35, 0x00, 0x35, 0x00,
            0x30, 0x00, 0x36, 0x00, 0x2D, 0x00, 0x30, 0x00, 0x33, 0x00, 0x2D, 0x00, 0x31, 0x00, 0x30, 0x00,
            0x33, 0x00, 0x33, 0x00, 0x2D, 0x00, 0x39, 0x00, 0x32, 0x00, 0x30, 0x00, 0x30, 0x00, 0x2E, 0x00,
            0x30, 0x00, 0x30, 0x00, 0x30, 0x00, 0x30, 0x00, 0x2D, 0x00, 0x30, 0x00, 0x36, 0x00, 0x35, 0x00,
            0x32, 0x00, 0x30, 0x00, 0x31, 0x00, 0x33, 0x00, 0x00, 0x00, 0xDE, 0x19, 0x02, 0xCF, 0x1F, 0x35,
            0x97, 0x4E, 0x8A, 0x8F, 0xB8, 0x07, 0xB1, 0x92, 0xB5, 0xB5, 0x97, 0x42, 0xEC, 0x3A, 0x76, 0x84,
            0xD5, 0x01, 0x32, 0x00, 0x00, 0x00, 0x78, 0x00, 0x00, 0x00, 0x60, 0x27, 0x00, 0x00, 0xC4, 0x1E,
            0xAA, 0x8B, 0xDD, 0x0C, 0xAB, 0x55, 0x6A, 0xCE, 0xAF, 0xAC, 0x7F, 0x5F, 0xBD, 0xE9
        };

        public static readonly byte[] KMSv5Response =
        {
            0x00, 0x00, 0x05, 0x00, 0xBE, 0x96, 0xF9, 0x04, 0x54, 0x17, 0x3F, 0xAF, 0xE3, 0x08, 0x50, 0xEB,
            0x22, 0xBA, 0x53, 0xBF, 0xF2, 0x6A, 0x7B, 0xC9, 0x05, 0x1D, 0xB5, 0x19, 0xDF, 0x98, 0xE2, 0x71,
            0x4D, 0x00, 0x61, 0xE9, 0x9D, 0x03, 0xFB, 0x31, 0xF9, 0x1F, 0x2E, 0x60, 0x59, 0xC7, 0x73, 0xC8,
            0xE8, 0xB6, 0xE1, 0x2B, 0x39, 0xC6, 0x35, 0x0E, 0x68, 0x7A, 0xAA, 0x4F, 0x28, 0x23, 0x12, 0x18,
            0xE3, 0xAA, 0x84, 0x81, 0x6E, 0x82, 0xF0, 0x3F, 0xD9, 0x69, 0xA9, 0xDF, 0xBA, 0x5F, 0xCA, 0x32,
            0x54, 0xB2, 0x52, 0x3B, 0x3E, 0xD1, 0x5C, 0x65, 0xBC, 0x3E, 0x59, 0x0D, 0x15, 0x9F, 0x37, 0xEC,
            0x30, 0x9C, 0xCC, 0x1B, 0x39, 0x0D, 0x21, 0x32, 0x29, 0xA2, 0xDD, 0xC7, 0xC1, 0x69, 0xF2, 0x72,
            0x3F, 0x00, 0x98, 0x1E, 0xF8, 0x9A, 0x79, 0x44, 0x5D, 0x25, 0x80, 0x7B, 0xF5, 0xE1, 0x7C, 0x68,
            0x25, 0xAA, 0x0D, 0x67, 0x98, 0xE5, 0x59, 0x9B, 0x04, 0xC1, 0x23, 0x33, 0x48, 0xFB, 0x28, 0xD0,
            0x76, 0xDF, 0x01, 0x56, 0xE7, 0xEC, 0xBF, 0x1A, 0xA2, 0x22, 0x28, 0xCA, 0xB1, 0xB4, 0x4C, 0x30,
            0x14, 0x6F, 0xD2, 0x2E, 0x01, 0x2A, 0x04, 0xE3, 0xBD, 0xA7, 0x41, 0x2F, 0xC9, 0xEF, 0x53, 0xC0,
            0x70, 0x48, 0xF1, 0xB2, 0xB6, 0xEA, 0xE7, 0x0F, 0x7A, 0x15, 0xD1, 0xA6, 0xFE, 0x23, 0xC8, 0xF3,
            0xE1, 0x02, 0x9E, 0xA0, 0x4E, 0xBD, 0xF5, 0xEA, 0x53, 0x74, 0x8E, 0x74, 0xA1, 0xA1, 0xBD, 0xBE,
            0x66, 0xC4, 0x73, 0x8F, 0x24, 0xA7, 0x2A, 0x2F, 0xE3, 0xD9, 0xF4, 0x28, 0xD9, 0xF8, 0xA3, 0x93,
            0x03, 0x9E, 0x29, 0xAB
        };

        public static readonly byte[] KMSv6Response =
        {
            0x00, 0x00, 0x06, 0x00, 0x54, 0xD3, 0x40, 0x08, 0xF3, 0xCD, 0x03, 0xEF, 0xC8, 0x15, 0x87, 0x9E,
            0xCA, 0x2E, 0x85, 0xFB, 0xE6, 0xF6, 0x73, 0x66, 0xFB, 0xDA, 0xBB, 0x7B, 0xB1, 0xBC, 0xD6, 0xF9,
            0x5C, 0x41, 0xA0, 0xFE, 0xE1, 0x74, 0xC4, 0xBB, 0x91, 0xE5, 0xDE, 0x6D, 0x3A, 0x11, 0xD5, 0xFC,
            0x68, 0xC0, 0x7B, 0x82, 0xB2, 0x24, 0xD1, 0x85, 0xBA, 0x45, 0xBF, 0xF1, 0x26, 0xFA, 0xA5, 0xC6,
            0x61, 0x70, 0x69, 0x69, 0x6E, 0x0F, 0x0B, 0x60, 0xB7, 0x3D, 0xE8, 0xF1, 0x47, 0x0B, 0x65, 0xFD,
            0xA7, 0x30, 0x1E, 0xF6, 0xA4, 0xD0, 0x79, 0xC4, 0x58, 0x8D, 0x81, 0xFD, 0xA7, 0xE7, 0x53, 0xF1,
            0x67, 0x78, 0xF0, 0x0F, 0x60, 0x8F, 0xC8, 0x16, 0x35, 0x22, 0x94, 0x48, 0xCB, 0x0F, 0x8E, 0xB2,
            0x1D, 0xF7, 0x3E, 0x28, 0x42, 0x55, 0x6B, 0x07, 0xE3, 0xE8, 0x51, 0xD5, 0xFA, 0x22, 0x0C, 0x86,
            0x65, 0x0D, 0x3F, 0xDD, 0x8D, 0x9B, 0x1B, 0xC9, 0xD3, 0xB8, 0x3A, 0xEC, 0xF1, 0x11, 0x19, 0x25,
            0xF7, 0x84, 0x4A, 0x4C, 0x0A, 0xB5, 0x31, 0x94, 0x37, 0x76, 0xCE, 0xE7, 0xAB, 0xA9, 0x69, 0xDF,
            0xA4, 0xC9, 0x22, 0x6C, 0x23, 0xFF, 0x6B, 0xFC, 0xDA, 0x78, 0xD8, 0xC4, 0x8F, 0x74, 0xBB, 0x26,
            0x05, 0x00, 0x98, 0x9B, 0xE5, 0xE2, 0xAD, 0x0D, 0x57, 0x95, 0x80, 0x66, 0x8E, 0x43, 0x74, 0x87,
            0x93, 0x1F, 0xF4, 0xB2, 0x2C, 0x20, 0x5F, 0xD8, 0x9C, 0x4C, 0x56, 0xB3, 0x57, 0x44, 0x62, 0x68,
            0x8D, 0xAA, 0x40, 0x11, 0x9D, 0x84, 0x62, 0x0E, 0x43, 0x8A, 0x1D, 0xF0, 0x1C, 0x49, 0xD8, 0x56,
            0xEF, 0x4C, 0xD3, 0x64, 0xBA, 0x0D, 0xEF, 0x87, 0xB5, 0x2C, 0x88, 0xF3, 0x18, 0xFF, 0x3A, 0x8C,
            0xF5, 0xA6, 0x78, 0x5C, 0x62, 0xE3, 0x9E, 0x4C, 0xB6, 0x31, 0x2D, 0x06, 0x80, 0x92, 0xBC, 0x2E,
            0x92, 0xA6, 0x56, 0x96
        };

        // 2^31 - 1 minutes
        public static ulong TimerMax = (ulong)TimeSpan.FromMinutes(2147483647).Ticks;

        public static readonly string ZeroCID = new string('0', 48);
    }

    public static class BinaryReaderExt
    {
        public static void Align(this BinaryReader reader, int to)
        {
            int pos = (int)reader.BaseStream.Position;
            reader.BaseStream.Seek(-pos & (to - 1), SeekOrigin.Current);
        }

        public static string ReadNullTerminatedString(this BinaryReader reader, int maxLen)
        {
            return Encoding.Unicode.GetString(reader.ReadBytes(maxLen)).Split(new char[] { '\0' }, 2)[0];
        }
    }

    public static class BinaryWriterExt
    {
        public static void Align(this BinaryWriter writer, int to)
        {
            int pos = (int)writer.BaseStream.Position;
            writer.WritePadding(-pos & (to - 1));
        }

        public static void WritePadding(this BinaryWriter writer, int len)
        {
            writer.Write(Enumerable.Repeat((byte)0, len).ToArray());
        }

        public static void WriteFixedString(this BinaryWriter writer, string str, int bLen)
        {
            writer.Write(Encoding.ASCII.GetBytes(str));
            writer.WritePadding(bLen - str.Length);
        }

        public static void WriteFixedString16(this BinaryWriter writer, string str, int bLen)
        {
            byte[] bstr = Utils.EncodeString(str);
            writer.Write(bstr);
            writer.WritePadding(bLen - bstr.Length);
        }

        public static byte[] GetBytes(this BinaryWriter writer)
        {
            return ((MemoryStream)writer.BaseStream).ToArray();
        }
    }

    public static class ByteArrayExt
    {
        public static byte[] CastToArray<T>(this T data) where T : struct
        {
            int size = Marshal.SizeOf(typeof(T));
            byte[] result = new byte[size];
            GCHandle handle = GCHandle.Alloc(result, GCHandleType.Pinned);
            try
            {
                Marshal.StructureToPtr(data, handle.AddrOfPinnedObject(), false);
            }
            finally
            {
                handle.Free();
            }
            return result;
        }

        public static T CastToStruct<T>(this byte[] data) where T : struct
        {
            GCHandle handle = GCHandle.Alloc(data, GCHandleType.Pinned);
            try
            {
                IntPtr ptr = handle.AddrOfPinnedObject();
                return (T)Marshal.PtrToStructure(ptr, typeof(T));
            }
            finally
            {
                handle.Free();
            }
        }
    }

    public static class FileStreamExt
    {
        public static byte[] ReadAllBytes(this FileStream fs)
        {
            BinaryReader br = new BinaryReader(fs);
            return br.ReadBytes((int)fs.Length);
        }

        public static void WriteAllBytes(this FileStream fs, byte[] data)
        {
            fs.Seek(0, SeekOrigin.Begin);
            fs.SetLength(data.Length);
            fs.Write(data, 0, data.Length);
        }
    }

    public static class Utils
    {
        public static string DecodeString(byte[] data)
        {
            return Encoding.Unicode.GetString(data).Trim('\0');
        }

        public static byte[] EncodeString(string str)
        {
            return Encoding.Unicode.GetBytes(str + '\0');
        }

        [DllImport("kernel32.dll")]
        public static extern uint GetSystemDefaultLCID();

        public static uint CRC32(byte[] data)
        {
            const uint polynomial = 0x04C11DB7;
            uint crc = 0xffffffff;

            foreach (byte b in data)
            {
                crc ^= (uint)b << 24;
                for (int bit = 0; bit < 8; bit++)
                {
                    if ((crc & 0x80000000) != 0)
                    {
                        crc = (crc << 1) ^ polynomial;
                    }
                    else
                    {
                        crc <<= 1;
                    }
                }
            }
            return ~crc;
        }

        public static void KillSPP()
        {
            ServiceController sc;

            try
            {
                sc = new ServiceController("sppsvc");

                if (sc.Status == ServiceControllerStatus.Stopped)
                    return;
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException("Unable to access sppsvc: " + ex.Message);
            }

            Logger.WriteLine("Stopping sppsvc...");

            bool stopped = false;

            for (int i = 0; stopped == false && i < 60; i++)
            {
                try
                {
                    if (sc.Status != ServiceControllerStatus.StopPending)
                        sc.Stop();

                    sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromMilliseconds(500));
                }
                catch (System.ServiceProcess.TimeoutException)
                {
                    continue;
                }
                catch (InvalidOperationException)
                {
                    System.Threading.Thread.Sleep(500);
                    continue;
                }

                stopped = true;
            }

            if (!stopped)
                throw new System.TimeoutException("Failed to stop sppsvc");

            Logger.WriteLine("sppsvc stopped successfully.");
        }

        public static string GetPSPath(PSVersion version)
        {
            switch (version)
            {
                case PSVersion.Win7:
                    return Directory.GetFiles(
                        Environment.GetFolderPath(Environment.SpecialFolder.System),
                        "7B296FB0-376B-497e-B012-9C450E1B7327-*.C7483456-A289-439d-8115-601632D005A0")
                    .FirstOrDefault() ?? "";
                case PSVersion.Win8Early:
                case PSVersion.WinBlue:
                case PSVersion.Win8:
                case PSVersion.WinModern:
                    return Path.Combine(
                        Environment.ExpandEnvironmentVariables(
                            (string)Registry.GetValue(
                                @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform",
                                "TokenStore",
                                string.Empty
                                )
                            ),
                            "data.dat"
                        );
                default:
                    return "";
            }
        }

        public static string GetTokensPath(PSVersion version)
        {
            switch (version)
            {
                case PSVersion.Win7:
                    return Path.Combine(
                        Environment.ExpandEnvironmentVariables("%WINDIR%"),
                        @"ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat"
                    );
                case PSVersion.Win8Early:
                case PSVersion.WinBlue:
                case PSVersion.Win8:
                case PSVersion.WinModern:
                    return Path.Combine(
                        Environment.ExpandEnvironmentVariables(
                            (string)Registry.GetValue(
                                @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform",
                                "TokenStore",
                                string.Empty
                                )
                            ),
                            "tokens.dat"
                        );
                default:
                    return "";
            }
        }

        public static IPhysicalStore GetStore(PSVersion version, bool production)
        {
            string psPath;

            try
            {
                psPath = GetPSPath(version);
            }
            catch
            {
                throw new FileNotFoundException("Failed to get path of physical store.");
            }

            if (string.IsNullOrEmpty(psPath) || !File.Exists(psPath))
            {
                throw new FileNotFoundException(string.Format("Physical store not found at expected path {0}.", psPath));
            }

            if (version == PSVersion.Vista)
            {
                throw new NotSupportedException("Physical store editing is not supported for Windows Vista.");
            }

            return version == PSVersion.Win7 ? new PhysicalStoreWin7(psPath, production) : (IPhysicalStore)new PhysicalStoreModern(psPath, production, version);
        }

        public static ITokenStore GetTokenStore(PSVersion version)
        {
            string tokPath;

            try
            {
                tokPath = GetTokensPath(version);
            }
            catch
            {
                throw new FileNotFoundException("Failed to get path of physical store.");
            }

            if (string.IsNullOrEmpty(tokPath) || !File.Exists(tokPath))
            {
                throw new FileNotFoundException(string.Format("Token store not found at expected path {0}.", tokPath));
            }

            return new TokenStoreModern(tokPath);
        }

        public static string GetArchitecture()
        {
            string arch = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE", EnvironmentVariableTarget.Machine).ToUpperInvariant();
            return arch == "AMD64" ? "X64" : arch;
        }

        public static PSVersion DetectVersion()
        {
            int build = Environment.OSVersion.Version.Build;

            if (build >= 9600) return PSVersion.WinModern;
            if (build >= 6000 && build <= 6003) return PSVersion.Vista;
            if (build >= 7600 && build <= 7602) return PSVersion.Win7;
            if (build == 9200) return PSVersion.Win8;

            throw new NotSupportedException("Unable to auto-detect version info, please specify one manually using the /ver argument.");
        }

        public static bool DetectCurrentKey()
        {
            SLApi.RefreshLicenseStatus();

            using (RegistryKey wpaKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\WPA"))
            {
                foreach (string subKey in wpaKey.GetSubKeyNames())
                {
                    if (subKey.StartsWith("8DEC0AF1") && subKey.EndsWith("-1"))
                    {
                        return subKey.Contains("P");
                    }
                }
            }

            throw new FileNotFoundException("Failed to autodetect key type, specify physical store key with /prod or /test arguments.");
        }

        public static void DumpStore(PSVersion version, bool production, string filePath, string encrFilePath)
        {
            if (encrFilePath == null)
            {
                encrFilePath = GetPSPath(version);
            }

            if (string.IsNullOrEmpty(encrFilePath) || !File.Exists(encrFilePath))
            {
                throw new FileNotFoundException("Store does not exist at expected path '" + encrFilePath + "'.");
            }

            KillSPP();

            using (FileStream fs = File.Open(encrFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            {
                byte[] encrData = fs.ReadAllBytes();
                File.WriteAllBytes(filePath, PhysStoreCrypto.DecryptPhysicalStore(encrData, production));
            }

            Logger.WriteLine("Store dumped successfully to '" + filePath + "'.");
        }

        public static void LoadStore(PSVersion version, bool production, string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException("Store file '" + filePath + "' does not exist.");
            }

            KillSPP();

            using (IPhysicalStore store = GetStore(version, production))
            {
                store.WriteRaw(File.ReadAllBytes(filePath));
            }

            Logger.WriteLine("Loaded store file succesfully.");
        }
    }

    public static class Logger
    {
        public static bool HideOutput = false;

        public static void WriteLine(string line)
        {
            if (!HideOutput) Console.WriteLine(line);
        }
    }
}


// SPP/PKeyConfig.cs
namespace LibTSforge.SPP
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Xml;

    public enum PKeyAlgorithm
    {
        PKEY2005,
        PKEY2009
    }

    public class KeyRange
    {
        public int Start;
        public int End;
        public string EulaType;
        public string PartNumber;
        public bool Valid;

        public bool Contains(int n)
        {
            return Start <= n && End <= n;
        }
    }

    public class ProductConfig
    {
        public int GroupId;
        public string Edition;
        public string Description;
        public string Channel;
        public bool Randomized;
        public PKeyAlgorithm Algorithm;
        public List<KeyRange> Ranges;
        public Guid ActivationId;

        private List<KeyRange> GetPkeyRanges()
        {
            if (Ranges.Count == 0)
            {
                throw new ArgumentException("No key ranges.");
            }

            if (Algorithm == PKeyAlgorithm.PKEY2005)
            {
                return Ranges;
            }

            List<KeyRange> FilteredRanges = Ranges.Where(r => !r.EulaType.Contains("WAU")).ToList();

            if (FilteredRanges.Count == 0)
            {
                throw new NotSupportedException("Specified Activation ID is usable only for Windows Anytime Upgrade. Please use a non-WAU Activation ID instead.");
            }

            return FilteredRanges;
        }

        public ProductKey GetRandomKey()
        {
            List<KeyRange> KeyRanges = GetPkeyRanges();
            Random rnd = new Random();

            KeyRange range = KeyRanges[rnd.Next(KeyRanges.Count)];
            int serial = rnd.Next(range.Start, range.End);

            return new ProductKey(serial, 0, false, Algorithm, this, range);
        }
    }

    public class PKeyConfig
    {
        public Dictionary<Guid, ProductConfig> Products = new Dictionary<Guid, ProductConfig>();
        private List<Guid> loadedPkeyConfigs = new List<Guid>();

        public void LoadConfig(Guid actId)
        {
            string pkcData;
            Guid pkcFileId = SLApi.GetPkeyConfigFileId(actId);

            if (loadedPkeyConfigs.Contains(pkcFileId)) return;

            string licConts = SLApi.GetLicenseContents(pkcFileId);

            using (TextReader tr = new StringReader(licConts))
            {
                XmlDocument lic = new XmlDocument();
                lic.Load(tr);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(lic.NameTable);
                nsmgr.AddNamespace("rg", "urn:mpeg:mpeg21:2003:01-REL-R-NS");
                nsmgr.AddNamespace("r", "urn:mpeg:mpeg21:2003:01-REL-R-NS");
                nsmgr.AddNamespace("tm", "http://www.microsoft.com/DRM/XrML2/TM/v2");

                XmlNode root = lic.DocumentElement;
                XmlNode pkcDataNode = root.SelectSingleNode("/rg:licenseGroup/r:license/r:otherInfo/tm:infoTables/tm:infoList/tm:infoBin[@name=\"pkeyConfigData\"]", nsmgr);
                pkcData = Encoding.UTF8.GetString(Convert.FromBase64String(pkcDataNode.InnerText));
            }

            using (TextReader tr = new StringReader(pkcData))
            {
                XmlDocument lic = new XmlDocument();
                lic.Load(tr);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(lic.NameTable);
                nsmgr.AddNamespace("p", "http://www.microsoft.com/DRM/PKEY/Configuration/2.0");
                XmlNodeList configNodes = lic.SelectNodes("//p:ProductKeyConfiguration/p:Configurations/p:Configuration", nsmgr);
                XmlNodeList rangeNodes = lic.SelectNodes("//p:ProductKeyConfiguration/p:KeyRanges/p:KeyRange", nsmgr);
                XmlNodeList pubKeyNodes = lic.SelectNodes("//p:ProductKeyConfiguration/p:PublicKeys/p:PublicKey", nsmgr);

                Dictionary<int, PKeyAlgorithm> algorithms = new Dictionary<int, PKeyAlgorithm>();
                Dictionary<string, List<KeyRange>> ranges = new Dictionary<string, List<KeyRange>>();

                Dictionary<string, PKeyAlgorithm> algoConv = new Dictionary<string, PKeyAlgorithm>
                {
                    { "msft:rm/algorithm/pkey/2005", PKeyAlgorithm.PKEY2005 },
                    { "msft:rm/algorithm/pkey/2009", PKeyAlgorithm.PKEY2009 }
                };

                foreach (XmlNode pubKeyNode in pubKeyNodes)
                {
                    int group = int.Parse(pubKeyNode.SelectSingleNode("./p:GroupId", nsmgr).InnerText);
                    algorithms[group] = algoConv[pubKeyNode.SelectSingleNode("./p:AlgorithmId", nsmgr).InnerText];
                }

                foreach (XmlNode rangeNode in rangeNodes)
                {
                    string refActIdStr = rangeNode.SelectSingleNode("./p:RefActConfigId", nsmgr).InnerText;

                    if (!ranges.ContainsKey(refActIdStr))
                    {
                        ranges[refActIdStr] = new List<KeyRange>();
                    }

                    KeyRange keyRange = new KeyRange();
                    keyRange.Start = int.Parse(rangeNode.SelectSingleNode("./p:Start", nsmgr).InnerText);
                    keyRange.End = int.Parse(rangeNode.SelectSingleNode("./p:End", nsmgr).InnerText);
                    keyRange.EulaType = rangeNode.SelectSingleNode("./p:EulaType", nsmgr).InnerText;
                    keyRange.PartNumber = rangeNode.SelectSingleNode("./p:PartNumber", nsmgr).InnerText;
                    keyRange.Valid = rangeNode.SelectSingleNode("./p:IsValid", nsmgr).InnerText.ToLower() == "true";

                    ranges[refActIdStr].Add(keyRange);
                }

                foreach (XmlNode configNode in configNodes)
                {
                    string refActIdStr = configNode.SelectSingleNode("./p:ActConfigId", nsmgr).InnerText;
                    Guid refActId = new Guid(refActIdStr);
                    int group = int.Parse(configNode.SelectSingleNode("./p:RefGroupId", nsmgr).InnerText);
                    List<KeyRange> keyRanges = ranges[refActIdStr];

                    if (keyRanges.Count > 0 && !Products.ContainsKey(refActId))
                    {
                        ProductConfig productConfig = new ProductConfig();
                        productConfig.GroupId = group;
                        productConfig.Edition = configNode.SelectSingleNode("./p:EditionId", nsmgr).InnerText;
                        productConfig.Description = configNode.SelectSingleNode("./p:ProductDescription", nsmgr).InnerText;
                        productConfig.Channel = configNode.SelectSingleNode("./p:ProductKeyType", nsmgr).InnerText;
                        productConfig.Randomized = configNode.SelectSingleNode("./p:ProductKeyType", nsmgr).InnerText.ToLower() == "true";
                        productConfig.Algorithm = algorithms[group];
                        productConfig.Ranges = keyRanges;
                        productConfig.ActivationId = refActId;

                        Products[refActId] = productConfig;
                    }
                }
            }

            loadedPkeyConfigs.Add(pkcFileId);
        }

        public ProductConfig MatchParams(int group, int serial)
        {
            foreach (ProductConfig config in Products.Values)
            {
                if (config.GroupId == group)
                {
                    foreach (KeyRange range in config.Ranges)
                    {
                        if (range.Contains(serial))
                        {
                            return config;
                        }
                    }
                }
            }

            throw new FileNotFoundException("Failed to find product matching supplied product key parameters.");
        }

        public void LoadAllConfigs(Guid appId)
        {
            foreach (Guid actId in SLApi.GetActivationIds(appId))
            {
                try
                {
                    LoadConfig(actId);
                } 
                catch (ArgumentException)
                {

                }
            }
        }

        public PKeyConfig()
        {

        }
    }
}


// SPP/ProductKey.cs
namespace LibTSforge.SPP
{
    using System;
    using System.IO;
    using System.Linq;
    using LibTSforge.Crypto;
    using LibTSforge.PhysicalStore;

    public class ProductKey
    {
        private static readonly string ALPHABET = "BCDFGHJKMPQRTVWXY2346789";

        private readonly ulong klow;
        private readonly ulong khigh;

        public int Group;
        public int Serial;
        public ulong Security;
        public bool Upgrade;
        public PKeyAlgorithm Algorithm;
        public string EulaType;
        public string PartNumber;
        public string Edition;
        public string Channel;
        public Guid ActivationId;

        private string mpc;
        private string pid2;

        public byte[] KeyBytes
        {
            get { return BitConverter.GetBytes(klow).Concat(BitConverter.GetBytes(khigh)).ToArray(); }
        }

        public ProductKey(int serial, ulong security, bool upgrade, PKeyAlgorithm algorithm, ProductConfig config, KeyRange range)
        {
            Group = config.GroupId;
            Serial = serial;
            Security = security;
            Upgrade = upgrade;
            Algorithm = algorithm;
            EulaType = range.EulaType;
            PartNumber = range.PartNumber.Split(':', ';')[0];
            Edition = config.Edition;
            Channel = config.Channel;
            ActivationId = config.ActivationId;

            klow = ((security & 0x3fff) << 50 | ((ulong)serial & 0x3fffffff) << 20 | ((ulong)Group & 0xfffff));
            khigh = ((upgrade ? (ulong)1 : 0) << 49 | ((security >> 14) & 0x7fffffffff));

            uint checksum = Utils.CRC32(KeyBytes) & 0x3ff;

            khigh |= ((ulong)checksum << 39);
        }

        public string GetAlgoUri()
        {
            return "msft:rm/algorithm/pkey/" + (Algorithm == PKeyAlgorithm.PKEY2005 ? "2005" : (Algorithm == PKeyAlgorithm.PKEY2009 ? "2009" : "Unknown"));
        }

        public Guid GetPkeyId()
        {
            VariableBag pkb = new VariableBag();
            pkb.Blocks.AddRange(new CRCBlock[]
            {
                new CRCBlock
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingProductKey",
                    ValueAsStr = ToString()
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingMiscData",
                    Value = new byte[] { }
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingAlgorithm",
                    ValueAsStr = GetAlgoUri()
                }
            });

            return new Guid(CryptoUtils.SHA256Hash(pkb.Serialize()).Take(16).ToArray());
        }

        public string GetDefaultMPC()
        {
            int build = Environment.OSVersion.Version.Build;
            string defaultMPC = build >= 10240 ? "03612" :
                                build >= 9600 ? "06401" :
                                build >= 9200 ? "05426" :
                                "55041";
            return defaultMPC;
        }

        public string GetMPC()
        {
            if (mpc != null)
            {
                return mpc;
            }

            mpc = GetDefaultMPC();

            // setup.cfg doesn't exist in Windows 8+
            string setupcfg = string.Format("{0}\\oobe\\{1}", Environment.SystemDirectory, "setup.cfg");

            if (!File.Exists(setupcfg) || Edition.Contains(";"))
            {
                return mpc;
            }

            string mpcKey = string.Format("{0}.{1}=", Utils.GetArchitecture(), Edition);
            string localMPC = File.ReadAllLines(setupcfg).FirstOrDefault(line => line.Contains(mpcKey));
            if (localMPC != null)
            {
                mpc = localMPC.Split('=')[1].Trim();
            }

            return mpc;
        }

        public string GetPid2()
        {
            if (pid2 != null)
            {
                return pid2;
            }

            pid2 = "";

            if (Algorithm == PKeyAlgorithm.PKEY2005)
            {
                string mpc = GetMPC();
                string serialHigh;
                int serialLow;
                int lastPart;

                if (EulaType == "OEM")
                {
                    serialHigh = "OEM";
                    serialLow = ((Group / 2) % 100) * 10000 + (Serial / 100000);
                    lastPart = Serial % 100000;
                }
                else
                {
                    serialHigh = (Serial / 1000000).ToString("D3");
                    serialLow = Serial % 1000000;
                    lastPart = ((Group / 2) % 100) * 1000 + new Random().Next(1000);
                }

                int checksum = 0;

                foreach (char c in serialLow.ToString())
                {
                    checksum += int.Parse(c.ToString());
                }
                checksum = 7 - (checksum % 7);

                pid2 = string.Format("{0}-{1}-{2:D6}{3}-{4:D5}", mpc, serialHigh, serialLow, checksum, lastPart);
            }

            return pid2;
        }

        public byte[] GetPid3()
        {
            BinaryWriter writer = new BinaryWriter(new MemoryStream());
            writer.Write(0xA4);
            writer.Write(0x3);
            writer.WriteFixedString(GetPid2(), 24);
            writer.Write(Group);
            writer.WriteFixedString(PartNumber, 16);
            writer.WritePadding(0x6C);
            byte[] data = writer.GetBytes();
            byte[] crc = BitConverter.GetBytes(~Utils.CRC32(data.Reverse().ToArray())).Reverse().ToArray();
            writer.Write(crc);

            return writer.GetBytes();
        }

        public byte[] GetPid4()
        {
            BinaryWriter writer = new BinaryWriter(new MemoryStream());
            writer.Write(0x4F8);
            writer.Write(0x4);
            writer.WriteFixedString16(GetExtendedPid(), 0x80);
            writer.WriteFixedString16(ActivationId.ToString(), 0x80);
            writer.WritePadding(0x10);
            writer.WriteFixedString16(Edition, 0x208);
            writer.Write(Upgrade ? (ulong)1 : 0);
            writer.WritePadding(0x50);
            writer.WriteFixedString16(PartNumber, 0x80);
            writer.WriteFixedString16(Channel, 0x80);
            writer.WriteFixedString16(EulaType, 0x80);

            return writer.GetBytes();
        }

        public string GetExtendedPid()
        {
            string mpc = GetMPC();
            int serialHigh = Serial / 1000000;
            int serialLow = Serial % 1000000;
            int licenseType;
            uint lcid = Utils.GetSystemDefaultLCID();
            int build = Environment.OSVersion.Version.Build;
            int dayOfYear = DateTime.Now.DayOfYear;
            int year = DateTime.Now.Year;

            switch (EulaType)
            {
                case "OEM":
                    licenseType = 2;
                    break;

                case "Volume":
                    licenseType = 3;
                    break;

                default:
                    licenseType = 0;
                    break;
            }

            return string.Format(
                "{0}-{1:D5}-{2:D3}-{3:D6}-{4:D2}-{5:D4}-{6:D4}.0000-{7:D3}{8:D4}",
                mpc,
                Group,
                serialHigh,
                serialLow,
                licenseType,
                lcid,
                build,
                dayOfYear,
                year
            );
        }

        public byte[] GetPhoneData(PSVersion version)
        {
            if (version == PSVersion.Win7)
            {
                Random rnd = new Random(Group * 1000000000 + Serial);
                byte[] data = new byte[8];
                rnd.NextBytes(data);
                return data;
            }

            int serialHigh = Serial / 1000000;
            int serialLow = Serial % 1000000;

            BinaryWriter writer = new BinaryWriter(new MemoryStream());
            writer.Write(new Guid("B8731595-A2F6-430B-A799-FBFFB81A8D73").ToByteArray());
            writer.Write(Group);
            writer.Write(serialHigh);
            writer.Write(serialLow);
            writer.Write(Upgrade ? 1 : 0);
            writer.Write(Security);

            return writer.GetBytes();
        }

        public override string ToString()
        {
            string keyStr = "";
            Random rnd = new Random(Group * 1000000000 + Serial);

            if (Algorithm == PKeyAlgorithm.PKEY2005)
            {
                keyStr = "H4X3DH4X3DH4X3DH4X3D";

                for (int i = 0; i < 5; i++)
                {
                    keyStr += ALPHABET[rnd.Next(24)];
                }
            }
            else if (Algorithm == PKeyAlgorithm.PKEY2009)
            {
                int last = 0;
                byte[] bKey = KeyBytes;

                for (int i = 24; i >= 0; i--)
                {
                    int current = 0;

                    for (int j = 14; j >= 0; j--)
                    {
                        current *= 0x100;
                        current += bKey[j];
                        bKey[j] = (byte)(current / 24);
                        current %= 24;
                        last = current;
                    }

                    keyStr = ALPHABET[current] + keyStr;
                }

                keyStr = keyStr.Substring(1, last) + "N" + keyStr.Substring(last + 1, keyStr.Length - last - 1);
            }

            for (int i = 5; i < keyStr.Length; i += 6)
            {
                keyStr = keyStr.Insert(i, "-");
            }

            return keyStr;
        }
    }
}


// SPP/SLAPI.cs
namespace LibTSforge.SPP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;

    public static class SLApi
    {
        private enum SLIDTYPE
        {
            SL_ID_APPLICATION,
            SL_ID_PRODUCT_SKU,
            SL_ID_LICENSE_FILE,
            SL_ID_LICENSE,
            SL_ID_PKEY,
            SL_ID_ALL_LICENSES,
            SL_ID_ALL_LICENSE_FILES,
            SL_ID_STORE_TOKEN,
            SL_ID_LAST
        }

        private enum SLDATATYPE
        {
            SL_DATA_NONE,
            SL_DATA_SZ,
            SL_DATA_DWORD,
            SL_DATA_BINARY,
            SL_DATA_MULTI_SZ,
            SL_DATA_SUM
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SL_LICENSING_STATUS
        {
            public Guid SkuId;
            public uint eStatus;
            public uint dwGraceTime;
            public uint dwTotalGraceDays;
            public uint hrReason;
            public ulong qwValidityExpiration;
        }

        public static readonly Guid WINDOWS_APP_ID = new Guid("55c92734-d682-4d71-983e-d6ec3f16059f");

        [DllImport("sppc.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SLOpen(out IntPtr hSLC);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SLClose(IntPtr hSLC);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetWindowsInformationDWORD(string ValueName, ref int Value);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLInstallProofOfPurchase(IntPtr hSLC, string pwszPKeyAlgorithm, string pwszPKeyString, uint cbPKeySpecificData, byte[] pbPKeySpecificData, ref Guid PKeyId);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLUninstallProofOfPurchase(IntPtr hSLC, ref Guid PKeyId);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetPKeyInformation(IntPtr hSLC, ref Guid pPKeyId, string pwszValueName, out SLDATATYPE peDataType, out uint pcbValue, out IntPtr ppbValue);

        [DllImport("sppcext.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLActivateProduct(IntPtr hSLC, ref Guid pProductSkuId, byte[] cbAppSpecificData, byte[] pvAppSpecificData, byte[] pActivationInfo, string pwszProxyServer, ushort wProxyPort);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGenerateOfflineInstallationId(IntPtr hSLC, ref Guid pProductSkuId, ref string ppwszInstallationId);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLDepositOfflineConfirmationId(IntPtr hSLC, ref Guid pProductSkuId, string pwszInstallationId, string pwszConfirmationId);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetSLIDList(IntPtr hSLC, SLIDTYPE eQueryIdType, ref Guid pQueryId, SLIDTYPE eReturnIdType, out uint pnReturnIds, out IntPtr ppReturnIds);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SLGetLicensingStatusInformation(IntPtr hSLC, ref Guid pAppID, IntPtr pProductSkuId, string pwszRightName, out uint pnStatusCount, out IntPtr ppLicensingStatus);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetInstalledProductKeyIds(IntPtr hSLC, ref Guid pProductSkuId, out uint pnProductKeyIds, out IntPtr ppProductKeyIds);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLConsumeWindowsRight(uint unknown);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetProductSkuInformation(IntPtr hSLC, ref Guid pProductSkuId, string pwszValueName, out SLDATATYPE peDataType, out uint pcbValue, out IntPtr ppbValue);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetProductSkuInformation(IntPtr hSLC, ref Guid pProductSkuId, string pwszValueName, IntPtr peDataType, out uint pcbValue, out IntPtr ppbValue);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetLicense(IntPtr hSLC, ref Guid pLicenseFileId, out uint pcbLicenseFile, out IntPtr ppbLicenseFile);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLSetCurrentProductKey(IntPtr hSLC, ref Guid pProductSkuId, ref Guid pProductKeyId);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLFireEvent(IntPtr hSLC, string pwszEventId, ref Guid pApplicationId);

        public class SLContext : IDisposable
        {
            public readonly IntPtr Handle;

            public SLContext()
            {
                SLOpen(out Handle);
            }

            public void Dispose()
            {
                SLClose(Handle);
                GC.SuppressFinalize(this);
            }

            ~SLContext()
            {
                Dispose();
            }
        }

        public static Guid GetDefaultActivationID(Guid appId, bool includeActivated)
        {
            using (SLContext sl = new SLContext())
            {
                uint count;
                IntPtr pLicStat;

                SLGetLicensingStatusInformation(sl.Handle, ref appId, IntPtr.Zero, null, out count, out pLicStat);

                unsafe
                {
                    SL_LICENSING_STATUS* licensingStatuses = (SL_LICENSING_STATUS*)pLicStat;
                    for (int i = 0; i < count; i++)
                    {
                        SL_LICENSING_STATUS slStatus = licensingStatuses[i];

                        Guid actId = slStatus.SkuId;
                        if (GetInstalledPkeyID(actId) == Guid.Empty) continue;
                        if (IsAddon(actId)) continue;
                        if (!includeActivated && (slStatus.eStatus == 1)) continue;

                        return actId;
                    }
                }

                return Guid.Empty;
            }
        }

        public static string GetInstallationID(Guid actId)
        {
            using (SLContext sl = new SLContext())
            {
                string installationId = null;
                return SLGenerateOfflineInstallationId(sl.Handle, ref actId, ref installationId) == 0 ? installationId : null;
            }
        }

        public static Guid GetInstalledPkeyID(Guid actId)
        {
            using (SLContext sl = new SLContext())
            {
                uint status;
                uint count;
                IntPtr pProductKeyIds;

                status = SLGetInstalledProductKeyIds(sl.Handle, ref actId, out count, out pProductKeyIds);

                if (status != 0 || count == 0)
                {
                    return Guid.Empty;
                }

                unsafe { return *(Guid*)pProductKeyIds; }
            }
        }

        public static uint DepositConfirmationID(Guid actId, string installationId, string confirmationId)
        {
            using (SLContext sl = new SLContext())
            {
                return SLDepositOfflineConfirmationId(sl.Handle, ref actId, installationId, confirmationId);
            }
        }

        public static void RefreshLicenseStatus()
        {
            SLConsumeWindowsRight(0);
        }

        public static bool RefreshTrustedTime(Guid actId)
        {
            using (SLContext sl = new SLContext())
            {
                SLDATATYPE type;
                uint count;
                IntPtr ppbValue;

                uint status = SLGetProductSkuInformation(sl.Handle, ref actId, "TrustedTime", out type, out count, out ppbValue);
                return (int)status >= 0 && status != 0xC004F012;
            }
        }

        public static void FireStateChangedEvent(Guid appId)
        {
            using (SLContext sl = new SLContext())
            {
                SLFireEvent(sl.Handle, "msft:rm/event/licensingstatechanged", ref appId);
            }
        }

        public static Guid GetAppId(Guid actId)
        {
            using (SLContext sl = new SLContext())
            {
                uint status;
                uint count;
                IntPtr pAppIds;

                status = SLGetSLIDList(sl.Handle, SLIDTYPE.SL_ID_PRODUCT_SKU, ref actId, SLIDTYPE.SL_ID_APPLICATION, out count, out pAppIds);

                if (status != 0 || count == 0)
                {
                    return Guid.Empty;
                }

                unsafe { return *(Guid*)pAppIds; }
            }
        }

        public static bool IsAddon(Guid actId)
        {
            using (SLContext sl = new SLContext())
            {
                uint count;
                SLDATATYPE type;
                IntPtr ppbValue;

                uint status = SLGetProductSkuInformation(sl.Handle, ref actId, "DependsOn", out type, out count, out ppbValue);
                return (int)status >= 0 && status != 0xC004F012;
            }
        }

        public static Guid GetLicenseFileId(Guid licId)
        {
            using (SLContext sl = new SLContext())
            {
                uint status;
                uint count;
                IntPtr ppReturnLics;

                status = SLGetSLIDList(sl.Handle, SLIDTYPE.SL_ID_LICENSE, ref licId, SLIDTYPE.SL_ID_LICENSE_FILE, out count, out ppReturnLics);

                if (status != 0 || count == 0)
                {
                    return Guid.Empty;
                }

                unsafe { return *(Guid*)ppReturnLics; }
            }
        }

        public static Guid GetPkeyConfigFileId(Guid actId)
        {
            using (SLContext sl = new SLContext())
            {
                SLDATATYPE type;
                uint len;
                IntPtr ppReturnLics;

                uint status = SLGetProductSkuInformation(sl.Handle, ref actId, "pkeyConfigLicenseId", out type, out len, out ppReturnLics);

                if (status != 0 || len == 0)
                {
                    return Guid.Empty;
                }

                Guid pkcId = new Guid(Marshal.PtrToStringAuto(ppReturnLics));
                return GetLicenseFileId(pkcId);
            }
        }

        public static string GetLicenseContents(Guid fileId)
        {
            if (fileId == Guid.Empty) throw new ArgumentException("License contents could not be retrieved.");

            using (SLContext sl = new SLContext())
            {
                uint dataLen;
                IntPtr dataPtr;

                if (SLGetLicense(sl.Handle, ref fileId, out dataLen, out dataPtr) != 0)
                {
                    return null;
                }

                byte[] data = new byte[dataLen];
                Marshal.Copy(dataPtr, data, 0, (int)dataLen);

                data = data.Skip(Array.IndexOf(data, (byte)'<')).ToArray();
                return Encoding.UTF8.GetString(data);
            }
        }

        public static bool IsPhoneActivatable(Guid actId)
        {
            using (SLContext sl = new SLContext())
            {
                uint count;
                SLDATATYPE type;
                IntPtr ppbValue;

                uint status = SLGetProductSkuInformation(sl.Handle, ref actId, "msft:sl/EUL/PHONE/PUBLIC", out type, out count, out ppbValue);
                return status >= 0 && status != 0xC004F012;
            }
        }

        public static string GetPKeyChannel(Guid pkeyId)
        {
            using (SLContext sl = new SLContext())
            {
                SLDATATYPE type;
                uint len;
                IntPtr ppbValue;

                uint status = SLGetPKeyInformation(sl.Handle, ref pkeyId, "Channel", out type, out len, out ppbValue);

                if (status != 0 || len == 0)
                {
                    return null;
                }

                return Marshal.PtrToStringAuto(ppbValue);
            }
        }

        public static string GetMetaStr(Guid actId, string value)
        {
            using (SLContext sl = new SLContext())
            {
                uint len;
                SLDATATYPE type;
                IntPtr ppbValue;

                uint status = SLGetProductSkuInformation(sl.Handle, ref actId, value, out type, out len, out ppbValue);

                if (status != 0 || len == 0 || type != SLDATATYPE.SL_DATA_SZ)
                {
                    return null;
                }

                return Marshal.PtrToStringAuto(ppbValue);
            }
        }

        public static List<Guid> GetActivationIds(Guid appId)
        {
            using (SLContext sl = new SLContext())
            {
                uint count;
                IntPtr pLicStat;

                SLGetLicensingStatusInformation(sl.Handle, ref appId, IntPtr.Zero, null, out count, out pLicStat);

                List<Guid> result = new List<Guid>();

                unsafe
                {
                    SL_LICENSING_STATUS* licensingStatuses = (SL_LICENSING_STATUS*)pLicStat;
                    for (int i = 0; i < count; i++)
                    {
                        result.Add(licensingStatuses[i].SkuId);
                    }
                }

                return result;
            }
        }

        public static uint SetCurrentProductKey(Guid actId, Guid pkeyId)
        {
            using (SLContext sl = new SLContext())
            {
                return SLSetCurrentProductKey(sl.Handle, ref actId, ref pkeyId);
            }
        }

        public static uint InstallProductKey(ProductKey pkey)
        {
            using (SLContext sl = new SLContext())
            {
                Guid pkeyId = Guid.Empty;
                return SLInstallProofOfPurchase(sl.Handle, pkey.GetAlgoUri(), pkey.ToString(), 0, null, ref pkeyId);
            }
        }

        public static uint UninstallProductKey(Guid pkeyId)
        {
            using (SLContext sl = new SLContext())
            {
                return SLUninstallProofOfPurchase(sl.Handle, ref pkeyId);
            }
        }

        public static void UninstallAllProductKeys(Guid appId)
        {
            foreach (Guid actId in GetActivationIds(appId))
            {
                Guid pkeyId = GetInstalledPkeyID(actId);
                if (pkeyId == Guid.Empty) continue;
                if (IsAddon(actId)) continue;
                UninstallProductKey(pkeyId);
            }
        }
    }
}


// Crypto/CryptoUtils.cs
namespace LibTSforge.Crypto
{
    using System;
    using System.Linq;
    using System.Security.Cryptography;

    public static class CryptoUtils
    {
        public static byte[] GenerateRandomKey(int len)
        {
            byte[] rand = new byte[len];
            Random r = new Random();
            r.NextBytes(rand);

            return rand;
        }

        public static byte[] AESEncrypt(byte[] data, byte[] key)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, Enumerable.Repeat((byte)0, 16).ToArray());
                byte[] encryptedData = encryptor.TransformFinalBlock(data, 0, data.Length);
                return encryptedData;
            }
        }

        public static byte[] AESDecrypt(byte[] data, byte[] key)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, Enumerable.Repeat((byte)0, 16).ToArray());
                byte[] decryptedData = decryptor.TransformFinalBlock(data, 0, data.Length);
                return decryptedData;
            }
        }

        public static byte[] RSADecrypt(byte[] rsaKey, byte[] data)
        {

            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.ImportCspBlob(rsaKey);
                return rsa.Decrypt(data, false);
            }
        }

        public static byte[] RSAEncrypt(byte[] rsaKey, byte[] data)
        {
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.ImportCspBlob(rsaKey);
                return rsa.Encrypt(data, false);
            }
        }

        public static byte[] RSASign(byte[] rsaKey, byte[] data)
        {
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.ImportCspBlob(rsaKey);
                RSAPKCS1SignatureFormatter formatter = new RSAPKCS1SignatureFormatter(rsa);
                formatter.SetHashAlgorithm("SHA1");

                byte[] hash;
                using (SHA1 sha1 = SHA1.Create())
                {
                    hash = sha1.ComputeHash(data);
                }

                return formatter.CreateSignature(hash);
            }
        }

        public static bool RSAVerifySignature(byte[] rsaKey, byte[] data, byte[] signature)
        {
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.ImportCspBlob(rsaKey);
                RSAPKCS1SignatureDeformatter deformatter = new RSAPKCS1SignatureDeformatter(rsa);
                deformatter.SetHashAlgorithm("SHA1");

                byte[] hash;
                using (SHA1 sha1 = SHA1.Create())
                {
                    hash = sha1.ComputeHash(data);
                }

                return deformatter.VerifySignature(hash, signature);
            }
        }

        public static byte[] HMACSign(byte[] key, byte[] data)
        {
            HMACSHA1 hmac = new HMACSHA1(key);
            return hmac.ComputeHash(data);
        }

        public static bool HMACVerify(byte[] key, byte[] data, byte[] signature)
        {
            HMACSHA1 hmac = new HMACSHA1(key);
            return Enumerable.SequenceEqual(signature, HMACSign(key, data));
        }

        public static byte[] SHA256Hash(byte[] data)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                return sha256.ComputeHash(data);
            }
        }
    }
}


// Crypto/Keys.cs
namespace LibTSforge.Crypto
{
    public static class Keys
    {
        public static readonly byte[] PRODUCTION = {
            0x07, 0x02, 0x00, 0x00, 0x00, 0xA4, 0x00, 0x00, 0x52, 0x53, 0x41, 0x32, 0x00, 0x04, 0x00, 0x00,
            0x01, 0x00, 0x01, 0x00, 0x29, 0x87, 0xBA, 0x3F, 0x52, 0x90, 0x57, 0xD8, 0x12, 0x26, 0x6B, 0x38,
            0xB2, 0x3B, 0xF9, 0x67, 0x08, 0x4F, 0xDD, 0x8B, 0xF5, 0xE3, 0x11, 0xB8, 0x61, 0x3A, 0x33, 0x42,
            0x51, 0x65, 0x05, 0x86, 0x1E, 0x00, 0x41, 0xDE, 0xC5, 0xDD, 0x44, 0x60, 0x56, 0x3D, 0x14, 0x39,
            0xB7, 0x43, 0x65, 0xE9, 0xF7, 0x2B, 0xA5, 0xF0, 0xA3, 0x65, 0x68, 0xE9, 0xE4, 0x8B, 0x5C, 0x03,
            0x2D, 0x36, 0xFE, 0x28, 0x4C, 0xD1, 0x3C, 0x3D, 0xC1, 0x90, 0x75, 0xF9, 0x6E, 0x02, 0xE0, 0x58,
            0x97, 0x6A, 0xCA, 0x80, 0x02, 0x42, 0x3F, 0x6C, 0x15, 0x85, 0x4D, 0x83, 0x23, 0x6A, 0x95, 0x9E,
            0x38, 0x52, 0x59, 0x38, 0x6A, 0x99, 0xF0, 0xB5, 0xCD, 0x53, 0x7E, 0x08, 0x7C, 0xB5, 0x51, 0xD3,
            0x8F, 0xA3, 0x0D, 0xA0, 0xFA, 0x8D, 0x87, 0x3C, 0xFC, 0x59, 0x21, 0xD8, 0x2E, 0xD9, 0x97, 0x8B,
            0x40, 0x60, 0xB1, 0xD7, 0x2B, 0x0A, 0x6E, 0x60, 0xB5, 0x50, 0xCC, 0x3C, 0xB1, 0x57, 0xE4, 0xB7,
            0xDC, 0x5A, 0x4D, 0xE1, 0x5C, 0xE0, 0x94, 0x4C, 0x5E, 0x28, 0xFF, 0xFA, 0x80, 0x6A, 0x13, 0x53,
            0x52, 0xDB, 0xF3, 0x04, 0x92, 0x43, 0x38, 0xB9, 0x1B, 0xD9, 0x85, 0x54, 0x7B, 0x14, 0xC7, 0x89,
            0x16, 0x8A, 0x4B, 0x82, 0xA1, 0x08, 0x02, 0x99, 0x23, 0x48, 0xDD, 0x75, 0x9C, 0xC8, 0xC1, 0xCE,
            0xB0, 0xD7, 0x1B, 0xD8, 0xFB, 0x2D, 0xA7, 0x2E, 0x47, 0xA7, 0x18, 0x4B, 0xF6, 0x29, 0x69, 0x44,
            0x30, 0x33, 0xBA, 0xA7, 0x1F, 0xCE, 0x96, 0x9E, 0x40, 0xE1, 0x43, 0xF0, 0xE0, 0x0D, 0x0A, 0x32,
            0xB4, 0xEE, 0xA1, 0xC3, 0x5E, 0x9B, 0xC7, 0x7F, 0xF5, 0x9D, 0xD8, 0xF2, 0x0F, 0xD9, 0x8F, 0xAD,
            0x75, 0x0A, 0x00, 0xD5, 0x25, 0x43, 0xF7, 0xAE, 0x51, 0x7F, 0xB7, 0xDE, 0xB7, 0xAD, 0xFB, 0xCE,
            0x83, 0xE1, 0x81, 0xFF, 0xDD, 0xA2, 0x77, 0xFE, 0xEB, 0x27, 0x1F, 0x10, 0xFA, 0x82, 0x37, 0xF4,
            0x7E, 0xCC, 0xE2, 0xA1, 0x58, 0xC8, 0xAF, 0x1D, 0x1A, 0x81, 0x31, 0x6E, 0xF4, 0x8B, 0x63, 0x34,
            0xF3, 0x05, 0x0F, 0xE1, 0xCC, 0x15, 0xDC, 0xA4, 0x28, 0x7A, 0x9E, 0xEB, 0x62, 0xD8, 0xD8, 0x8C,
            0x85, 0xD7, 0x07, 0x87, 0x90, 0x2F, 0xF7, 0x1C, 0x56, 0x85, 0x2F, 0xEF, 0x32, 0x37, 0x07, 0xAB,
            0xB0, 0xE6, 0xB5, 0x02, 0x19, 0x35, 0xAF, 0xDB, 0xD4, 0xA2, 0x9C, 0x36, 0x80, 0xC6, 0xDC, 0x82,
            0x08, 0xE0, 0xC0, 0x5F, 0x3C, 0x59, 0xAA, 0x4E, 0x26, 0x03, 0x29, 0xB3, 0x62, 0x58, 0x41, 0x59,
            0x3A, 0x37, 0x43, 0x35, 0xE3, 0x9F, 0x34, 0xE2, 0xA1, 0x04, 0x97, 0x12, 0x9D, 0x8C, 0xAD, 0xF7,
            0xFB, 0x8C, 0xA1, 0xA2, 0xE9, 0xE4, 0xEF, 0xD9, 0xC5, 0xE5, 0xDF, 0x0E, 0xBF, 0x4A, 0xE0, 0x7A,
            0x1E, 0x10, 0x50, 0x58, 0x63, 0x51, 0xE1, 0xD4, 0xFE, 0x57, 0xB0, 0x9E, 0xD7, 0xDA, 0x8C, 0xED,
            0x7D, 0x82, 0xAC, 0x2F, 0x25, 0x58, 0x0A, 0x58, 0xE6, 0xA4, 0xF4, 0x57, 0x4B, 0xA4, 0x1B, 0x65,
            0xB9, 0x4A, 0x87, 0x46, 0xEB, 0x8C, 0x0F, 0x9A, 0x48, 0x90, 0xF9, 0x9F, 0x76, 0x69, 0x03, 0x72,
            0x77, 0xEC, 0xC1, 0x42, 0x4C, 0x87, 0xDB, 0x0B, 0x3C, 0xD4, 0x74, 0xEF, 0xE5, 0x34, 0xE0, 0x32,
            0x45, 0xB0, 0xF8, 0xAB, 0xD5, 0x26, 0x21, 0xD7, 0xD2, 0x98, 0x54, 0x8F, 0x64, 0x88, 0x20, 0x2B,
            0x14, 0xE3, 0x82, 0xD5, 0x2A, 0x4B, 0x8F, 0x4E, 0x35, 0x20, 0x82, 0x7E, 0x1B, 0xFE, 0xFA, 0x2C,
            0x79, 0x6C, 0x6E, 0x66, 0x94, 0xBB, 0x0A, 0xEB, 0xBA, 0xD9, 0x70, 0x61, 0xE9, 0x47, 0xB5, 0x82,
            0xFC, 0x18, 0x3C, 0x66, 0x3A, 0x09, 0x2E, 0x1F, 0x61, 0x74, 0xCA, 0xCB, 0xF6, 0x7A, 0x52, 0x37,
            0x1D, 0xAC, 0x8D, 0x63, 0x69, 0x84, 0x8E, 0xC7, 0x70, 0x59, 0xDD, 0x2D, 0x91, 0x1E, 0xF7, 0xB1,
            0x56, 0xED, 0x7A, 0x06, 0x9D, 0x5B, 0x33, 0x15, 0xDD, 0x31, 0xD0, 0xE6, 0x16, 0x07, 0x9B, 0xA5,
            0x94, 0x06, 0x7D, 0xC1, 0xE9, 0xD6, 0xC8, 0xAF, 0xB4, 0x1E, 0x2D, 0x88, 0x06, 0xA7, 0x63, 0xB8,
            0xCF, 0xC8, 0xA2, 0x6E, 0x84, 0xB3, 0x8D, 0xE5, 0x47, 0xE6, 0x13, 0x63, 0x8E, 0xD1, 0x7F, 0xD4,
            0x81, 0x44, 0x38, 0xBF
        };

        public static readonly byte[] TEST = {
            0x07, 0x02, 0x00, 0x00, 0x00, 0xA4, 0x00, 0x00, 0x52, 0x53, 0x41, 0x32, 0x00, 0x04, 0x00, 0x00,
            0x01, 0x00, 0x01, 0x00, 0x0F, 0xBE, 0x77, 0xB8, 0xDD, 0x54, 0x36, 0xDD, 0x67, 0xD4, 0x17, 0x66,
            0xC4, 0x13, 0xD1, 0x3F, 0x1E, 0x16, 0x0C, 0x16, 0x35, 0xAB, 0x6D, 0x3D, 0x34, 0x51, 0xED, 0x3F,
            0x57, 0x14, 0xB6, 0xB7, 0x08, 0xE9, 0xD9, 0x7A, 0x80, 0xB3, 0x5F, 0x9B, 0x3A, 0xFD, 0x9E, 0x37,
            0x3A, 0x53, 0x72, 0x67, 0x92, 0x60, 0xC3, 0xEF, 0xB5, 0x8E, 0x1E, 0xCF, 0x9D, 0x9C, 0xD3, 0x90,
            0xE5, 0xDD, 0xF4, 0xDB, 0xF3, 0xD6, 0x65, 0xB3, 0xC1, 0xBD, 0x69, 0xE1, 0x76, 0x95, 0xD9, 0x37,
            0xB8, 0x5E, 0xCA, 0x3D, 0x98, 0xFC, 0x50, 0x5C, 0x98, 0xAE, 0xE3, 0x7C, 0x4C, 0x27, 0xC3, 0xD0,
            0xCE, 0x78, 0x06, 0x51, 0x68, 0x23, 0xE6, 0x70, 0xF8, 0x7C, 0xAE, 0x36, 0xBE, 0x41, 0x57, 0xE2,
            0xC3, 0x2D, 0xAF, 0x21, 0xB1, 0xB3, 0x15, 0x81, 0x19, 0x26, 0x6B, 0x10, 0xB3, 0xE9, 0xD1, 0x45,
            0x21, 0x77, 0x9C, 0xF6, 0xE1, 0xDD, 0xB6, 0x78, 0x9D, 0x1D, 0x32, 0x61, 0xBC, 0x2B, 0xDB, 0x86,
            0xFB, 0x07, 0x24, 0x10, 0x19, 0x4F, 0x09, 0x6D, 0x03, 0x90, 0xD4, 0x5E, 0x30, 0x85, 0xC5, 0x58,
            0x7E, 0x5D, 0xAE, 0x9F, 0x64, 0x93, 0x04, 0x82, 0x09, 0x0E, 0x1C, 0x66, 0xA8, 0x95, 0x91, 0x51,
            0xB2, 0xED, 0x9A, 0x75, 0x04, 0x87, 0x50, 0xAC, 0xCC, 0x20, 0x06, 0x45, 0xB9, 0x7B, 0x42, 0x53,
            0x9A, 0xD1, 0x29, 0xFC, 0xEF, 0xB9, 0x47, 0x16, 0x75, 0x69, 0x05, 0x87, 0x2B, 0xCB, 0x54, 0x9C,
            0x21, 0x2D, 0x50, 0x8E, 0x12, 0xDE, 0xD3, 0x6B, 0xEC, 0x92, 0xA1, 0xB1, 0xE9, 0x4B, 0xBF, 0x6B,
            0x9A, 0x38, 0xC7, 0x13, 0xFA, 0x78, 0xA1, 0x3C, 0x1E, 0xBB, 0x38, 0x31, 0xBB, 0x0C, 0x9F, 0x70,
            0x1A, 0x31, 0x00, 0xD7, 0x5A, 0xA5, 0x84, 0x24, 0x89, 0x80, 0xF5, 0x88, 0xC2, 0x31, 0x18, 0xDC,
            0x53, 0x05, 0x5D, 0xFA, 0x81, 0xDC, 0xE1, 0xCE, 0xA4, 0xAA, 0xBA, 0x07, 0xDA, 0x28, 0x4F, 0x64,
            0x0E, 0x84, 0x9B, 0x06, 0xDE, 0xC8, 0x78, 0x66, 0x2F, 0x17, 0x25, 0xA8, 0x9C, 0x99, 0xFC, 0xBC,
            0x7D, 0x01, 0x42, 0xD7, 0x35, 0xBF, 0x19, 0xF6, 0x3F, 0x20, 0xD9, 0x98, 0x9B, 0x5D, 0xDD, 0x39,
            0xBE, 0x81, 0x00, 0x0B, 0xDE, 0x6F, 0x14, 0xCA, 0x7E, 0xF8, 0xC0, 0x26, 0xA8, 0x1D, 0xD1, 0x16,
            0x88, 0x64, 0x87, 0x36, 0x45, 0x37, 0x50, 0xDA, 0x6C, 0xEB, 0x85, 0xB5, 0x43, 0x29, 0x88, 0x6F,
            0x2F, 0xFE, 0x8D, 0x12, 0x8B, 0x72, 0xB7, 0x5A, 0xCB, 0x66, 0xC2, 0x2E, 0x1D, 0x7D, 0x42, 0xA6,
            0xF4, 0xFE, 0x26, 0x5D, 0x54, 0x9E, 0x77, 0x1D, 0x97, 0xC2, 0xF3, 0xFD, 0x60, 0xB3, 0x22, 0x88,
            0xCA, 0x27, 0x99, 0xDF, 0xC8, 0xB1, 0xD7, 0xC6, 0x54, 0xA6, 0x50, 0xB9, 0x54, 0xF5, 0xDE, 0xFE,
            0xE1, 0x81, 0xA2, 0xBE, 0x81, 0x9F, 0x48, 0xFF, 0x2F, 0xB8, 0xA4, 0xB3, 0x17, 0xD8, 0xC1, 0xB9,
            0x5D, 0x21, 0x3D, 0xA2, 0xED, 0x1C, 0x96, 0x66, 0xEE, 0x1F, 0x47, 0xCF, 0x62, 0xFA, 0xD6, 0xC1,
            0x87, 0x5B, 0xC4, 0xE5, 0xD9, 0x08, 0x38, 0x22, 0xFA, 0x21, 0xBD, 0xF2, 0x88, 0xDA, 0xE2, 0x24,
            0x25, 0x1F, 0xF1, 0x0B, 0x2D, 0xAE, 0x04, 0xBE, 0xA6, 0x7F, 0x75, 0x8C, 0xD9, 0x97, 0xE1, 0xCA,
            0x35, 0xB9, 0xFC, 0x6F, 0x01, 0x68, 0x11, 0xD3, 0x68, 0x32, 0xD0, 0xC1, 0x69, 0xA3, 0xCF, 0x9B,
            0x10, 0xE4, 0x69, 0xA7, 0xCF, 0xE1, 0xFE, 0x2A, 0x07, 0x9E, 0xC1, 0x37, 0x84, 0x68, 0xE5, 0xC5,
            0xAB, 0x25, 0xEC, 0x7D, 0x7D, 0x74, 0x6A, 0xD1, 0xD5, 0x4D, 0xD7, 0xE1, 0x7D, 0xDE, 0x30, 0x4B,
            0xE6, 0x5D, 0xCD, 0x91, 0x59, 0xF6, 0x80, 0xFD, 0xC6, 0x3C, 0xDD, 0x94, 0x7F, 0x15, 0x9D, 0xEF,
            0x2F, 0x00, 0x62, 0xD7, 0xDA, 0xB9, 0xB3, 0xD9, 0x8D, 0xE8, 0xD7, 0x3C, 0x96, 0x45, 0x5D, 0x1E,
            0x50, 0xFB, 0xAA, 0x43, 0xD3, 0x47, 0x77, 0x81, 0xE9, 0x67, 0xE4, 0xFE, 0xDF, 0x42, 0x79, 0xCB,
            0xA7, 0xAD, 0x5D, 0x48, 0xF5, 0xB7, 0x74, 0x96, 0x12, 0x23, 0x06, 0x70, 0x42, 0x68, 0x7A, 0x44,
            0xFC, 0xA0, 0x31, 0x7F, 0x68, 0xCA, 0xA2, 0x14, 0x5D, 0xA3, 0xCF, 0x42, 0x23, 0xAB, 0x47, 0xF6,
            0xB2, 0xFC, 0x6D, 0xF1
        };
    }
}


// Crypto/PhysStoreCrypto.cs
namespace LibTSforge.Crypto
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;

    public static class PhysStoreCrypto
    {
        public static byte[] DecryptPhysicalStore(byte[] data, bool production)
        {
            byte[] rsaKey = production ? Keys.PRODUCTION : Keys.TEST;
            BinaryReader br = new BinaryReader(new MemoryStream(data));
            br.BaseStream.Seek(0x10, SeekOrigin.Begin);
            byte[] aesKeySig = br.ReadBytes(0x80);
            byte[] encAesKey = br.ReadBytes(0x80);

            if (CryptoUtils.RSAVerifySignature(rsaKey, encAesKey, aesKeySig))
            {
                byte[] aesKey = CryptoUtils.RSADecrypt(rsaKey, encAesKey);
                byte[] decData = CryptoUtils.AESDecrypt(br.ReadBytes((int)br.BaseStream.Length - 0x110), aesKey);
                byte[] hmacKey = decData.Take(0x10).ToArray();
                byte[] hmacSig = decData.Skip(0x10).Take(0x14).ToArray();
                byte[] psData = decData.Skip(0x28).ToArray();

                if (!CryptoUtils.HMACVerify(hmacKey, psData, hmacSig))
                {
                    Logger.WriteLine("Warning: Failed to verify HMAC. Physical store is either corrupt or in Vista format.");
                }

                return psData;
            }

            throw new Exception("Failed to decrypt physical store.");
        }

        public static byte[] EncryptPhysicalStore(byte[] data, bool production, PSVersion version)
        {
            Dictionary<PSVersion, int> versionTable = new Dictionary<PSVersion, int>
            {
                {PSVersion.Win7, 5},
                {PSVersion.Win8, 1},
                {PSVersion.WinBlue, 2},
                {PSVersion.WinModern, 3}
            };

            byte[] rsaKey = production ? Keys.PRODUCTION : Keys.TEST;

            byte[] aesKey = Encoding.UTF8.GetBytes("massgrave.dev :3");
            byte[] hmacKey = CryptoUtils.GenerateRandomKey(0x10);

            byte[] encAesKey = CryptoUtils.RSAEncrypt(rsaKey, aesKey);
            byte[] aesKeySig = CryptoUtils.RSASign(rsaKey, encAesKey);
            byte[] hmacSig = CryptoUtils.HMACSign(hmacKey, data);

            byte[] decData = new byte[] { };
            decData = decData.Concat(hmacKey).Concat(hmacSig).Concat(BitConverter.GetBytes(0)).Concat(data).ToArray();
            byte[] encData = CryptoUtils.AESEncrypt(decData, aesKey);

            BinaryWriter bw = new BinaryWriter(new MemoryStream());
            bw.Write(versionTable[version]);
            bw.Write(Encoding.UTF8.GetBytes("UNTRUSTSTORE"));
            bw.Write(aesKeySig);
            bw.Write(encAesKey);
            bw.Write(encData);

            return bw.GetBytes();
        }
    }
}


// Modifiers/GenPKeyInstall.cs
namespace LibTSforge.Modifiers
{
    using System;
    using System.IO;
    using Microsoft.Win32;
    using LibTSforge.PhysicalStore;
    using LibTSforge.SPP;
    using LibTSforge.TokenStore;

    public static class GenPKeyInstall
    {
        private static void WritePkey2005RegistryValues(PSVersion version, ProductKey pkey)
        {
            Logger.WriteLine("Writing registry data for Windows product key...");
            Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductId", pkey.GetPid2());
            Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "DigitalProductId", pkey.GetPid3());
            Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "DigitalProductId4", pkey.GetPid4());

            if (Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Registration", "ProductId", null) != null)
            {
                Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Registration", "ProductId", pkey.GetPid2());
                Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Registration", "DigitalProductId", pkey.GetPid3());
                Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Registration", "DigitalProductId4", pkey.GetPid4());
            }

            if (pkey.Channel == "Volume:CSVLK" && version != PSVersion.Win7)
            {
                Registry.SetValue(@"HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform", "KmsHostConfig", 1);
            }
        }

        public static void InstallGenPKey(PSVersion version, bool production, Guid actId)
        {
            if (actId == Guid.Empty) throw new ArgumentException("Activation ID must be specified for generated product key install.");

            PKeyConfig pkc = new PKeyConfig();
            
            try
            {
                pkc.LoadConfig(actId);
            }
            catch (ArgumentException)
            {
                pkc.LoadAllConfigs(SLApi.GetAppId(actId));
            }

            ProductConfig config;
            pkc.Products.TryGetValue(actId, out config);

            if (config == null) throw new ArgumentException("Activation ID " + actId + " not found in PKeyConfig.");

            ProductKey pkey = config.GetRandomKey();

            Guid instPkeyId = SLApi.GetInstalledPkeyID(actId);
            if (instPkeyId != Guid.Empty) SLApi.UninstallProductKey(instPkeyId);

            if (pkey.Algorithm == PKeyAlgorithm.PKEY2009)
            {
                uint status = SLApi.InstallProductKey(pkey);
                Logger.WriteLine(string.Format("Installing generated product key {0} status {1:X}", pkey.ToString(), status));

                if ((int)status < 0)
                {
                    throw new ApplicationException("Failed to install generated product key.");
                }

                Logger.WriteLine("Successfully deposited generated product key.");
                return;
            }

            Logger.WriteLine("Key range is PKEY2005, creating fake key data...");

            if (pkey.Channel == "Volume:GVLK" && version == PSVersion.Win7) throw new NotSupportedException("Fake GVLK generation is not supported on Windows 7.");

            VariableBag pkb = new VariableBag();
            pkb.Blocks.AddRange(new CRCBlock[]
            {
                new CRCBlock
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingProductKey",
                    ValueAsStr = pkey.ToString()
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingMPC",
                    ValueAsStr = pkey.GetMPC()
                },
                new CRCBlock {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingPid2",
                    ValueAsStr = pkey.GetPid2()
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingPid3",
                    Value = pkey.GetPid3()
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingPid4",
                    Value = pkey.GetPid4()
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyChannelId",
                    ValueAsStr = pkey.Channel
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingEditionId",
                    ValueAsStr = pkey.Edition
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = (version == PSVersion.Win7) ? "SppPkeyShortAuthenticator" : "SppPkeyPhoneActivationData",
                    Value = pkey.GetPhoneData(version)
                },
                new CRCBlock
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingMiscData",
                    Value = new byte[] { }
                }
            });

            Guid appId = SLApi.GetAppId(actId);
            string pkeyId = pkey.GetPkeyId().ToString();
            bool isAddon = SLApi.IsAddon(actId);
            string currEdition = SLApi.GetMetaStr(actId, "Family");

            if (appId == SLApi.WINDOWS_APP_ID && !isAddon)
            {
                SLApi.UninstallAllProductKeys(appId);
            }

            Utils.KillSPP();

            using (IPhysicalStore ps = Utils.GetStore(version, production))
            {
                using (ITokenStore tks = Utils.GetTokenStore(version))
                {
                    Logger.WriteLine("Writing to physical store and token store...");

                    string suffix = (version == PSVersion.Win8 || version == PSVersion.WinBlue || version == PSVersion.WinModern) ? "_--" : "";
                    string metSuffix = suffix + "_met";

                    if (appId == SLApi.WINDOWS_APP_ID && !isAddon)
                    {
                        string edTokName = "msft:spp/token/windows/productkeyid/" + currEdition;

                        TokenMeta edToken = tks.GetMetaEntry(edTokName);
                        edToken.Data["windowsComponentEditionPkeyId"] = pkeyId;
                        edToken.Data["windowsComponentEditionSkuId"] = actId.ToString();
                        tks.SetEntry(edTokName, "xml", edToken.Serialize());

                        WritePkey2005RegistryValues(version, pkey);
                    }

                    string uriMapName = "msft:spp/token/PKeyIdUriMapper" + metSuffix;
                    TokenMeta uriMap = tks.GetMetaEntry(uriMapName);
                    uriMap.Data[pkeyId] = pkey.GetAlgoUri();
                    tks.SetEntry(uriMapName, "xml", uriMap.Serialize());

                    string skuMetaName = actId.ToString() + metSuffix;
                    TokenMeta skuMeta = tks.GetMetaEntry(skuMetaName);

                    foreach (string k in skuMeta.Data.Keys)
                    {
                        if (k.StartsWith("pkeyId_"))
                        {
                            skuMeta.Data.Remove(k);
                            break;
                        }
                    }

                    skuMeta.Data["pkeyId"] = pkeyId;
                    skuMeta.Data["pkeyIdList"] = pkeyId;
                    tks.SetEntry(skuMetaName, "xml", skuMeta.Serialize());

                    string psKey = string.Format("SPPSVC\\{0}\\{1}", appId, actId);
                    ps.DeleteBlock(psKey, pkeyId);
                    ps.AddBlock(new PSBlock
                    {
                        Type = BlockType.NAMED,
                        Flags = (version == PSVersion.WinModern) ? (uint)0x402 : 0x2,
                        KeyAsStr = psKey,
                        ValueAsStr = pkeyId,
                        Data = pkb.Serialize()
                    });

                    string cachePath = Utils.GetTokensPath(version).Replace("tokens.dat", @"cache\cache.dat");
                    if (File.Exists(cachePath)) File.Delete(cachePath);
                }
            }

            SLApi.RefreshTrustedTime(actId);
            Logger.WriteLine("Successfully deposited fake product key.");
        }
    }
}


// Modifiers/GracePeriodReset.cs
namespace LibTSforge.Modifiers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using LibTSforge.PhysicalStore;

    public static class GracePeriodReset
    {
        public static void Reset(PSVersion version, bool production)
        {
            Utils.KillSPP();
            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = Utils.GetStore(version, production))
            {
                string value = "msft:sl/timer";
                List<PSBlock> blocks = store.FindBlocks(value).ToList();

                foreach (PSBlock block in blocks)
                {
                    store.DeleteBlock(block.KeyAsStr, block.ValueAsStr);
                }
            }

            Logger.WriteLine("Successfully reset all grace and evaluation period timers.");
        }
    }
}


// Modifiers/KeyChangeLockDelete.cs
namespace LibTSforge.Modifiers
{
    using System.Collections.Generic;
    using System.Linq;
    using LibTSforge.PhysicalStore;
    using LibTSforge;
    public static class KeyChangeLockDelete
    {
        public static void Delete(PSVersion version, bool production)
        {
            Utils.KillSPP();
            Logger.WriteLine("Writing TrustedStore data...");
            using (IPhysicalStore store = Utils.GetStore(version, production))
            {
                List<string> values = new List<string>
                {
                    "msft:spp/timebased/AB",
                    "msft:spp/timebased/CD"
                };
                List<PSBlock> blocks = new List<PSBlock>();
                foreach (string value in values)
                {
                    blocks.AddRange(store.FindBlocks(value).ToList());
                }
                foreach (PSBlock block in blocks)
                {
                    store.DeleteBlock(block.KeyAsStr, block.ValueAsStr);
                }
            }
            Logger.WriteLine("Successfully removed the key change lock.");
        }
    }
}


// Modifiers/KMSHostCharge.cs
namespace LibTSforge.Modifiers
{
    using System;
    using System.IO;
    using LibTSforge.PhysicalStore;
    using LibTSforge.SPP;

    public static class KMSHostCharge
    {
        public static void Charge(PSVersion version, Guid actId, bool production)
        {
            if (actId == Guid.Empty)
            {
                actId = SLApi.GetDefaultActivationID(SLApi.WINDOWS_APP_ID, true);

                if (actId == Guid.Empty)
                {
                    throw new NotSupportedException("No applicable activation IDs found.");
                }
            }

            if (SLApi.GetPKeyChannel(SLApi.GetInstalledPkeyID(actId)) != "Volume:CSVLK")
            {
                throw new NotSupportedException("Non-Volume:CSVLK product key installed.");
            }

            Guid appId = SLApi.GetAppId(actId);
            int totalClients = 50;
            int currClients = 25;
            byte[] hwidBlock = Constants.UniversalHWIDBlock;
            string key = string.Format("SPPSVC\\{0}", appId);
            long ldapTimestamp = DateTime.Now.ToFileTime();

            BinaryWriter writer = new BinaryWriter(new MemoryStream());

            for (int i = 0; i < currClients; i++)
            {
                writer.Write(ldapTimestamp - (10 * (i + 1)));
                writer.Write(Guid.NewGuid().ToByteArray());
            }

            byte[] cmidGuids = writer.GetBytes();

            writer = new BinaryWriter(new MemoryStream());

            writer.Write(new byte[40]);

            writer.Seek(4, SeekOrigin.Begin);
            writer.Write((byte)currClients);

            writer.Seek(24, SeekOrigin.Begin);
            writer.Write((byte)currClients);
            byte[] reqCounts = writer.GetBytes();

            Utils.KillSPP();

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = Utils.GetStore(version, production))
            {
                VariableBag kmsCountData = new VariableBag();
                kmsCountData.Blocks.AddRange(new CRCBlock[]
                {
                    new CRCBlock
                    {
                        DataType = CRCBlockType.BINARY,
                        KeyAsStr = "SppBindingLicenseData",
                        Value = hwidBlock
                    },
                    new CRCBlock
                    {
                        DataType = CRCBlockType.UINT,
                        Key = new byte[] { },
                        ValueAsInt = (uint)totalClients
                    },
                    new CRCBlock
                    {
                        DataType = CRCBlockType.UINT,
                        Key = new byte[] { },
                        ValueAsInt = 1051200000
                    },
                    new CRCBlock
                    {
                        DataType = CRCBlockType.UINT,
                        Key = new byte[] { },
                        ValueAsInt = (uint)currClients
                    },
                    new CRCBlock
                    {
                        DataType = CRCBlockType.BINARY,
                        Key = new byte[] { },
                        Value = cmidGuids
                    },
                    new CRCBlock
                    {
                        DataType = CRCBlockType.BINARY,
                        Key = new byte[] { },
                        Value = reqCounts
                    }
                });

                byte[] kmsChargeData = kmsCountData.Serialize();
                string countVal = string.Format("msft:spp/kms/host/2.0/store/counters/{0}", appId);

                store.DeleteBlock(key, countVal);
                store.AddBlock(new PSBlock
                {
                    Type = BlockType.NAMED,
                    Flags = (version == PSVersion.WinModern) ? (uint)0x400 : 0,
                    KeyAsStr = key,
                    ValueAsStr = countVal,
                    Data = kmsChargeData
                });

                Logger.WriteLine(string.Format("Set charge count to {0} successfully.", currClients));
            }
        }
    }
}


// Modifiers/RearmReset.cs
namespace LibTSforge.Modifiers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using LibTSforge.PhysicalStore;

    public static class RearmReset
    {
        public static void Reset(PSVersion version, bool production)
        {
            Utils.KillSPP();

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = Utils.GetStore(version, production))
            {
                List<PSBlock> blocks;

                if (version == PSVersion.Win7)
                {
                    blocks = store.FindBlocks(0xA0000).ToList();
                }
                else
                {
                    blocks = store.FindBlocks("__##USERSEP-RESERVED##__$$REARM-COUNT$$").ToList();
                }

                foreach (PSBlock block in blocks)
                {
                    if (version == PSVersion.Win7)
                    {
                        store.SetBlock(block.KeyAsStr, block.ValueAsInt, new byte[8]);
                    }
                    else
                    {
                        store.SetBlock(block.KeyAsStr, block.ValueAsStr, new byte[8]);
                    }
                }

                Logger.WriteLine("Successfully reset all rearm counters.");
            }
        }
    }
}


// Modifiers/TamperedFlagsDelete.cs
namespace LibTSforge.Modifiers
{
    using System;
    using System.Linq;
    using LibTSforge.PhysicalStore;

    public static class TamperedFlagsDelete
    {
        public static void DeleteTamperFlags(PSVersion version, bool production)
        {
            Utils.KillSPP();

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = Utils.GetStore(version, production))
            {
                if (version != PSVersion.Win7)
                {
                    string recreatedFlag = "__##USERSEP-RESERVED##__$$RECREATED-FLAG$$";
                    string recoveredFlag = "__##USERSEP-RESERVED##__$$RECOVERED-FLAG$$";

                    DeleteFlag(store, recreatedFlag);
                    DeleteFlag(store, recoveredFlag);
                }
                else
                {
                    SetFlag(store, 0xA0001);
                }

                Logger.WriteLine("Successfully cleared the tamper state.");
            }
        }

        private static void DeleteFlag(IPhysicalStore store, string flag)
        {
            store.FindBlocks(flag).ToList().ForEach(block => store.DeleteBlock(block.KeyAsStr, block.ValueAsStr));
        }

        private static void SetFlag(IPhysicalStore store, uint flag)
        {
            store.FindBlocks(flag).ToList().ForEach(block => store.SetBlock(block.KeyAsStr, block.ValueAsInt, new byte[8]));
        }
    }
}


// Modifiers/UniqueIdDelete.cs
namespace LibTSforge.Modifiers
{
    using System;
    using LibTSforge.PhysicalStore;
    using LibTSforge.SPP;

    public static class UniqueIdDelete
    {
        public static void DeleteUniqueId(PSVersion version, bool production, Guid actId)
        {
            Guid appId;

            if (actId == Guid.Empty)
            {
                appId = SLApi.WINDOWS_APP_ID;
                actId = SLApi.GetDefaultActivationID(appId, true);

                if (actId == Guid.Empty)
                {
                    throw new Exception("No applicable activation IDs found.");
                }
            }
            else
            {
                appId = SLApi.GetAppId(actId);
            }

            string instId = SLApi.GetInstallationID(actId);
            Guid pkeyId = SLApi.GetInstalledPkeyID(actId);

            Utils.KillSPP();

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = Utils.GetStore(version, production))
            {
                string key = string.Format("SPPSVC\\{0}\\{1}", appId, actId);
                PSBlock keyBlock = store.GetBlock(key, pkeyId.ToString());

                if (keyBlock == null)
                {
                    throw new Exception("No product key found.");
                }

                VariableBag pkb = new VariableBag(keyBlock.Data);

                pkb.DeleteBlock("SppPkeyUniqueIdToken");

                store.SetBlock(key, pkeyId.ToString(), pkb.Serialize());
            }

            Logger.WriteLine("Successfully removed Unique ID for product key ID " + pkeyId);
        }
    }
}


// Activators/ZeroCID.cs
namespace LibTSforge.Activators
{
    using System;
    using System.IO;
    using LibTSforge.Crypto;
    using LibTSforge.PhysicalStore;
    using LibTSforge.SPP;

    public static class ZeroCID
    {
        public static void Deposit(Guid actId, string instId)
        {
            uint status = SLApi.DepositConfirmationID(actId, instId, Constants.ZeroCID);
            Logger.WriteLine(string.Format("Depositing fake CID status {0:X}", status));

            if (status != 0)
            {
                throw new InvalidOperationException("Failed to deposit fake CID.");
            }
        }

        public static void Activate(PSVersion version, bool production, Guid actId)
        {
            Guid appId;

            if (actId == Guid.Empty)
            {
                appId = SLApi.WINDOWS_APP_ID;
                actId = SLApi.GetDefaultActivationID(appId, false);

                if (actId == Guid.Empty)
                {
                    throw new NotSupportedException("No applicable activation IDs found.");
                }
            }
            else
            {
                appId = SLApi.GetAppId(actId);
            }

            if (!SLApi.IsPhoneActivatable(actId))
            {
                throw new NotSupportedException("Phone license is unavailable for this product.");
            }

            string instId = SLApi.GetInstallationID(actId);
            Guid pkeyId = SLApi.GetInstalledPkeyID(actId);

            if (version == PSVersion.Win7)
            {
                Deposit(actId, instId);
            }

            Utils.KillSPP();

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = Utils.GetStore(version, production))
            {
                byte[] hwidBlock = Constants.UniversalHWIDBlock;

                Logger.WriteLine("Activation ID: " + actId);
                Logger.WriteLine("Installation ID: " + instId);
                Logger.WriteLine("Product Key ID: " + pkeyId);

                byte[] iidHash;

                if (version == PSVersion.Win7)
                {
                    iidHash = CryptoUtils.SHA256Hash(Utils.EncodeString(instId));
                }
                else
                {
                    iidHash = CryptoUtils.SHA256Hash(Utils.EncodeString(instId + '\0' + Constants.ZeroCID));
                }

                string key = string.Format("SPPSVC\\{0}\\{1}", appId, actId);
                PSBlock keyBlock = store.GetBlock(key, pkeyId.ToString());

                if (keyBlock == null)
                {
                    throw new InvalidDataException("Failed to get product key data for activation ID " + actId + ".");
                }

                VariableBag pkb = new VariableBag(keyBlock.Data);

                byte[] pkeyData;

                if (version == PSVersion.Win7)
                {
                    pkeyData = pkb.GetBlock("SppPkeyShortAuthenticator").Value;
                }
                else
                {
                    pkeyData = pkb.GetBlock("SppPkeyPhoneActivationData").Value;
                }

                pkb.DeleteBlock("SppPkeyVirtual");
                store.SetBlock(key, pkeyId.ToString(), pkb.Serialize());

                BinaryWriter writer = new BinaryWriter(new MemoryStream());
                writer.Write(0x20);
                writer.Write(iidHash);
                writer.Write(hwidBlock.Length);
                writer.Write(hwidBlock);
                byte[] tsHwidData = writer.GetBytes();

                writer = new BinaryWriter(new MemoryStream());
                writer.Write(0x20);
                writer.Write(iidHash);
                writer.Write(pkeyData.Length);
                writer.Write(pkeyData);
                byte[] tsPkeyInfoData = writer.GetBytes();

                store.AddBlocks(new PSBlock[] {
                    new PSBlock
                    {
                        Type = BlockType.NAMED,
                        Flags = 0,
                        KeyAsStr = key,
                        ValueAsStr = "msft:Windows/7.0/Phone/Cached/HwidBlock/" + pkeyId,
                        Data = tsHwidData
                    }, 
                    new PSBlock
                    {
                        Type = BlockType.NAMED,
                        Flags = 0,
                        KeyAsStr = key,
                        ValueAsStr = "msft:Windows/7.0/Phone/Cached/PKeyInfo/" + pkeyId,
                        Data = tsPkeyInfoData
                    }
                });
            }

            if (version != PSVersion.Win7)
            {
                Deposit(actId, instId);
            }

            SLApi.RefreshLicenseStatus();
            SLApi.FireStateChangedEvent(appId);
            Logger.WriteLine("Activated using ZeroCID successfully.");
        }
    }
}


// TokenStore/Common.cs
namespace LibTSforge.TokenStore
{
    using System.Collections.Generic;
    using System.IO;

    public class TokenEntry
    {
        public string Name;
        public string Extension;
        public byte[] Data;
        public bool Populated;
    }

    public class TokenMeta
    {
        public string Name;
        public Dictionary<string, string> Data = new Dictionary<string, string>();

        public byte[] Serialize()
        {
            BinaryWriter writer = new BinaryWriter(new MemoryStream());
            writer.Write(1);
            byte[] nameBytes = Utils.EncodeString(Name);
            writer.Write(nameBytes.Length);
            writer.Write(nameBytes);

            foreach (KeyValuePair<string, string> kv in Data)
            {
                byte[] keyBytes = Utils.EncodeString(kv.Key);
                byte[] valueBytes = Utils.EncodeString(kv.Value);
                writer.Write(keyBytes.Length);
                writer.Write(valueBytes.Length);
                writer.Write(keyBytes);
                writer.Write(valueBytes);
            }

            return writer.GetBytes();
        }

        public void Deserialize(byte[] data)
        {
            BinaryReader reader = new BinaryReader(new MemoryStream(data));
            reader.ReadInt32();
            int nameLen = reader.ReadInt32();
            Name = reader.ReadNullTerminatedString(nameLen);

            while (reader.BaseStream.Position < data.Length - 0x8)
            {
                int keyLen = reader.ReadInt32();
                int valueLen = reader.ReadInt32();
                string key = reader.ReadNullTerminatedString(keyLen);
                string value = reader.ReadNullTerminatedString(valueLen);
                Data[key] = value;
            }
        }

        public TokenMeta(byte[] data)
        {
            Deserialize(data);
        }

        public TokenMeta()
        {

        }
    }
}


// TokenStore/ITokenStore.cs
namespace LibTSforge.TokenStore
{
    using System;

    public interface ITokenStore : IDisposable
    {
        void Deserialize();
        void Serialize();
        void AddEntry(TokenEntry entry);
        void AddEntries(TokenEntry[] entries);
        void DeleteEntry(string name, string ext);
        void DeleteUnpopEntry(string name, string ext);
        TokenEntry GetEntry(string name, string ext);
        TokenMeta GetMetaEntry(string name);
        void SetEntry(string name, string ext, byte[] data);
    }
}


// TokenStore/TokenStoreModern.cs
namespace LibTSforge.TokenStore
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using LibTSforge.Crypto;

    public class TokenStoreModern : ITokenStore
    {
        private static readonly uint VERSION = 3;
        private static readonly int ENTRY_SIZE = 0x9E;
        private static readonly int BLOCK_SIZE = 0x4020;
        private static readonly int ENTRIES_PER_BLOCK = BLOCK_SIZE / ENTRY_SIZE;
        private static readonly int BLOCK_PAD_SIZE = 0x66;

        private static readonly byte[] CONTS_HEADER = Enumerable.Repeat((byte)0x55, 0x20).ToArray();
        private static readonly byte[] CONTS_FOOTER = Enumerable.Repeat((byte)0xAA, 0x20).ToArray();

        private List<TokenEntry> Entries = new List<TokenEntry>();
        public FileStream TokensFile;

        public void Deserialize()
        {
            if (TokensFile.Length < BLOCK_SIZE) return;

            TokensFile.Seek(0x24, SeekOrigin.Begin);
            uint nextBlock = 0;

            BinaryReader reader = new BinaryReader(TokensFile);
            do
            {
                uint curOffset = reader.ReadUInt32();
                nextBlock = reader.ReadUInt32();

                for (int i = 0; i < ENTRIES_PER_BLOCK; i++)
                {
                    curOffset = reader.ReadUInt32();
                    bool populated = reader.ReadUInt32() == 1;
                    uint contentOffset = reader.ReadUInt32();
                    uint contentLength = reader.ReadUInt32();
                    uint allocLength = reader.ReadUInt32();
                    byte[] contentData = new byte[] { };

                    if (populated)
                    {
                        reader.BaseStream.Seek(contentOffset + 0x20, SeekOrigin.Begin);
                        uint dataLength = reader.ReadUInt32();

                        if (dataLength != contentLength)
                        {
                            throw new FormatException("Data length in tokens content is inconsistent with entry.");
                        }

                        reader.ReadBytes(0x20);
                        contentData = reader.ReadBytes((int)contentLength);
                    }

                    reader.BaseStream.Seek(curOffset + 0x14, SeekOrigin.Begin);

                    Entries.Add(new TokenEntry
                    {
                        Name = reader.ReadNullTerminatedString(0x82),
                        Extension = reader.ReadNullTerminatedString(0x8),
                        Data = contentData,
                        Populated = populated
                    });
                }

                reader.BaseStream.Seek(nextBlock, SeekOrigin.Begin);
            } while (nextBlock != 0);
        }

        public void Serialize()
        {
            MemoryStream tokens = new MemoryStream();

            using (BinaryWriter writer = new BinaryWriter(tokens))
            {
                writer.Write(VERSION);
                writer.Write(CONTS_HEADER);

                int curBlockOffset = (int)writer.BaseStream.Position;
                int curEntryOffset = curBlockOffset + 0x8;
                int curContsOffset = curBlockOffset + BLOCK_SIZE;

                for (int eIndex = 0; eIndex < ((Entries.Count / ENTRIES_PER_BLOCK) + 1) * ENTRIES_PER_BLOCK; eIndex++)
                {
                    TokenEntry entry;

                    if (eIndex < Entries.Count)
                    {
                        entry = Entries[eIndex];
                    }
                    else
                    {
                        entry = new TokenEntry
                        {
                            Name = "",
                            Extension = "",
                            Populated = false,
                            Data = new byte[] { }
                        };
                    }

                    writer.BaseStream.Seek(curBlockOffset, SeekOrigin.Begin);
                    writer.Write(curBlockOffset);
                    writer.Write(0);

                    writer.BaseStream.Seek(curEntryOffset, SeekOrigin.Begin);
                    writer.Write(curEntryOffset);
                    writer.Write(entry.Populated ? 1 : 0);
                    writer.Write(entry.Populated ? curContsOffset : 0);
                    writer.Write(entry.Populated ? entry.Data.Length : -1);
                    writer.Write(entry.Populated ? entry.Data.Length : -1);
                    writer.WriteFixedString16(entry.Name, 0x82);
                    writer.WriteFixedString16(entry.Extension, 0x8);
                    curEntryOffset = (int)writer.BaseStream.Position;

                    if (entry.Populated)
                    {
                        writer.BaseStream.Seek(curContsOffset, SeekOrigin.Begin);
                        writer.Write(CONTS_HEADER);
                        writer.Write(entry.Data.Length);
                        writer.Write(CryptoUtils.SHA256Hash(entry.Data));
                        writer.Write(entry.Data);
                        writer.Write(CONTS_FOOTER);
                        curContsOffset = (int)writer.BaseStream.Position;
                    }

                    if ((eIndex + 1) % ENTRIES_PER_BLOCK == 0 && eIndex != 0)
                    {
                        if (eIndex < Entries.Count)
                        {
                            writer.BaseStream.Seek(curBlockOffset + 0x4, SeekOrigin.Begin);
                            writer.Write(curContsOffset);
                        }

                        writer.BaseStream.Seek(curEntryOffset, SeekOrigin.Begin);
                        writer.WritePadding(BLOCK_PAD_SIZE);

                        writer.BaseStream.Seek(curBlockOffset, SeekOrigin.Begin);
                        byte[] blockHash;
                        byte[] blockData = new byte[BLOCK_SIZE - 0x20];

                        tokens.Read(blockData, 0, BLOCK_SIZE - 0x20);
                        blockHash = CryptoUtils.SHA256Hash(blockData);

                        writer.BaseStream.Seek(curBlockOffset + BLOCK_SIZE - 0x20, SeekOrigin.Begin);
                        writer.Write(blockHash);

                        curBlockOffset = curContsOffset;
                        curEntryOffset = curBlockOffset + 0x8;
                        curContsOffset = curBlockOffset + BLOCK_SIZE;
                    }
                }

                tokens.SetLength(curBlockOffset);
            }

            byte[] tokensData = tokens.ToArray();
            byte[] tokensHash = CryptoUtils.SHA256Hash(tokensData.Take(0x4).Concat(tokensData.Skip(0x24)).ToArray());

            tokens = new MemoryStream(tokensData);

            BinaryWriter tokWriter = new BinaryWriter(TokensFile);
            using (BinaryReader reader = new BinaryReader(tokens))
            {
                TokensFile.Seek(0, SeekOrigin.Begin);
                TokensFile.SetLength(tokens.Length);
                tokWriter.Write(reader.ReadBytes(0x4));
                reader.ReadBytes(0x20);
                tokWriter.Write(tokensHash);
                tokWriter.Write(reader.ReadBytes((int)reader.BaseStream.Length - 0x4));
            }
        }

        public void AddEntry(TokenEntry entry)
        {
            Entries.Add(entry);
        }

        public void AddEntries(TokenEntry[] entries)
        {
            Entries.AddRange(entries);
        }

        public void DeleteEntry(string name, string ext)
        {
            foreach (TokenEntry entry in Entries)
            {
                if (entry.Name == name && entry.Extension == ext)
                {
                    Entries.Remove(entry);
                    return;
                }
            }
        }

        public void DeleteUnpopEntry(string name, string ext)
        {
            List<TokenEntry> delEntries = new List<TokenEntry>();
            foreach (TokenEntry entry in Entries)
            {
                if (entry.Name == name && entry.Extension == ext && !entry.Populated)
                {
                    delEntries.Add(entry);
                }
            }

            Entries = Entries.Except(delEntries).ToList();
        }

        public TokenEntry GetEntry(string name, string ext)
        {
            foreach (TokenEntry entry in Entries)
            {
                if (entry.Name == name && entry.Extension == ext)
                {
                    if (!entry.Populated) continue;
                    return entry;
                }
            }

            return null;
        }

        public TokenMeta GetMetaEntry(string name)
        {
            DeleteUnpopEntry(name, "xml");
            TokenEntry entry = GetEntry(name, "xml");
            TokenMeta meta;

            if (entry == null)
            {
                meta = new TokenMeta
                {
                    Name = name
                };
            }
            else
            {
                meta = new TokenMeta(entry.Data);
            }

            return meta;
        }

        public void SetEntry(string name, string ext, byte[] data)
        {
            for (int i = 0; i < Entries.Count; i++)
            {
                TokenEntry entry = Entries[i];

                if (entry.Name == name && entry.Extension == ext && entry.Populated)
                {
                    entry.Data = data;
                    Entries[i] = entry;
                    return;
                }
            }

            Entries.Add(new TokenEntry
            {
                Populated = true,
                Name = name,
                Extension = ext,
                Data = data
            });
        }

        public TokenStoreModern(string tokensPath)
        {
            TokensFile = File.Open(tokensPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            Deserialize();
        }

        public TokenStoreModern()
        {

        }

        public void Dispose()
        {
            Serialize();
            TokensFile.Close();
        }
    }
}


// PhysicalStore/Common.cs
namespace LibTSforge.PhysicalStore
{
    using System.Runtime.InteropServices;

    public enum BlockType : uint
    {
        NONE,
        NAMED,
        ATTRIBUTE,
        TIMER
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct Timer
    {
        public ulong Unknown;
        public ulong Time1;
        public ulong Time2;
        public ulong Expiry;
    }
}


// PhysicalStore/IPhysicalStore.cs
namespace LibTSforge.PhysicalStore
{
    using System;
    using System.Collections.Generic;

    public class PSBlock
    {
        public BlockType Type;
        public uint Flags;
        public uint Unknown = 0;
        public byte[] Key;
        public string KeyAsStr
        {
            get
            {
                return Utils.DecodeString(Key);
            }
            set
            {
                Key = Utils.EncodeString(value);
            }
        }
        public byte[] Value;
        public string ValueAsStr
        {
            get
            {
                return Utils.DecodeString(Value);
            }
            set
            {
                Value = Utils.EncodeString(value);
            }
        }
        public uint ValueAsInt
        {
            get
            {
                return BitConverter.ToUInt32(Value, 0);
            }
            set
            {
                Value = BitConverter.GetBytes(value);
            }
        }
        public byte[] Data;
        public string DataAsStr
        {
            get
            {
                return Utils.DecodeString(Data);
            }
            set
            {
                Data = Utils.EncodeString(value);
            }
        }
        public uint DataAsInt
        {
            get
            {
                return BitConverter.ToUInt32(Data, 0);
            }
            set
            {
                Data = BitConverter.GetBytes(value);
            }
        }
    }

    public interface IPhysicalStore : IDisposable
    {
        PSBlock GetBlock(string key, string value);
        PSBlock GetBlock(string key, uint value);
        void AddBlock(PSBlock block);
        void AddBlocks(IEnumerable<PSBlock> blocks);
        void SetBlock(string key, string value, byte[] data);
        void SetBlock(string key, string value, string data);
        void SetBlock(string key, string value, uint data);
        void SetBlock(string key, uint value, byte[] data);
        void SetBlock(string key, uint value, string data);
        void SetBlock(string key, uint value, uint data);
        void DeleteBlock(string key, string value);
        void DeleteBlock(string key, uint value);
        byte[] Serialize();
        void Deserialize(byte[] data);
        byte[] ReadRaw();
        void WriteRaw(byte[] data);
        IEnumerable<PSBlock> FindBlocks(string valueSearch);
        IEnumerable<PSBlock> FindBlocks(uint valueSearch);
    }
}


// PhysicalStore/PhysicalStoreModern.cs
namespace LibTSforge.PhysicalStore
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using LibTSforge.Crypto;

    public class ModernBlock
    {
        public BlockType Type;
        public uint Flags;
        public uint Unknown;
        public byte[] Value;
        public string ValueAsStr
        {
            get
            {
                return Utils.DecodeString(Value);
            }
            set
            {
                Value = Utils.EncodeString(value);
            }
        }
        public uint ValueAsInt
        {
            get
            {
                return BitConverter.ToUInt32(Value, 0);
            }
            set
            {
                Value = BitConverter.GetBytes(value);
            }
        }
        public byte[] Data;
        public string DataAsStr
        {
            get
            {
                return Utils.DecodeString(Data);
            }
            set
            {
                Data = Utils.EncodeString(value);
            }
        }
        public uint DataAsInt
        {
            get
            {
                return BitConverter.ToUInt32(Data, 0);
            }
            set
            {
                Data = BitConverter.GetBytes(value);
            }
        }

        public void Encode(BinaryWriter writer)
        {
            writer.Write((uint)Type);
            writer.Write(Flags);
            writer.Write((uint)Value.Length);
            writer.Write((uint)Data.Length);
            writer.Write(Unknown);
            writer.Write(Value);
            writer.Write(Data);
        }

        public static ModernBlock Decode(BinaryReader reader)
        {
            uint type = reader.ReadUInt32();
            uint flags = reader.ReadUInt32();

            uint valueLen = reader.ReadUInt32();
            uint dataLen = reader.ReadUInt32();
            uint unk3 = reader.ReadUInt32();

            byte[] value = reader.ReadBytes((int)valueLen);
            byte[] data = reader.ReadBytes((int)dataLen);

            return new ModernBlock
            {
                Type = (BlockType)type,
                Flags = flags,
                Unknown = unk3,
                Value = value,
                Data = data,
            };
        }
    }

    public sealed class PhysicalStoreModern : IPhysicalStore
    {
        private byte[] PreHeaderBytes = new byte[] { };
        private readonly Dictionary<string, List<ModernBlock>> Data = new Dictionary<string, List<ModernBlock>>();
        private readonly FileStream TSFile;
        private readonly PSVersion Version;
        private readonly bool Production;

        public byte[] Serialize()
        {
            BinaryWriter writer = new BinaryWriter(new MemoryStream());
            writer.Write(PreHeaderBytes);
            writer.Write(Data.Keys.Count);

            foreach (string key in Data.Keys)
            {
                List<ModernBlock> blocks = Data[key];
                byte[] keyNameEnc = Utils.EncodeString(key);

                writer.Write(keyNameEnc.Length);
                writer.Write(keyNameEnc);
                writer.Write(blocks.Count);
                writer.Align(4);

                foreach (ModernBlock block in blocks)
                {
                    block.Encode(writer);
                    writer.Align(4);
                }
            }

            return writer.GetBytes();
        }

        public void Deserialize(byte[] data)
        {
            BinaryReader reader = new BinaryReader(new MemoryStream(data));
            PreHeaderBytes = reader.ReadBytes(8);

            while (reader.BaseStream.Position < data.Length - 0x4)
            {
                uint numKeys = reader.ReadUInt32();

                for (int i = 0; i < numKeys; i++)
                {
                    uint lenKeyName = reader.ReadUInt32();
                    string keyName = Utils.DecodeString(reader.ReadBytes((int)lenKeyName)); uint numValues = reader.ReadUInt32();

                    reader.Align(4);

                    Data[keyName] = new List<ModernBlock>();

                    for (int j = 0; j < numValues; j++)
                    {
                        Data[keyName].Add(ModernBlock.Decode(reader));
                        reader.Align(4);
                    }
                }
            }
        }

        public void AddBlock(PSBlock block)
        {
            if (!Data.ContainsKey(block.KeyAsStr))
            {
                Data[block.KeyAsStr] = new List<ModernBlock>();
            }

            Data[block.KeyAsStr].Add(new ModernBlock
            {
                Type = block.Type,
                Flags = block.Flags,
                Unknown = block.Unknown,
                Value = block.Value,
                Data = block.Data
            });
        }

        public void AddBlocks(IEnumerable<PSBlock> blocks)
        {
            foreach (PSBlock block in blocks)
            {
                AddBlock(block);
            }
        }

        public PSBlock GetBlock(string key, string value)
        {
            List<ModernBlock> blocks = Data[key];

            foreach (ModernBlock block in blocks)
            {
                if (block.ValueAsStr == value)
                {
                    return new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = Utils.EncodeString(key),
                        Value = block.Value,
                        Data = block.Data
                    };
                }
            }

            return null;
        }

        public PSBlock GetBlock(string key, uint value)
        {
            List<ModernBlock> blocks = Data[key];

            foreach (ModernBlock block in blocks)
            {
                if (block.ValueAsInt == value)
                {
                    return new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = Utils.EncodeString(key),
                        Value = block.Value,
                        Data = block.Data
                    };
                }
            }

            return null;
        }

        public void SetBlock(string key, string value, byte[] data)
        {
            List<ModernBlock> blocks = Data[key];

            for (int i = 0; i < blocks.Count; i++)
            {
                ModernBlock block = blocks[i];

                if (block.ValueAsStr == value)
                {
                    block.Data = data;
                    blocks[i] = block;
                    break;
                }
            }

            Data[key] = blocks;
        }

        public void SetBlock(string key, uint value, byte[] data)
        {
            List<ModernBlock> blocks = Data[key];

            for (int i = 0; i < blocks.Count; i++)
            {
                ModernBlock block = blocks[i];

                if (block.ValueAsInt == value)
                {
                    block.Data = data;
                    blocks[i] = block;
                    break;
                }
            }

            Data[key] = blocks;
        }

        public void SetBlock(string key, string value, string data)
        {
            SetBlock(key, value, Utils.EncodeString(data));
        }

        public void SetBlock(string key, string value, uint data)
        {
            SetBlock(key, value, BitConverter.GetBytes(data));
        }

        public void SetBlock(string key, uint value, string data)
        {
            SetBlock(key, value, Utils.EncodeString(data));
        }

        public void SetBlock(string key, uint value, uint data)
        {
            SetBlock(key, value, BitConverter.GetBytes(data));
        }

        public void DeleteBlock(string key, string value)
        {
            if (Data.ContainsKey(key))
            {
                List<ModernBlock> blocks = Data[key];

                foreach (ModernBlock block in blocks)
                {
                    if (block.ValueAsStr == value)
                    {
                        blocks.Remove(block);
                        break;
                    }
                }

                Data[key] = blocks;
            }
        }

        public void DeleteBlock(string key, uint value)
        {
            if (Data.ContainsKey(key))
            {
                List<ModernBlock> blocks = Data[key];

                foreach (ModernBlock block in blocks)
                {
                    if (block.ValueAsInt == value)
                    {
                        blocks.Remove(block);
                        break;
                    }
                }

                Data[key] = blocks;
            }
        }

        public PhysicalStoreModern(string tsPath, bool production, PSVersion version)
        {
            TSFile = File.Open(tsPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            Deserialize(PhysStoreCrypto.DecryptPhysicalStore(TSFile.ReadAllBytes(), production));
            TSFile.Seek(0, SeekOrigin.Begin);
            Version = version;
            Production = production;
        }

        public void Dispose()
        {
            if (TSFile.CanWrite)
            {
                byte[] data = PhysStoreCrypto.EncryptPhysicalStore(Serialize(), Production, Version);
                TSFile.SetLength(data.LongLength);
                TSFile.Seek(0, SeekOrigin.Begin);
                TSFile.WriteAllBytes(data);
                TSFile.Close();
            }
        }

        public byte[] ReadRaw()
        {
            byte[] data = PhysStoreCrypto.DecryptPhysicalStore(TSFile.ReadAllBytes(), Production);
            TSFile.Seek(0, SeekOrigin.Begin);
            return data;
        }

        public void WriteRaw(byte[] data)
        {
            byte[] encrData = PhysStoreCrypto.EncryptPhysicalStore(data, Production, Version);
            TSFile.SetLength(encrData.LongLength);
            TSFile.Seek(0, SeekOrigin.Begin);
            TSFile.WriteAllBytes(encrData);
            TSFile.Close();
        }

        public IEnumerable<PSBlock> FindBlocks(string valueSearch)
        {
            List<PSBlock> results = new List<PSBlock>();

            foreach (string key in Data.Keys)
            {
                List<ModernBlock> values = Data[key];

                foreach (ModernBlock block in values)
                {
                    if (block.ValueAsStr.Contains(valueSearch))
                    {
                        results.Add(new PSBlock
                        {
                            Type = block.Type,
                            Flags = block.Flags,
                            KeyAsStr = key,
                            Value = block.Value,
                            Data = block.Data
                        });
                    }
                }
            }

            return results;
        }

        public IEnumerable<PSBlock> FindBlocks(uint valueSearch)
        {
            List<PSBlock> results = new List<PSBlock>();

            foreach (string key in Data.Keys)
            {
                List<ModernBlock> values = Data[key];

                foreach (ModernBlock block in values)
                {
                    if (block.ValueAsInt == valueSearch)
                    {
                        results.Add(new PSBlock
                        {
                            Type = block.Type,
                            Flags = block.Flags,
                            KeyAsStr = key,
                            Value = block.Value,
                            Data = block.Data
                        });
                    }
                }
            }

            return results;
        }
    }
}


// PhysicalStore/PhysicalStoreWin7.cs
namespace LibTSforge.PhysicalStore
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using LibTSforge.Crypto;

    public class Win7Block
    {
        public BlockType Type;
        public uint Flags;
        public byte[] Key;
        public string KeyAsStr
        {
            get
            {
                return Utils.DecodeString(Key);
            }
            set
            {
                Key = Utils.EncodeString(value);
            }
        }
        public byte[] Value;
        public string ValueAsStr
        {
            get
            {
                return Utils.DecodeString(Value);
            }
            set
            {
                Value = Utils.EncodeString(value);
            }
        }
        public uint ValueAsInt
        {
            get
            {
                return BitConverter.ToUInt32(Value, 0);
            }
            set
            {
                Value = BitConverter.GetBytes(value);
            }
        }
        public byte[] Data;
        public string DataAsStr
        {
            get
            {
                return Utils.DecodeString(Data);
            }
            set
            {
                Data = Utils.EncodeString(value);
            }
        }
        public uint DataAsInt
        {
            get
            {
                return BitConverter.ToUInt32(Data, 0);
            }
            set
            {
                Data = BitConverter.GetBytes(value);
            }
        }

        internal void Encode(BinaryWriter writer)
        {
            writer.Write((uint)Type);
            writer.Write(Flags);
            writer.Write(Key.Length);
            writer.Write(Value.Length);
            writer.Write(Data.Length);
            writer.Write(Key);
            writer.Write(Value);
            writer.Write(Data);
        }

        internal static Win7Block Decode(BinaryReader reader)
        {
            uint type = reader.ReadUInt32();
            uint flags = reader.ReadUInt32();

            int keyLen = reader.ReadInt32();
            int valueLen = reader.ReadInt32();
            int dataLen = reader.ReadInt32();

            byte[] key = reader.ReadBytes(keyLen);
            byte[] value = reader.ReadBytes(valueLen);
            byte[] data = reader.ReadBytes(dataLen);
            return new Win7Block
            {
                Type = (BlockType)type,
                Flags = flags,
                Key = key,
                Value = value,
                Data = data,
            };
        }
    }

    public sealed class PhysicalStoreWin7 : IPhysicalStore
    {
        private byte[] PreHeaderBytes = new byte[] { };
        private readonly List<Win7Block> Blocks = new List<Win7Block>();
        private readonly FileStream TSPrimary;
        private readonly FileStream TSSecondary;
        private readonly bool Production;

        public byte[] Serialize()
        {
            BinaryWriter writer = new BinaryWriter(new MemoryStream());
            writer.Write(PreHeaderBytes);

            foreach (Win7Block block in Blocks)
            {
                block.Encode(writer);
                writer.Align(4);
            }

            return writer.GetBytes();
        }

        public void Deserialize(byte[] data)
        {
            int len = data.Length;

            BinaryReader reader = new BinaryReader(new MemoryStream(data));
            PreHeaderBytes = reader.ReadBytes(8);

            while (reader.BaseStream.Position < len - 0x14)
            {
                Blocks.Add(Win7Block.Decode(reader));
                reader.Align(4);
            }
        }

        public void AddBlock(PSBlock block)
        {
            Blocks.Add(new Win7Block
            {
                Type = block.Type,
                Flags = block.Flags,
                Key = block.Key,
                Value = block.Value,
                Data = block.Data
            });
        }

        public void AddBlocks(IEnumerable<PSBlock> blocks)
        {
            foreach (PSBlock block in blocks)
            {
                AddBlock(block);
            }
        }

        public PSBlock GetBlock(string key, string value)
        {
            foreach (Win7Block block in Blocks)
            {
                if (block.KeyAsStr == key && block.ValueAsStr == value)
                {
                    return new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = block.Key,
                        Value = block.Value,
                        Data = block.Data
                    };
                }
            }

            return null;
        }

        public PSBlock GetBlock(string key, uint value)
        {
            foreach (Win7Block block in Blocks)
            {
                if (block.KeyAsStr == key && block.ValueAsInt == value)
                {
                    return new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = block.Key,
                        Value = block.Value,
                        Data = block.Data
                    };
                }
            }

            return null;
        }

        public void SetBlock(string key, string value, byte[] data)
        {
            for (int i = 0; i < Blocks.Count; i++)
            {
                Win7Block block = Blocks[i];

                if (block.KeyAsStr == key && block.ValueAsStr == value)
                {
                    block.Data = data;
                    Blocks[i] = block;
                    break;
                }
            }
        }

        public void SetBlock(string key, uint value, byte[] data)
        {
            for (int i = 0; i < Blocks.Count; i++)
            {
                Win7Block block = Blocks[i];

                if (block.KeyAsStr == key && block.ValueAsInt == value)
                {
                    block.Data = data;
                    Blocks[i] = block;
                    break;
                }
            }
        }

        public void SetBlock(string key, string value, string data)
        {
            SetBlock(key, value, Utils.EncodeString(data));
        }

        public void SetBlock(string key, string value, uint data)
        {
            SetBlock(key, value, BitConverter.GetBytes(data));
        }

        public void SetBlock(string key, uint value, string data)
        {
            SetBlock(key, value, Utils.EncodeString(data));
        }

        public void SetBlock(string key, uint value, uint data)
        {
            SetBlock(key, value, BitConverter.GetBytes(data));
        }

        public void DeleteBlock(string key, string value)
        {
            foreach (Win7Block block in Blocks)
            {
                if (block.KeyAsStr == key && block.ValueAsStr == value)
                {
                    Blocks.Remove(block);
                    return;
                }
            }
        }

        public void DeleteBlock(string key, uint value)
        {
            foreach (Win7Block block in Blocks)
            {
                if (block.KeyAsStr == key && block.ValueAsInt == value)
                {
                    Blocks.Remove(block);
                    return;
                }
            }
        }

        public PhysicalStoreWin7(string primaryPath, bool production)
        {
            TSPrimary = File.Open(primaryPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            TSSecondary = File.Open(primaryPath.Replace("-0.", "-1."), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            Production = production;

            Deserialize(PhysStoreCrypto.DecryptPhysicalStore(TSPrimary.ReadAllBytes(), production));
            TSPrimary.Seek(0, SeekOrigin.Begin);
        }

        public void Dispose()
        {
            if (TSPrimary.CanWrite && TSSecondary.CanWrite)
            {
                byte[] data = PhysStoreCrypto.EncryptPhysicalStore(Serialize(), Production, PSVersion.Win7);

                TSPrimary.SetLength(data.LongLength);
                TSSecondary.SetLength(data.LongLength);

                TSPrimary.Seek(0, SeekOrigin.Begin);
                TSSecondary.Seek(0, SeekOrigin.Begin);

                TSPrimary.WriteAllBytes(data);
                TSSecondary.WriteAllBytes(data);

                TSPrimary.Close();
                TSSecondary.Close();
            }
        }

        public byte[] ReadRaw()
        {
            byte[] data = PhysStoreCrypto.DecryptPhysicalStore(TSPrimary.ReadAllBytes(), Production);
            TSPrimary.Seek(0, SeekOrigin.Begin);
            return data;
        }

        public void WriteRaw(byte[] data)
        {
            byte[] encrData = PhysStoreCrypto.EncryptPhysicalStore(data, Production, PSVersion.Win7);

            TSPrimary.SetLength(encrData.LongLength);
            TSSecondary.SetLength(encrData.LongLength);

            TSPrimary.Seek(0, SeekOrigin.Begin);
            TSSecondary.Seek(0, SeekOrigin.Begin);

            TSPrimary.WriteAllBytes(encrData);
            TSSecondary.WriteAllBytes(encrData);

            TSPrimary.Close();
            TSSecondary.Close();
        }

        public IEnumerable<PSBlock> FindBlocks(string valueSearch)
        {
            List<PSBlock> results = new List<PSBlock>();

            foreach (Win7Block block in Blocks)
            {
                if (block.ValueAsStr.Contains(valueSearch))
                {
                    results.Add(new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = block.Key,
                        Value = block.Value,
                        Data = block.Data
                    });
                }
            }

            return results;
        }

        public IEnumerable<PSBlock> FindBlocks(uint valueSearch)
        {
            List<PSBlock> results = new List<PSBlock>();

            foreach (Win7Block block in Blocks)
            {
                if (block.ValueAsInt == valueSearch)
                {
                    results.Add(new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = block.Key,
                        Value = block.Value,
                        Data = block.Data
                    });
                }
            }

            return results;
        }
    }
}


// PhysicalStore/VariableBag.cs
namespace LibTSforge.PhysicalStore
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    public enum CRCBlockType : uint
    {
        UINT = 1 << 0,
        STRING = 1 << 1,
        BINARY = 1 << 2
    }

    public class CRCBlock
    {
        public CRCBlockType DataType;
        public byte[] Key;
        public string KeyAsStr
        {
            get
            {
                return Utils.DecodeString(Key);
            }
            set
            {
                Key = Utils.EncodeString(value);
            }
        }
        public byte[] Value;
        public string ValueAsStr
        {
            get
            {
                return Utils.DecodeString(Value);
            }
            set
            {
                Value = Utils.EncodeString(value);
            }
        }
        public uint ValueAsInt
        {
            get
            {
                return BitConverter.ToUInt32(Value, 0);
            }
            set
            {
                Value = BitConverter.GetBytes(value);
            }
        }

        public void Encode(BinaryWriter writer)
        {
            uint crc = CRC();
            writer.Write(crc);
            writer.Write((uint)DataType);
            writer.Write(Key.Length);
            writer.Write(Value.Length);

            writer.Write(Key);
            writer.Align(8);

            writer.Write(Value);
            writer.Align(8);
        }

        public static CRCBlock Decode(BinaryReader reader)
        {
            uint crc = reader.ReadUInt32();
            uint type = reader.ReadUInt32();
            uint lenName = reader.ReadUInt32();
            uint lenVal = reader.ReadUInt32();

            byte[] key = reader.ReadBytes((int)lenName);
            reader.Align(8);

            byte[] value = reader.ReadBytes((int)lenVal);
            reader.Align(8);

            CRCBlock block = new CRCBlock
            {
                DataType = (CRCBlockType)type,
                Key = key,
                Value = value,
            };

            if (block.CRC() != crc)
            {
                throw new InvalidDataException("Invalid CRC in variable bag.");
            }

            return block;
        }

        public uint CRC()
        {
            BinaryWriter wtemp = new BinaryWriter(new MemoryStream());
            wtemp.Write(0);
            wtemp.Write((uint)DataType);
            wtemp.Write(Key.Length);
            wtemp.Write(Value.Length);
            wtemp.Write(Key);
            wtemp.Write(Value);
            return Utils.CRC32(wtemp.GetBytes());
        }
    }

    public class VariableBag
    {
        public List<CRCBlock> Blocks = new List<CRCBlock>();

        public void Deserialize(byte[] data)
        {
            int len = data.Length;

            BinaryReader reader = new BinaryReader(new MemoryStream(data));

            while (reader.BaseStream.Position < len - 0x10)
            {
                Blocks.Add(CRCBlock.Decode(reader));
            }
        }

        public byte[] Serialize()
        {
            BinaryWriter writer = new BinaryWriter(new MemoryStream());

            foreach (CRCBlock block in Blocks)
            {
                block.Encode(writer);
            }

            return writer.GetBytes();
        }

        public CRCBlock GetBlock(string key)
        {
            foreach (CRCBlock block in Blocks)
            {
                if (block.KeyAsStr == key)
                {
                    return block;
                }
            }

            return null;
        }

        public void SetBlock(string key, byte[] value)
        {
            for (int i = 0; i < Blocks.Count; i++)
            {
                CRCBlock block = Blocks[i];

                if (block.KeyAsStr == key)
                {
                    block.Value = value;
                    Blocks[i] = block;
                    break;
                }
            }
        }

        public void DeleteBlock(string key)
        {
            foreach (CRCBlock block in Blocks)
            {
                if (block.KeyAsStr == key)
                {
                    Blocks.Remove(block);
                    return;
                }
            }
        }

        public VariableBag(byte[] data)
        {
            Deserialize(data);
        }

        public VariableBag()
        {

        }
    }
}
'@
$ErrorActionPreference = 'Stop'
$cp = [CodeDom.Compiler.CompilerParameters] [string[]]@("System.dll", "System.Core.dll", "System.ServiceProcess.dll", "System.Xml.dll")
$cp.CompilerOptions = "/unsafe"
$lang = If ((Get-Host).Version.Major -gt 2) { "CSharp" } Else { "CSharpVersion3" }

$ctemp = "$env:SystemRoot\Temp\"
if (-Not (Test-Path -Path $ctemp)) { New-Item -Path $ctemp -ItemType Directory > $null }
$env:TMP = $ctemp
$env:TEMP = $ctemp

$cp.GenerateInMemory = $true
Add-Type -Language $lang -TypeDefinition $src -CompilerParameters $cp
if ($env:_debug -eq '0') {
    [LibTSforge.Logger]::HideOutput = $true
}
$ver = [LibTSforge.Utils]::DetectVersion()
$prod = [LibTSforge.Utils]::DetectCurrentKey()
$tsactids = @($args)

function Get-WmiInfo {
    param ([string]$tsactid, [string]$property)
    
    $query = "SELECT ID, $property FROM SoftwareLicensingProduct WHERE ID='$tsactid'"
    $record = Get-WmiObject -Query $query
    if ($record) {
        return $record.$property
    }
}

if ($env:resetstuff -eq $null) {
    foreach ($tsactid in $tsactids) {
        try {
            $prodDes = Get-WmiInfo -tsactid $tsactid -property "Description"
            $prodName = Get-WmiInfo -tsactid $tsactid -property "Name"
            if ($prodName) {
                $nameParts = $prodName -split ',', 2
                $prodName = if ($nameParts.Count -gt 1) { ($nameParts[1].Trim() -split '[ ,]')[0] } else { $null }
            }
            [LibTSforge.Modifiers.GenPKeyInstall]::InstallGenPKey($ver, $prod, $tsactid)
            [LibTSforge.Activators.ZeroCID]::Activate($ver, $prod, $tsactid)
            $licenseStatus = Get-WmiInfo -tsactid $tsactid -property "LicenseStatus"
            if ($licenseStatus -eq 1) {
                if ($prodDes -match 'KMS' -and $prodDes -notmatch 'CLIENT') {
                    [LibTSforge.Modifiers.KMSHostCharge]::Charge($ver, $tsactid, $prod)
                    Write-Host "[$prodName] CSVLK is permanently activated with ZeroCID." -ForegroundColor White -BackgroundColor DarkGreen
                    Write-Host "[$prodName] CSVLK is charged with 25 clients for 30 days." -ForegroundColor White -BackgroundColor DarkGreen
                }
                else {
                    Write-Host "[$prodName] is permanently activated with ZeroCID." -ForegroundColor White -BackgroundColor DarkGreen
                }
            }
            else {
                Write-Host "[$prodName] activation has failed." -ForegroundColor White -BackgroundColor DarkRed
                $errcode = 3
            }
        }
        catch {
            $errcode = 3
            Write-Host "$($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
            Write-Host "[$prodName] activation has failed." -ForegroundColor White -BackgroundColor DarkRed
        }
    }
}

if ($env:resetstuff -eq '1') {
    try {
        [LibTSforge.Modifiers.TamperedFlagsDelete]::DeleteTamperFlags($ver, $prod)
        [LibTSforge.SPP.SLApi]::RefreshLicenseStatus()
        [LibTSforge.Modifiers.RearmReset]::Reset($ver, $prod)
        [LibTSforge.Modifiers.GracePeriodReset]::Reset($ver, $prod)
        [LibTSforge.Modifiers.KeyChangeLockDelete]::Delete($ver, $prod)
    }
    catch {
        $errcode = 3
        Write-Host "$($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
    }
}

Exit $errcode
:tsforge:

::========================================================================================================================================

::  Get Windows Activation ID

:wintsid:
$SysPath = "$env:SystemRoot\System32"
if (Test-Path "$env:SystemRoot\Sysnative\reg.exe") {
    $SysPath = "$env:SystemRoot\Sysnative"
}

function Windows-ActID {
    param (
        [string]$edition,
        [string]$keytype
    )
    
    $filePatterns = @(
        "$SysPath\spp\tokens\skus\$edition\$edition*.xrm-ms",
        "$SysPath\spp\tokens\skus\Security-SPP-Component-SKU-$edition\*-$edition-*.xrm-ms"
    )
    
    switch ($keytype) {
        "zero" {
            $licenseTypes = @('OEM_DM', 'OEM_COA_SLP', 'OEM_COA_NSLP', 'MAK', 'RETAIL')
        }
        "ks" {
            $licenseTypes = @('KMSCLIENT')
        }
        "avma" {
            $licenseTypes = @('VIRTUAL_MACHINE')
        }
        "kmshost" {
            $licenseTypes = @('KMS_')
        }
    }
    
    $softwareLicensingProducts = Get-WmiObject -Query "SELECT ID, Description, LicenseFamily FROM SoftwareLicensingProduct WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f'" | Where-Object { $_.LicenseFamily -eq $edition }
    
    $orderedLicenses = @()
    foreach ($type in $licenseTypes) {
        $orderedLicenses += $softwareLicensingProducts | Where-Object { $_.Description -match $type } | Select-Object -ExpandProperty ID
    }
    
    $fileIds = @()
    $muiLockedIds = @()
    $kmsCountedIdCounts = @{}

    $t = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0)

    $methods = @(
        @{name = 'SLOpen'; returnType = [Int32]; parameters = @([IntPtr].MakeByRefType()) },
        @{name = 'SLClose'; returnType = [Int32]; parameters = @([IntPtr]) },
        @{name = 'SLGetProductSkuInformation'; returnType = [Int32]; parameters = @([IntPtr], [Guid].MakeByRefType(), [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()) },
        @{name = 'SLGetLicense'; returnType = [Int32]; parameters = @([IntPtr], [Guid].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()) }
    )
    
    foreach ($method in $methods) {
        $t.DefinePInvokeMethod($method.name, 'slc.dll', 22, 1, $method.returnType, $method.parameters, 1, 3).SetImplementationFlags(128)
    }
    
    $w = $t.CreateType()
    $m = [Runtime.InteropServices.Marshal]

    function GetLicenseInfo($SkuId, $checkType) {
        $result = $false
        $c = 0; $b = 0
        
        [void]$w::SLGetProductSkuInformation($hSLC, [ref][Guid]$SkuId, "fileId", [ref]$null, [ref]$c, [ref]$b)
        $FileId = $m::PtrToStringUni($b)
        
        $c = 0; $b = 0
        [void]$w::SLGetLicense($hSLC, [ref][Guid]$FileId, [ref]$c, [ref]$b)
        $blob = New-Object byte[] $c; $m::Copy($b, $blob, 0, $c)
        $cont = [Text.Encoding]::UTF8.GetString($blob)
        $xml = [xml]$cont.SubString($cont.IndexOf('<r'))
        
        if ($checkType -eq 'MUI') {
            $xml.licenseGroup.license[0].grant | foreach {
                $_.allConditions.allConditions.productPolicies.policyStr | where { $_.name -eq 'Kernel-MUI-Language-Allowed' } | foreach {
                    if ($_.InnerText -ne 'EMPTY') { $result = $true }
                }
            }
        }
        elseif ($checkType -eq 'KMS') {
            $xml.licenseGroup.license[0].grant | foreach {
                $_.allConditions.allConditions.productPolicies.policyStr | where { $_.name -eq 'Security-SPP-KmsCountedIdList' } | foreach {
                    $result = ($_.InnerText.Replace(' ', '').Replace("`n", '') -split ',').Count
                }
            }
        }
        
        return $result
    }

    $hSLC = 0; [void]$w::SLOpen([ref]$hSLC)

    foreach ($id in $orderedLicenses) {
        if ($keytype -eq 'kmshost') {
            $kmsCount = GetLicenseInfo $id 'KMS'
            if ($kmsCount -gt 0) {
                $kmsCountedIdCounts[$id] = $kmsCount
            }
        }
        if ($edition -notcontains "CountrySpecific" -and (GetLicenseInfo $id 'MUI')) {
            $muiLockedIds += $id
        }
    }
    
    foreach ($filePattern in $filePatterns) {
        $files = Get-ChildItem -Path $filePattern -Filter '*.xrm-ms' -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            if ($null -ne $file.FullName) {
                $content = Get-Content -Path $file.FullName -ErrorAction SilentlyContinue | Out-String
                foreach ($id in $orderedLicenses) {
                    if ($content -match "name=`"productSkuId`">\{$id\}" -and -not ($file.Name -match 'Beta|Test')) {
                        $fileIds += $id
                    }
                }
            }
        }
    }
    
    if ($kmsCountedIdCounts.Count -gt 0) {
        $idWithMostIds = $kmsCountedIdCounts.GetEnumerator() | Sort-Object Value -Descending
        $fileIds = $idWithMostIds | Select-Object -ExpandProperty Key
    }
    else {
        if ($fileIds.Count -eq 0) {
            $fileIds = $orderedLicenses
        }
    
        $fileIds = $orderedLicenses | Where-Object { $fileIds -contains $_ -and $muiLockedIds -notcontains $_ } | Select-Object -Unique
    }
    
    [void]$w::SLClose($hSLC)
    
    $pkeyconfig = "$SysPath\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms"
    if ($keytype -eq 'kmshost') {
        $csvlkPath = "$SysPath\spp\tokens\pkeyconfig\pkeyconfig-csvlk.xrm-ms"
        if (Test-Path $csvlkPath) {
            $pkeyconfig = $csvlkPath
        }
    }
    
    $data = [xml][Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(([xml](get-content $pkeyconfig)).licenseGroup.license.otherInfo.infoTables.infoList.infoBin.InnerText))
    
    $betaIds = @()
    $excludedIds = @()
    $checkedIds = @()
    
    foreach ($id in $fileIds) {
        $actConfig = $data.ProductKeyConfiguration.Configurations.Configuration | Where-Object { $_.ActConfigId -eq "{$id}" }
        if ($actConfig) {
            $productDescription = $actConfig.ProductDescription
            $productEditionID = $actConfig.EditionID
            if ($productDescription -match 'MUI locked|Tencent|Qihoo|WAU') {
                $excludedIds += $id
                continue
            }
    
            if ($productDescription -match 'Beta|RC |Next |Test|Pre-') {
                $betaIds += $id
                continue
            }
    
            if ($keytype -ne 'kmshost' -and $productEditionID -eq '$edition') {
                $checkedIds += $id
                continue
            }
    
            $refGroupId = $actConfig.RefGroupId
            $publicKey = $data.ProductKeyConfiguration.PublicKeys.PublicKey | Where-Object { $_.GroupId -eq $refGroupId -and $_.AlgorithmId -eq 'msft:rm/algorithm/pkey/2009' }
            if ($publicKey) {
                $keyRanges = $data.ProductKeyConfiguration.KeyRanges.KeyRange | Where-Object { $_.RefActConfigId -eq "{$id}" }
                foreach ($keyRange in $keyRanges) {
                    if ($keyRange.EulaType -match 'WAU') {
                        $excludedIds += $id
                        break
                    }
                }
            }
        }
    }
    
    $prefinalIds = @()
    $finalIds = @()
    
    $prefinalIds = $fileIds | Where-Object { $excludedIds -notcontains $_ } | Select-Object -Unique
    $finalIds = $prefinalIds | Where-Object { $betaIds -notcontains $_ } | Select-Object -Unique
    
    if ($finalIds.Count -eq 0) {
        $finalIds = $prefinalIds
    }
    
    if ($checkedIds.Count -gt 0) {
        $finalIds = $checkedIds + $finalIds
    }
    
    $firstId = $finalIds | Select-Object -First 1
    return $firstId.ToLower()
}

Windows-ActID -edition "$env:tsedition" -keytype "$env:keytype"
:wintsid:

::========================================================================================================================================

::  Get Office Activation ID

:offtsid:
function Office-ActID {
    param (
        [string]$pkeypath,
        [string]$edition,
        [string]$keytype
    )

    switch ($keytype) {
        "zero" { $productKeyTypes = @("OEM:NONSLP","Volume:MAK","Retail") }
        "ks" { $productKeyTypes = @("Volume:GVLK") }
    }

    $data = [xml][Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(([xml](Get-Content $pkeypath)).licenseGroup.license.otherInfo.infoTables.infoList.infoBin.InnerText))
    $configurations = $data.ProductKeyConfiguration.Configurations.Configuration

    $filteredConfigs = @()
    foreach ($type in $productKeyTypes) {
        $filteredConfigs += $configurations | Where-Object { 
            $_.EditionId -eq $edition -and 
            $_.ProductKeyType -eq $type -and 
            $_.ProductDescription -notmatch 'demo|MSDN|PIN'
        }
    }

    $filterPreview = $filteredConfigs | Where-Object { $_.ProductDescription -notmatch 'preview' }

    if ($filterPreview.Count -ne 0) {
        $filteredConfigs = $filterPreview
    } 

    $firstConfig = ($filteredConfigs | Select-Object -First 1).ActConfigID -replace '^\{|\}$', ''
    return $firstConfig.ToLower()
}

Office-ActID -pkeypath "$env:pkeypath" -edition "$env:_License" -keytype "$env:keytype"
:offtsid:

::========================================================================================================================================

::  1st column = Office version number
::  2nd column = Activation ID
::  3rd column = Edition
::  Separator  = "_"

:ts_msiofficedata

for %%# in (
:: Office 2013
15_ab4d047b-97cf-4126-a69f-34df08e2f254_AccessRetail
15_259de5be-492b-44b3-9d78-9645f848f7b0_AccessRuntimeRetail
15_4374022d-56b8-48c1-9bb7-d8f2fc726343_AccessVolume
15_1b1d9bd5-12ea-4063-964c-16e7e87d6e08_ExcelRetail
15_ac1ae7fd-b949-4e04-a330-849bc40638cf_ExcelVolume
15_cfaf5356-49e3-48a8-ab3c-e729ab791250_GrooveRetail
15_4825ac28-ce41-45a7-9e6e-1fed74057601_GrooveVolume
15_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_HomeBusinessPipcRetail
15_a2b90e7a-a797-4713-af90-f0becf52a1dd_HomeBusinessRetail
15_1fdfb4e4-f9c9-41c4-b055-c80daf00697d_HomeStudentARMRetail
15_ebef9f05-5273-404a-9253-c5e252f50555_HomeStudentPlusARMRetail
15_f2de350d-3028-410a-bfae-283e00b44d0e_HomeStudentRetail
15_44984381-406e-4a35-b1c3-e54f499556e2_InfoPathRetail
15_9e016989-4007-42a6-8051-64eb97110cf2_InfoPathVolume
15_9103f3ce-1084-447a-827e-d6097f68c895_LyncAcademicRetail
15_ff693bf4-0276-4ddb-bb42-74ef1a0c9f4d_LyncEntryRetail
15_fada6658-bfc6-4c4e-825a-59a89822cda8_LyncRetail
15_e1264e10-afaf-4439-a98b-256df8bb156f_LyncVolume
15_69ec9152-153b-471a-bf35-77ec88683eae_MondoRetail
15_f33485a0-310b-4b72-9a0e-b1d605510dbd_MondoVolume
15_3391e125-f6e4-4b1e-899c-a25e6092d40d_OneNoteFreeRetail
15_8b524bcc-67ea-4876-a509-45e46f6347e8_OneNoteRetail
15_b067e965-7521-455b-b9f7-c740204578a2_OneNoteVolume
15_12004b48-e6c8-4ffa-ad5a-ac8d4467765a_OutlookRetail
15_8d577c50-ae5e-47fd-a240-24986f73d503_OutlookVolume
15_5aab8561-1686-43f7-9ff5-2c861da58d17_PersonalPipcRetail
15_17e9df2d-ed91-4382-904b-4fed6a12caf0_PersonalRetail
15_31743b82-bfbc-44b6-aa12-85d42e644d5b_PowerPointRetail
15_e40dcb44-1d5c-4085-8e8f-943f33c4f004_PowerPointVolume
15_064383fa-1538-491c-859b-0ecab169a0ab_ProPlusRetail
15_2b88c4f2-ea8f-43cd-805e-4d41346e18a7_ProPlusVolume
15_4e26cac1-e15a-4467-9069-cb47b67fe191_ProfessionalPipcRetail
15_44bc70e2-fb83-4b09-9082-e5557e0c2ede_ProfessionalRetail
15_2f72340c-b555-418d-8b46-355944fe66b8_ProjectProRetail
15_ed34dc89-1c27-4ecd-8b2f-63d0f4cedc32_ProjectProVolume
15_58d95b09-6af6-453d-a976-8ef0ae0316b1_ProjectStdRetail
15_2b9e4a37-6230-4b42-bee2-e25ce86c8c7a_ProjectStdVolume
15_c3a0814a-70a4-471f-af37-2313a6331111_PublisherRetail
15_38ea49f6-ad1d-43f1-9888-99a35d7c9409_PublisherVolume
15_ba3e3833-6a7e-445a-89d0-7802a9a68588_SPDRetail
15_32255c0a-16b4-4ce2-b388-8a4267e219eb_StandardRetail
15_a24cca51-3d54-4c41-8a76-4031f5338cb2_StandardVolume
15_a56a3b37-3a35-4bbb-a036-eee5f1898eee_VisioProRetail
15_3e4294dd-a765-49bc-8dbd-cf8b62a4bd3d_VisioProVolume
15_980f9e3e-f5a8-41c8-8596-61404addf677_VisioStdRetail
15_44a1f6ff-0876-4edb-9169-dbb43101ee89_VisioStdVolume
15_191509f2-6977-456f-ab30-cf0492b1e93a_WordRetail
15_9cedef15-be37-4ff0-a08a-13a045540641_WordVolume
:: Office 365 - 15.0 version
15_742178ed-6b28-42dd-b3d7-b7c0ea78741b_O365BusinessRetail
15_a96f8dae-da54-4fad-bdc6-108da592707a_O365HomePremRetail
15_e3dacc06-3bc2-4e13-8e59-8e05f3232325_O365ProPlusRetail
15_0bc1dae4-6158-4a1c-a893-807665b934b2_O365SmallBusPremRetail
:: Office 365 - 16.0 version
16_742178ed-6b28-42dd-b3d7-b7c0ea78741b_O365BusinessRetail
16_2f5c71b4-5b7a-4005-bb68-f9fac26f2ea3_O365EduCloudRetail
16_a96f8dae-da54-4fad-bdc6-108da592707a_O365HomePremRetail
16_e3dacc06-3bc2-4e13-8e59-8e05f3232325_O365ProPlusRetail
16_0bc1dae4-6158-4a1c-a893-807665b934b2_O365SmallBusPremRetail
:: Office 2016
16_bfa358b0-98f1-4125-842e-585fa13032e6_AccessRetail
16_9d9faf9e-d345-4b49-afce-68cb0a539c7c_AccessRuntimeRetail
16_3b2fa33f-cd5a-43a5-bd95-f49f3f546b0b_AccessVolume
16_424d52ff-7ad2-4bc7-8ac6-748d767b455d_ExcelRetail
16_685062a7-6024-42e7-8c5f-6bb9e63e697f_ExcelVolume
16_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_HomeBusinessPipcRetail
16_86834d00-7896-4a38-8fae-32f20b86fa2b_HomeBusinessRetail
16_090896a0-ea98-48ac-b545-ba5da0eb0c9c_HomeStudentARMRetail
16_6bbe2077-01a4-4269-bf15-5bf4d8efc0b2_HomeStudentPlusARMRetail
16_c28acdb8-d8b3-4199-baa4-024d09e97c99_HomeStudentRetail
16_e2127526-b60c-43e0-bed1-3c9dc3d5a468_HomeStudentVNextRetail
16_69ec9152-153b-471a-bf35-77ec88683eae_MondoRetail
16_2cd0ea7e-749f-4288-a05e-567c573b2a6c_MondoVolume
16_436366de-5579-4f24-96db-3893e4400030_OneNoteFreeRetail
16_83ac4dd9-1b93-40ed-aa55-ede25bb6af38_OneNoteRetail
16_23b672da-a456-4860-a8f3-e062a501d7e8_OneNoteVolume
16_5a670809-0983-4c2d-8aad-d3c2c5b7d5d1_OutlookRetail
16_50059979-ac6f-4458-9e79-710bcb41721a_OutlookVolume
16_5aab8561-1686-43f7-9ff5-2c861da58d17_PersonalPipcRetail
16_a9f645a1-0d6a-4978-926a-abcb363b72a6_PersonalRetail
16_f32d1284-0792-49da-9ac6-deb2bc9c80b6_PowerPointRetail
16_9b4060c9-a7f5-4a66-b732-faf248b7240f_PowerPointVolume
16_de52bd50-9564-4adc-8fcb-a345c17f84f9_ProPlusRetail
16_c47456e3-265d-47b6-8ca0-c30abbd0ca36_ProPlusVolume
16_4e26cac1-e15a-4467-9069-cb47b67fe191_ProfessionalPipcRetail
16_d64edc00-7453-4301-8428-197343fafb16_ProfessionalRetail
16_2f72340c-b555-418d-8b46-355944fe66b8_ProjectProRetail
16_82f502b5-b0b0-4349-bd2c-c560df85b248_ProjectProVolume
16_16728639-a9ab-4994-b6d8-f81051e69833_ProjectProXVolume
16_58d95b09-6af6-453d-a976-8ef0ae0316b1_ProjectStdRetail
16_82e6b314-2a62-4e51-9220-61358dd230e6_ProjectStdVolume
16_431058f0-c059-44c5-b9e7-ed2dd46b6789_ProjectStdXVolume
16_6e0c1d99-c72e-4968-bcb7-ab79e03e201e_PublisherRetail
16_fcc1757b-5d5f-486a-87cf-c4d6dedb6032_PublisherVolume
16_9103f3ce-1084-447a-827e-d6097f68c895_SkypeServiceBypassRetail
16_971cd368-f2e1-49c1-aedd-330909ce18b6_SkypeforBusinessEntryRetail
16_418d2b9f-b491-4d7f-84f1-49e27cc66597_SkypeforBusinessRetail
16_03ca3b9a-0869-4749-8988-3cbc9d9f51bb_SkypeforBusinessVolume
16_4a31c291-3a12-4c64-b8ab-cd79212be45e_StandardRetail
16_0ed94aac-2234-4309-ba29-74bdbb887083_StandardVolume
16_a56a3b37-3a35-4bbb-a036-eee5f1898eee_VisioProRetail
16_295b2c03-4b1c-4221-b292-1411f468bd02_VisioProVolume
16_0594dc12-8444-4912-936a-747ca742dbdb_VisioProXVolume
16_980f9e3e-f5a8-41c8-8596-61404addf677_VisioStdRetail
16_44151c2d-c398-471f-946f-7660542e3369_VisioStdVolume
16_1d1c6879-39a3-47a5-9a6d-aceefa6a289d_VisioStdXVolume
16_cacaa1bf-da53-4c3b-9700-11738ef1c2a5_WordRetail
16_c3000759-551f-4f4a-bcac-a4b42cbf1de2_WordVolume
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do (

if "%oVer%"=="%%A" (
for /f "tokens=*" %%x in ('findstr /i /c:"%%B" "%_oBranding%"') do set "prodId=%%x"
set prodId=!prodId:"/>=!
set prodId=!prodId:~-4!
if "%oVer%"=="14" (
REM Exception case for Visio because wrong primary product ID is mentioned in Branding.xml
echo %%C | find /i "Visio" %nul% && set prodId=0057
)
reg query "%1\Registration\{%%B}" /v ProductCode %nul2% | find /i "-!prodId!-" %nul% && (
reg query "%1\Common\InstalledPackages" %nul2% | find /i "-!prodId!-" %nul% && (
if defined _oIds (set _oIds=!_oIds! %%C) else (set _oIds=%%C)
)
)
)

)
)
exit /b

::========================================================================================================================================

:ts_getedition

set tsedition=
set _wtarget=

if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL) get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND LicenseDependsOn is NULL').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
%chkedi% do if not errorlevel 1 call set "_wtarget= !_wtarget! %%a "

::  SKU and Edition ID database

for %%# in (
1:Ultimate
2:HomeBasic
3:HomePremium
4:Enterprise
5:HomeBasicN
6:Business
7:ServerStandard
8:ServerDatacenter
9:ServerSBSStandard
10:ServerEnterprise
11:Starter
12:ServerDatacenterCore
13:ServerStandardCore
14:ServerEnterpriseCore
15:ServerEnterpriseIA64
16:BusinessN
17:ServerWeb
18:ServerHPC
19:ServerHomeStandard
20:ServerStorageExpress
21:ServerStorageStandard
22:ServerStorageWorkgroup
23:ServerStorageEnterprise
24:ServerWinSB
25:ServerSBSPremium
26:HomePremiumN
27:EnterpriseN
28:UltimateN
29:ServerWebCore
30:ServerMediumBusinessManagement
31:ServerMediumBusinessSecurity
32:ServerMediumBusinessMessaging
33:ServerWinFoundation
34:ServerHomePremium
35:ServerWinSBV
36:ServerStandardV
37:ServerDatacenterV
38:ServerEnterpriseV
39:ServerDatacenterVCore
40:ServerStandardVCore
41:ServerEnterpriseVCore
42:ServerHyperCore
43:ServerStorageExpressCore
44:ServerStorageStandardCore
45:ServerStorageWorkgroupCore
46:ServerStorageEnterpriseCore
47:StarterN
48:Professional
49:ProfessionalN
50:ServerSolution
51:ServerForSBSolutions
52:ServerSolutionsPremium
53:ServerSolutionsPremiumCore
54:ServerSolutionEM
55:ServerForSBSolutionsEM
56:ServerEmbeddedSolution
57:ServerEmbeddedSolutionCore
58:ProfessionalEmbedded
59:ServerEssentialManagement
60:ServerEssentialAdditional
61:ServerEssentialManagementSvc
62:ServerEssentialAdditionalSvc
63:ServerSBSPremiumCore
64:ServerHPCV
65:Embedded
66:StarterE
67:HomeBasicE
68:HomePremiumE
69:ProfessionalE
70:EnterpriseE
71:UltimateE
72:EnterpriseEval
74:Prerelease
76:ServerMultiPointStandard
77:ServerMultiPointPremium
79:ServerStandardEval
80:ServerDatacenterEval
81:PrereleaseARM
82:PrereleaseN
84:EnterpriseNEval
85:EmbeddedAutomotive
86:EmbeddedIndustryA
87:ThinPC
88:EmbeddedA
89:EmbeddedIndustry
90:EmbeddedE
91:EmbeddedIndustryE
92:EmbeddedIndustryAE
93:ProfessionalPlus
95:ServerStorageWorkgroupEval
96:ServerStorageStandardEval
97:CoreARM
98:CoreN
99:CoreCountrySpecific
100:CoreSingleLanguage
101:Core
103:ProfessionalWMC
104:MobileCore
105:EmbeddedIndustryEval
106:EmbeddedIndustryEEval
107:EmbeddedEval
108:EmbeddedEEval
109:CoreSystemServer
110:ServerCloudStorage
111:CoreConnected
112:ProfessionalStudent
113:CoreConnectedN
114:ProfessionalStudentN
115:CoreConnectedSingleLanguage
116:CoreConnectedCountrySpecific
117:ConnectedCar
118:IndustryHandheld
119:PPIPRO
120:ServerARM64
121:Education
122:EducationN
123:IoTUAP
124:ServerHI
125:EnterpriseS
126:EnterpriseSN
127:ProfessionalS
128:ProfessionalSN
129:EnterpriseSEval
130:EnterpriseSNEval
131:IoTUAPCommercial
133:MobileEnterprise
134:AnalogOneCoreEnterprise
135:AnalogOneCore
136:Holographic
138:ProfessionalSingleLanguage
139:ProfessionalCountrySpecific
140:EnterpriseSubscription
141:EnterpriseSubscriptionN
143:ServerDatacenterNano
144:ServerStandardNano
145:ServerDatacenterACor
146:ServerStandardACor
147:ServerDatacenterCor
148:ServerStandardCor
149:UtilityVM
159:ServerDatacenterEvalCor
160:ServerStandardEvalCor
161:ProfessionalWorkstation
162:ProfessionalWorkstationN
163:ServerAzure
164:ProfessionalEducation
165:ProfessionalEducationN
168:ServerAzureCor
169:ServerAzureNano
171:EnterpriseG
172:EnterpriseGN
173:BusinessSubscription
174:BusinessSubscriptionN
175:ServerRdsh
178:Cloud
179:CloudN
180:HubOS
182:OneCoreUpdateOS
183:CloudE
184:Andromeda
185:IoTOS
186:CloudEN
187:IoTEdgeOS
188:IoTEnterprise
189:ModernPC
191:IoTEnterpriseS
192:SystemOS
193:NativeOS
194:GameCoreXbox
195:GameOS
196:DurangoHostOS
197:ScarlettHostOS
198:Keystone
199:CloudHost
200:CloudMOS
201:CloudCore
202:CloudEditionN
203:CloudEdition
204:WinVOS
205:IoTEnterpriseSK
206:IoTEnterpriseK
207:IoTEnterpriseSEval
208:AgentBridge
209:NanoHost
210:WNC
406:ServerAzureStackHCICor
407:ServerTurbine
408:ServerTurbineCor

REM Some old edition names with same SKU ID

4:ProEnterprise
6:ProStandard
10:ProSBS
16:ProStandardN
18:ServerComputeCluster
19:ServerHome
30:ServerMidmarketStandard
31:ServerMidmarketEdge
32:ServerMidmarketPremium
33:ServerSBSPrime
42:ServerHyper
64:ServerComputeClusterV
85:EmbeddedIapetus
86:EmbeddedTethys
88:EmbeddedDione
89:EmbeddedRhea
90:EmbeddedEnceladus
109:ServerNano
124:ServerCloudHostInfrastructure
133:MobileBusiness
134:HololensEnterprise
145:ServerDatacenterSCor
146:ServerStandardSCor
147:ServerDatacenterWSCor
148:ServerStandardWSCor
189:Lite
) do (
for /f "tokens=1-2 delims=:" %%A in ("%%#") do if "%osSKU%"=="%%A" if not defined tsedition (
echo "%_wtarget%" | find /i " %%B " %nul% && set tsedition=%%B
)
)

if defined tsedition exit /b

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('GetEditionNameFromId', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([int], [IntPtr].MakeByRefType()), 1, 3);
set d1=%d1% $out = 0; [void]$TypeBuilder.CreateType()::GetEditionNameFromId(%osSKU%, [ref]$out);$s=[Runtime.InteropServices.Marshal]::PtrToStringUni($out); $s

for %%# in (pkeyhelper.dll) do @if not "%%~$PATH:#"=="" (
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (
echo "%_wtarget%" | find /i " %%a " %nul% && set tsedition=%%a
)
)

exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:KMS38Activation

::  To activate, run the script with "/KMS38" parameter or change 0 to 1 in below line
set _act=0

::  To remove KMS38 protection, run the script with /KMS38-RemoveProtection parameter or change 0 to 1 in below line
set _rem=0

::  To disable changing edition if current edition doesn't support KMS38 activation, change the value to 1 from 0 or run the script with "/KMS38-NoEditionChange" parameter
set _NoEditionChange=0

::  If value is changed in above lines or parameter is used then script will run in unattended mode

::========================================================================================================================================

cls
color 07
title  KMS38 Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/KMS38"                  set _act=1
if /i "%%A"=="/KMS38-RemoveProtection" set _rem=1
if /i "%%A"=="/KMS38-NoEditionChange"  set _NoEditionChange=1
if /i "%%A"=="-el"                     set _elev=1
)
)

for %%A in (%_act% %_rem% %_NoEditionChange%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

set _k38=
call :dk_setvar
set "specific_kms=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f"

if %winbuild% LSS 14393 (
%eline%
echo Unsupported OS version detected [%winbuild%].
echo KMS38 activation is only supported on Windows 10/11/Server, build 14393 and later.
echo:
if %winbuild% LSS 10240 (
call :dk_color %Blue% "Use TSforge activation option from the main menu."
) else (
call :dk_color %Blue% "Use HWID activation option from the main menu."
)
goto dk_done
)

::========================================================================================================================================

if %_rem%==1 goto :k_uninstall

:k_menu

if %_unattended%==0 (
cls
if not defined terminal mode 76, 25
title  KMS38 Activation %masver%

echo:
echo:
echo:
echo:
echo:           ______________________________________________________
echo:
echo                 [1] KMS38 Activation
echo                 ____________________________________________
echo:
echo                 [2] Remove KM38 Protection
echo:
echo                 [0] %_exitmsg%
echo:           ______________________________________________________
echo: 
call :dk_color2 %_White% "             " %_Green% "Choose a menu option using your keyboard [1,2,0]"
choice /C:120 /N
set _el=!errorlevel!
if !_el!==3  exit /b
if !_el!==2  goto :k_uninstall
if !_el!==1  goto :k_menu2
goto :k_menu
)

::========================================================================================================================================

:k_menu2

cls
if not defined terminal (
mode 110, 34
if exist "%SysPath%\spp\store_test\" mode 134, 34
)
title  KMS38 Activation %masver%

echo:
echo Initializing...
call :dk_chkmal

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" if not exist "%SysPath%\clipup.exe" set a_cor=1
if not exist %SysPath%\sppsvc.exe (set _fmiss=sppsvc.exe)
if not exist %SysPath%\ClipUp.exe if not defined a_cor (set _fmiss=%_fmiss%ClipUp.exe)

if defined _fmiss (
%eline%
echo [%_fmiss%] file is missing, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_ckeckwmic
call :dk_checksku
call :dk_product
call :dk_sppissue

::========================================================================================================================================

::  Check if system is permanently activated or not

call :dk_checkperm
if defined _perm (
cls
echo ___________________________________________________________________________________________
echo:
call :dk_color2 %_White% "     " %Green% "%winos% is already permanently activated."
call :dk_color2 %_White% "     " %Gray% "Activation is not required."
echo ___________________________________________________________________________________________
if %_unattended%==1 goto dk_done
echo:
choice /C:10 /N /M ">    [1] Activate Anyway [0] %_exitmsg% : "
if errorlevel 2 exit /b
)
cls

::========================================================================================================================================

::  Check Evaluation version

set _eval=
set _evalserv=

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _eval=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" set _evalserv=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" set _eval=1 & set _evalserv=1

if defined _eval (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% && (
%eline%
echo [%winos% ^| %winbuild%]
if defined _evalserv (
echo Server Evaluation cannot be activated. Convert it to full Server OS.
echo:
call :dk_color %Blue% "Go Back to main menu and use [Change Edition] option."
) else (
echo Evaluation editions cannot be activated outside of their evaluation period.
call :dk_color %Blue% "Use TSforge activation option from the main menu to reset evaluation period."
echo:
set fixes=%fixes% %mas%evaluation_editions
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%evaluation_editions"
)
goto dk_done
)
)

::========================================================================================================================================

:: Check clipup.exe for the detection and activation of server cor and acor editions

if defined a_cor (
if not exist "!_work!\clipup.exe" (
%eline%
echo clipup.exe doesn't exist in Server Cor/Acor [No GUI] versions.
echo The file is required for KMS38 activation.
echo Check the below page for instructions on how to activate it.
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%kms38"
goto dk_done
)
)

:: Check file signature

if defined a_cor (
%psc% "if (-not (Get-AuthenticodeSignature -FilePath '!_work!\clipup.exe').IsOSBinary) {Exit 3}" %nul%
if !errorlevel!==3 (
%eline%
echo Valid digital signature not found in clipup.exe file.
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)
)

::========================================================================================================================================

set error=

cls
echo:
call :dk_showosinfo

::========================================================================================================================================

echo Initiating Diagnostic Tests...

set "_serv=ClipSVC sppsvc KeyIso Winmgmt"

::  Client License Service (ClipSVC)
::  Software Protection
::  CNG Key Isolation
::  Windows Management Instrumentation

call :dk_errorcheck

::========================================================================================================================================

::  Check if GVLK (KMS key) is already installed or not

call :k_channel

::  Detect Key

set key=
set pkey=
set altkey=
set changekey=
set altedition=

call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f
if defined allapps call :kms38data key
if not defined key call :k_gvlk %nul%
if defined allapps if not defined key call :kms38fallback

if defined altkey (set key=%altkey%&set changekey=1)

set /a UBR=0
if %osSKU%==191 if defined altkey if defined altedition (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR %nul6%') do if not errorlevel 1 set /a UBR=%%b
if %winbuild% LSS 22598 if !UBR! LSS 2788 (
call :dk_color %Blue% "Windows must be updated to build 19044.2788 or higher for IotEnterpriseS KMS38 activation."
)
)

if not defined key if defined notfoundaltactID (
call :dk_color %Red% "Checking Alternate Edition for KMS38    [%altedition% Activation ID Not Found]"
)

if not defined key if not defined _gvlk (
echo:
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]

if exist "%SysPath%\spp\tokens\skus\%osedition%\*GVLK*.xrm-ms" set sppks=1

if defined skunotfound (
call :dk_color %Red% "Required license files not found in %SysPath%\spp\tokens\skus\"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

if defined sppks (
call :dk_color %Red% "KMS38 activation is supported but failed to find the key."
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

if not defined skunotfound if not defined sppks (
call :dk_color %Red% "This product does not support KMS38 activation."
call :dk_color %Blue% "Use TSforge activation option from the main menu."
set fixes=%fixes% %mas%
echo %mas%
)
echo:
goto dk_done
)

::========================================================================================================================================

::  Install key

echo:
if defined changekey (
call :dk_color %Blue% "[%altedition%] edition product key will be used to enable KMS38 activation."
echo:
)

if defined winsub (
call :dk_color %Blue% "Windows Subscription edition [SKU ID-%slcSKU%] found. Script will activate the base edition [SKU ID-%regSKU%]."
echo:
)

set _partial=
if not defined key (
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path %spp% where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null AND LicenseDependsOn is NULL" Get PartialProductKey /value %nul6%') do set "_partial=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT PartialProductKey FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').Get()).PartialProductKey | %% {echo ('PartialProductKey='+$_)}" %nul6%') do set "_partial=%%#"
call echo Checking Installed Product Key          [Partial Key - %%_partial%%] [Volume:GVLK]
)

if defined key (
call :dk_inskey "[%key%]"
)

::========================================================================================================================================

::  Check activation ID for setting specific KMS host

set app=
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL AND LicenseDependsOn is NULL) get ID /VALUE" %nul6%') do call set "app=%%a"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').Get()).ID | %% {echo ('ID='+$_)}" %nul6%') do call set "app=%%a"

if not defined app (
call :dk_color %Red% "Checking Installed GVLK Activation ID   [Not Found] Aborting..."
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto :dk_done
)

::========================================================================================================================================

::  Set specific KMS host to Local Host
::  By doing this, global KMS IP can not replace KMS38 activation but can be used with Office and other Windows Editions

echo:
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f

%nul% reg query "HKLM\%specific_kms%" && (
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regdel\:.*';iex ($f[1])"
%nul% reg delete "HKLM\%specific_kms%" /f
)

set k_error=
%nul% reg add "HKLM\%specific_kms%\%app%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2" || set k_error=1
%nul% reg add "HKLM\%specific_kms%\%app%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" || set k_error=1

if not defined k_error (
echo Adding Specific KMS Host                [LocalHost 127.0.0.2] [Successful]
) else (
call :dk_color %Red% "Adding Specific KMS Host                [LocalHost 127.0.0.2] [Failed]"
)

::========================================================================================================================================

::  Copy clipup.exe to System32 directory to activate Server Cor/Acor editions

if defined a_cor (
set "_clipup=%systemroot%\System32\clipup.exe"
pushd "!_work!\"
copy /y /b "ClipUp.exe" "!_clipup!" %nul%
popd

echo:
if exist "!_clipup!" (
echo Copying clipup.exe File to              [%systemroot%\System32\] [Successful]
) else (
call :dk_color %Red% "Copying clipup.exe File to              [%systemroot%\System32\] [Failed] Aborting..."
goto :k_final
)
)

::========================================================================================================================================

::  Generate GenuineTicket.xml and apply
::  In some cases clipup -v -o method fails and in some cases service restart method fails as well
::  To maximize success rate and get better error details, script will install tickets two times (service restart + clipup -v -o)

set "tdir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if not exist "%tdir%\" md "%tdir%\" %nul%

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" del /f /q "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" %nul%

::  Signature value is as it is, it's not encoded
::  Session ID is in Base64 encoded format. It's decoded value is "OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;GVLKExp=2038-01-19T03:14:07Z;DownlevelGenuineState=1;"
::  Check mass grave [.] dev/kms38#manual-activation to see how it's generated

set "signature=C52iGEoH+1VqzI6kEAqOhUyrWuEObnivzaVjyef8WqItVYd/xGDTZZ3bkxAI9hTpobPFNJyJx6a3uriXq3HVd7mlXfSUK9ydeoUdG4eqMeLwkxeb6jQWJzLOz41rFVSMtBL0e+ycCATebTaXS4uvFYaDHDdPw2lKY8ADj3MLgsA="
set "sessionId=TwBTAE0AYQBqAG8AcgBWAGUAcgBzAGkAbwBuAD0ANQA7AE8AUwBNAGkAbgBvAHIAVgBlAHIAcwBpAG8AbgA9ADEAOwBPAFMAUABsAGEAdABmAG8AcgBtAEkAZAA9ADIAOwBQAFAAPQAwADsARwBWAEwASwBFAHgAcAA9ADIAMAAzADgALQAwADEALQAxADkAVAAwADMAOgAxADQAOgAwADcAWgA7AEQAbwB3AG4AbABlAHYAZQBsAEcAZQBuAHUAaQBuAGUAUwB0AGEAdABlAD0AMQA7AAAA"
<nul set /p "=<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=%sessionId%;TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">%signature%</signature></signatures></genuineProperties></genuineAuthorization>" >"%tdir%\GenuineTicket"

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "Generating GenuineTicket.xml            [Failed, aborting...]"
if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
goto :k_final
) else (
echo Generating GenuineTicket.xml            [Successful]
)

set "_xmlexist=if exist "%tdir%\GenuineTicket.xml""

::  Stop sppsvc

%psc% "Start-Job { Stop-Service sppsvc -force } | Wait-Job -Timeout 20 | Out-Null"

sc query sppsvc | find /i "STOPPED" %nul% && (
echo Stopping sppsvc Service                 [Successful]
) || (
call :dk_color %Gray% "Stopping sppsvc Service                 [Failed]"
)

%_xmlexist% (
%psc% "Start-Job { Restart-Service ClipSVC } | Wait-Job -Timeout 20 | Out-Null"
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

%_xmlexist% (
set error=1
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
call :dk_color %Gray% "Installing GenuineTicket.xml            [Failed with ClipSVC service restart, wait...]"
)
)

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%
clipup -v -o

set rebuildinfo=

if not exist %ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat (
set error=1
set rebuildinfo=1
call :dk_color %Red% "Checking ClipSVC tokens.dat             [Not Found]"
)

%_xmlexist% (
set error=1
set rebuildinfo=1
call :dk_color %Red% "Installing GenuineTicket.xml            [Failed With clipup -v -o]"
)

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*.xml" (
set error=1
set rebuildinfo=1
call :dk_color %Red% "Checking Ticket Migration               [Failed]"
)

if not defined showfix if defined rebuildinfo (
set showfix=1
call :dk_color %Blue% "%_fixmsg%"
)

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%

::==========================================================================================================================================

call :dk_product

echo:
echo Activating...
echo:

call :k_checkexp
if defined _k38 (
call :k_actinfo
goto :k_final
)

::  Clear 180 Days KMS Activation lock with Windows SKU specific rearm and without the need to restart the system

if %_wmic% EQU 1 wmic path %spp% where ID='%app%' call ReArmsku %nul%
if %_wmic% EQU 0 %psc% "$null=([WMI]'%spp%=''%app%''').ReArmsku()" %nul%

if %errorlevel%==0 (
echo Applying SKU-ID Rearm                   [Successful]
) else (
call :dk_color %Red% "Applying SKU-ID Rearm                   [Failed]"
)
call :dk_refresh

echo:
call :k_checkexp
if defined _k38 (
call :k_actinfo
goto :k_final
)

call :dk_color %Red% "Activation Failed"
if not defined error call :dk_color %Blue% "%_fixmsg%"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"

::========================================================================================================================================

:k_final

::  Remove the specific KMS host (LocalHost) added by the script if activation is not completed

echo:
if not defined _k38 (
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f
%nul% reg query "HKLM\%specific_kms%" && (
call :dk_color %Red% "Removing the Added Specific KMS Host    [Failed]"
) || (
echo Removing the Added Specific KMS Host    [Successful]
)
)

::  Protect KMS38 if opted by the user and conditions are correct

if defined _k38 (
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regdel\:.*';& ([ScriptBlock]::Create($f[1])) -protect"
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg query "HKLM\%specific_kms%" && (
echo Protect KMS38 From KMS                  [Successful] [Locked a Registry Key]
) || (
call :dk_color %Red% "Protect KMS38 From KMS                  [Failed to Lock a Registry Key]"
)
)

::  clipup.exe does not exist in server cor and acor editions by default, it was copied there with this script

if defined a_cor if exist "%_clipup%" del /f /q "%_clipup%" %nul%

if defined a_cor (
if exist "%_clipup%" (
call :dk_color %Red% "Deleting Copied clipup.exe File         [Failed]"
) else (
echo Deleting Copied clipup.exe File         [Successful]
)
)

for %%# in (407) do if %osSKU%==%%# (
call :dk_color %Red% "%winos% does not support activation on non-azure platforms."
)

::  Trigger reevaluation of SPP's Scheduled Tasks

if defined _k38 (
call :dk_reeval %nul%
)
goto :dk_done

::========================================================================================================================================

:k_uninstall

cls
if not defined terminal mode 99, 28
title  Remove KMS38 Protection %masver%

%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f

%nul% reg query "HKLM\%specific_kms%" && (
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regdel\:.*';iex ($f[1])"
%nul% reg delete "HKLM\%specific_kms%" /f
)

echo:
%nul% reg query "HKLM\%specific_kms%" && (
call :dk_color %Red% "Removing Specific KMS Host              [Failed]"
) || (
echo Removing Specific KMS Host              [Successful]
)

goto :dk_done

::========================================================================================================================================

::  This code runs to protect/undo below registry key for KMS38 protection
::  HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f

::  KMS38 protection stops 180 days KMS Activation from replacing KMS38 activation

:regdel:
param (
    [switch]$protect
)

$SID = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
$Admin = ($SID.Translate([System.Security.Principal.NTAccount])).Value

if($protect) {
$ruleArgs = @("$Admin", "Delete, SetValue", "ContainerInherit", "None", "Deny")
} else {
$ruleArgs = @("$Admin", "FullControl", "Allow")
}

$path = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f'
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey($path, 'ReadWriteSubTree', 'ChangePermissions')
$acl = $key.GetAccessControl()

$rule = [System.Security.AccessControl.RegistryAccessRule]::new.Invoke($ruleArgs)
$acl.ResetAccessRule($rule)
$key.SetAccessControl($acl)
:regdel:

::========================================================================================================================================

::  Check KMS activation status

:k_actinfo

set xpr=
for /f "tokens=* delims=" %%# in ('%psc% "$([DateTime]::Now.addMinutes(%gpr%)).ToString('yyyy-MM-dd HH:mm:ss')" %nul6%') do set "xpr=%%#"
call :dk_color %Green% "%winos% is activated till !xpr!"
exit /b

::  Check remaining KMS activation grace period

:k_checkexp

set gpr=0
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL AND LicenseDependsOn is NULL) get GracePeriodRemaining /VALUE" %nul6%') do set "gpr=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT GracePeriodRemaining FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').Get()).GracePeriodRemaining | %% {echo ('GracePeriodRemaining='+$_)}" %nul6%') do set "gpr=%%#"
if %gpr% GTR 259200 (set _k38=1) else (set _k38=)
exit /b

::  Get Windows installed key channel

:k_channel

set _gvlk=
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path %spp% where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL and Description like '%%KMSCLIENT%%'" Get Name /value %nul6%') do (echo %%# findstr /i "Windows" %nul1% && set _gvlk=1)
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT Name FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL and Description like ''%%KMSCLIENT%%''').Get()).Name | %% {echo ('Name='+$_)}" %nul6%') do (echo %%# findstr /i "Windows" %nul1% && set _gvlk=1)
exit /b

::========================================================================================================================================

::  Get Product Key from pkeyhelper.dll for future new editions
::  It works on Windows 10 1803 (17134) and later builds.

:k_pkey

call :dk_reflection

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SkuGetProductKeyForEdition', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([int], [String], [String].MakeByRefType(), [String].MakeByRefType()), 1, 3);
set d1=%d1% $out = ''; [void]$TypeBuilder.CreateType()::SkuGetProductKeyForEdition(%1, %2, [ref]$out, [ref]$null); $out

set pkey=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkey=%%a)
exit /b

::  Get channel name for the key which was extracted from pkeyhelper.dll

:k_pkeychannel

set k=%1
set m=[Runtime.InteropServices.Marshal]
set p=%SysPath%\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('PidGenX', 'pidgenx.dll', 'Public, Static', 1, [int], @([String], [String], [String], [int], [IntPtr], [IntPtr], [IntPtr]), 1, 3);
set d1=%d1% $r = [byte[]]::new(0x04F8); $r[0] = 0xF8; $r[1] = 0x04; $f = %m%::AllocHGlobal(0x04F8); %m%::Copy($r, 0, $f, 0x04F8);
set d1=%d1% [void]$TypeBuilder.CreateType()::PidGenX('%k%', '%p%', '00000', 0, 0, 0, $f); %m%::Copy($f, $r, 0, 0x04F8); %m%::FreeHGlobal($f); [Text.Encoding]::Unicode.GetString($r, 1016, 128)

set pkeychannel=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkeychannel=%%a)
exit /b

:k_gvlk

for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b
for %%# in (Volume:GVLK) do (
call :k_pkey %osSKU% '%%#'
if defined pkey call :k_pkeychannel !pkey!
if /i "!pkeychannel!"=="%%#" (
set key=!pkey!
exit /b
)
)
exit /b

::========================================================================================================================================

::  1st column = Activation ID
::  2nd column = GVLK (Generic volume licensing key)
::  3rd column = SKU ID
::  4th column = WMI Edition ID (For reference only)
::  5th column = Build Branch name incase same Edition ID is used in different OS versions with different key (For reference only)
::  Separator  = "_"

:kms38data

set f=
for %%# in (
:: Windows 10/11
73111121-5638-40f6-bc11-f1d7b0d64300_NPPR9-FWDCX-D2C8J-H872K-2Y%f%T43___4_Enterprise
e272e3e2-732f-4c65-a8f0-484747d0d947_DPH2V-TTNVB-4X9Q3-TJR4H-KH%f%JW4__27_EnterpriseN
2de67392-b7a7-462a-b1ca-108dd189f588_W269N-WFGWX-YVC9B-4J6C9-T8%f%3GX__48_Professional
a80b5abf-76ad-428b-b05d-a47d2dffeebf_MH37W-N47XK-V7XM9-C7227-GC%f%QG9__49_ProfessionalN
7b9e1751-a8da-4f75-9560-5fadfe3d8e38_3KHY7-WNT83-DGQKR-F7HPR-84%f%4BM__98_CoreN
a9107544-f4a0-4053-a96a-1479abdef912_PVMJN-6DFY6-9CCP6-7BKTT-D3%f%WVR__99_CoreCountrySpecific
cd918a57-a41b-4c82-8dce-1a538e221a83_7HNRX-D7KGG-3K4RQ-4WPJ4-YT%f%DFH_100_CoreSingleLanguage
58e97c99-f377-4ef1-81d5-4ad5522b5fd8_TX9XD-98N7V-6WMQ6-BX7FG-H8%f%Q99_101_Core
e0c42288-980c-4788-a014-c080d2e1926e_NW6C2-QMPVW-D7KKK-3GKT6-VC%f%FB2_121_Education
3c102355-d027-42c6-ad23-2e7ef8a02585_2WH4N-8QGBV-H22JP-CT43Q-MD%f%WWJ_122_EducationN
32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7XTQ-FN8P6-TTKYV-9D4CC-J4%f%62D_125_EnterpriseS_RS5,VB,Ge
2d5a5a60-3040-48bf-beb0-fcd770c20ce0_DCPHK-NFMTC-H88MJ-PFHPY-QJ%f%4BJ_125_EnterpriseS_RS1
7b51a46c-0c04-4e8f-9af4-8496cca90d5e_WNMTR-4C88C-JK8YV-HQ7T2-76%f%DF9_125_EnterpriseS_TH1
7103a333-b8c8-49cc-93ce-d37c09687f92_92NFX-8DJQP-P6BBQ-THF9C-7C%f%G2H_126_EnterpriseSN_RS5,VB,Ge
9f776d83-7156-45b2-8a5c-359b9c9f22a3_QFFDN-GRT3P-VKWWX-X7T3R-8B%f%639_126_EnterpriseSN_RS1
87b838b7-41b6-4590-8318-5797951d8529_2F77B-TNFGY-69QQF-B8YKP-D6%f%9TJ_126_EnterpriseSN_TH1
82bbc092-bc50-4e16-8e18-b74fc486aec3_NRG8B-VKK3Q-CXVCJ-9G2XF-6Q%f%84J_161_ProfessionalWorkstation
4b1571d3-bafb-4b40-8087-a961be2caf65_9FNHH-K3HBT-3W4TD-6383H-6X%f%YWF_162_ProfessionalWorkstationN
3f1afc82-f8ac-4f6c-8005-1d233e606eee_6TP4R-GNPTD-KYYHQ-7B7DP-J4%f%47Y_164_ProfessionalEducation
5300b18c-2e33-4dc2-8291-47ffcec746dd_YVWGF-BXNMC-HTQYQ-CPQ99-66%f%QFC_165_ProfessionalEducationN
e0b2d383-d112-413f-8a80-97f373a5820c_YYVX9-NTFWV-6MDM3-9PT4T-4M%f%68B_171_EnterpriseG
e38454fb-41a4-4f59-a5dc-25080e354730_44RPN-FTY23-9VTTB-MP9BX-T8%f%4FV_172_EnterpriseGN
ec868e65-fadf-4759-b23e-93fe37f2cc29_CPWHC-NT2C7-VYW78-DHDB2-PG%f%3GK_175_ServerRdsh_RS5
e4db50ea-bda1-4566-b047-0ca50abc6f07_7NBT4-WGBQX-MP4H7-QXFF8-YP%f%3KX_175_ServerRdsh_RS3
0df4f814-3f57-4b8b-9a9d-fddadcd69fac_NBTWJ-3DR69-3C4V8-C26MC-GQ%f%9M6_183_CloudE
59eb965c-9150-42b7-a0ec-22151b9897c5_KBN8V-HFGQ4-MGXVD-347P6-PD%f%QGT_191_IoTEnterpriseS_VB,NI
d30136fc-cb4b-416e-a23d-87207abc44a9_6XN7V-PCBDC-BDBRH-8DQY7-G6%f%R44_202_CloudEditionN
ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69_37D7F-N49CB-WQR8W-TBJ73-FM%f%8RX_203_CloudEdition
:: Windows 2016/19/22/25 LTSC/SAC
7dc26449-db21-4e09-ba37-28f2958506a6_TVRH6-WHNXV-R9WG3-9XRFY-MY%f%832___7_ServerStandard_Ge
9bd77860-9b31-4b7b-96ad-2564017315bf_VDYBN-27WPP-V4HQT-9VMD4-VM%f%K7H___7_ServerStandard_FE
de32eafd-aaee-4662-9444-c1befb41bde2_N69G4-B89J2-4G8F4-WWYCC-J4%f%64C___7_ServerStandard_RS5
8c1c5410-9f39-4805-8c9d-63a07706358f_WC2BQ-8NRM3-FDDYY-2BFGV-KH%f%KQY___7_ServerStandard_RS1
c052f164-cdf6-409a-a0cb-853ba0f0f55a_D764K-2NDRG-47T6Q-P8T8W-YP%f%6DF___8_ServerDatacenter_Ge
ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03_WX4NM-KYWYW-QJJR4-XV3QB-6V%f%M33___8_ServerDatacenter_FE
34e1ae55-27f8-4950-8877-7a03be5fb181_WMDGN-G9PQG-XVVXX-R3X43-63%f%DFG___8_ServerDatacenter_RS5
21c56779-b449-4d20-adfc-eece0e1ad74b_CB7KF-BWN84-R7R2Y-793K2-8X%f%DDG___8_ServerDatacenter_RS1
034d3cbb-5d4b-4245-b3f8-f84571314078_WVDHN-86M7X-466P6-VHXV7-YY%f%726__50_ServerSolution_RS5
2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283_JCKRF-N37P4-C2D82-9YXRT-4M%f%63B__50_ServerSolution_RS1
7b4433f4-b1e7-4788-895a-c45378d38253_QN4C6-GBJD2-FB422-GHWJK-GJ%f%G2R_110_ServerCloudStorage
8de8eb62-bbe0-40ac-ac17-f75595071ea3_GRFBW-QNDC4-6QBHG-CCK3B-2P%f%R88_120_ServerARM64_RS5
43d9af6e-5e86-4be8-a797-d072a046896c_K9FYF-G6NCK-73M32-XMVPY-F9%f%DRR_120_ServerARM64_RS4
39e69c41-42b4-4a0a-abad-8e3c10a797cc_QFND9-D3Y9C-J3KKY-6RPVP-2D%f%PYV_145_ServerDatacenterACor_FE
90c362e5-0da1-4bfd-b53b-b87d309ade43_6NMRW-2C8FM-D24W7-TQWMY-CW%f%H2D_145_ServerDatacenterACor_RS5
e49c08e7-da82-42f8-bde2-b570fbcae76c_2HXDN-KRXHB-GPYC7-YCKFJ-7F%f%VDG_145_ServerDatacenterACor_RS3
f5e9429c-f50b-4b98-b15c-ef92eb5cff39_67KN8-4FYJW-2487Q-MQ2J7-4C%f%4RG_146_ServerStandardACor_FE
73e3957c-fc0c-400d-9184-5f7b6f2eb409_N2KJX-J94YW-TQVFB-DG9YT-72%f%4CC_146_ServerStandardACor_RS5
61c5ef22-f14f-4553-a824-c4b31e84b100_PTXN8-JFHJM-4WC78-MPCBR-9W%f%4KR_146_ServerStandardACor_RS3
45b5aff2-60a0-42f2-bc4b-ec6e5f7b527e_FCNV3-279Q9-BQB46-FTKXX-9H%f%PRH_168_ServerAzureCor_Ge
8c8f0ad3-9a43-4e05-b840-93b8d1475cbc_6N379-GGTMK-23C6M-XVVTC-CK%f%FRQ_168_ServerAzureCor_FE
a99cc1f0-7719-4306-9645-294102fbff95_FDNH6-VW9RW-BXPJ7-4XTYG-23%f%9TB_168_ServerAzureCor_RS5
3dbf341b-5f6c-4fa7-b936-699dce9e263f_VP34G-4NPPG-79JTQ-864T4-R3%f%MQX_168_ServerAzureCor_RS1
c2e946d1-cfa2-4523-8c87-30bc696ee584_XGN3F-F394H-FD2MY-PP6FD-8M%f%CRC_407_ServerTurbine_Ge
19b5e0fb-4431-46bc-bac1-2f1873e4ae73_NTBV8-9K7Q8-V27C6-M2BTV-KH%f%MXV_407_ServerTurbine_RS5
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do if %osSKU%==%%C (
if %1==key if not defined key echo "!allapps!" | find /i "%%A" %nul1% && set key=%%B
)
)
exit /b

::========================================================================================================================================

::  Below code is used to get alternate edition name and key if current edition doesn't support KMS38 activation

::  1st column = Current SKU ID
::  2nd column = Current Edition Name
::  3rd column = Current Edition Activation ID
::  4th column = Alternate Edition Activation ID
::  5th column = Alternate Edition GVLK
::  6th column = Alternate Edition Name
::  Separator  = _


:kms38fallback

set notfoundaltactID=
if %_NoEditionChange%==1 exit /b

for %%# in (
188_IoTEnterprise__________________8ab9bdd1-1f67-4997-82d9-8878520837d9_73111121-5638-40f6-bc11-f1d7b0d64300_NPPR9-FWDCX-D2C8J-H872K-2Y%f%T43_Enterprise
206_IoTEnterpriseK_________________80083eae-7031-4394-9e88-4901973d56fe_73111121-5638-40f6-bc11-f1d7b0d64300_NPPR9-FWDCX-D2C8J-H872K-2Y%f%T43_Enterprise
191_IoTEnterpriseS-2021____________ed655016-a9e8-4434-95d9-4345352c2552_32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7XTQ-FN8P6-TTKYV-9D4CC-J4%f%62D_EnterpriseS-2021
205_IoTEnterpriseSK________________d4f9b41f-205c-405e-8e08-3d16e88e02be_59eb965c-9150-42b7-a0ec-22151b9897c5_KBN8V-HFGQ4-MGXVD-347P6-PD%f%QGT_IoTEnterpriseS
138_ProfessionalSingleLanguage_____a48938aa-62fa-4966-9d44-9f04da3f72f2_2de67392-b7a7-462a-b1ca-108dd189f588_W269N-WFGWX-YVC9B-4J6C9-T8%f%3GX_Professional
139_ProfessionalCountrySpecific____f7af7d09-40e4-419c-a49b-eae366689ebd_2de67392-b7a7-462a-b1ca-108dd189f588_W269N-WFGWX-YVC9B-4J6C9-T8%f%3GX_Professional
139_ProfessionalCountrySpecific-Zn_01eb852c-424d-4060-94b8-c10d799d7364_2de67392-b7a7-462a-b1ca-108dd189f588_W269N-WFGWX-YVC9B-4J6C9-T8%f%3GX_Professional
) do (
for /f "tokens=1-6 delims=_" %%A in ("%%#") do if %osSKU%==%%A (
echo "!allapps!" | find /i "%%C" %nul1% && (
echo "!allapps!" | find /i "%%D" %nul1% && (
set altkey=%%E
set altedition=%%F
) || (
set altedition=%%F
set notfoundaltactID=1
)
)
)
)
exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:KMSActivation

::  To activate Windows with K-M-S activation, run the script with "/K-Windows" parameter or change 0 to 1 in below line
set _actwin=0

::  To activate all Office apps (including Project/Visio) with K-M-S activation, run the script with "/K-Office" parameter or change 0 to 1 in below line
set _actoff=0

::  To activate only Project/Visio with K-M-S activation, run the script with "/K-ProjectVisio" parameter or change 0 to 1 in below line
set _actprojvis=0

::  To activate all Windows/Office with K-M-S activation, run the script with "/K-WindowsOffice" parameter or change 0 to 1 in below line
set _actwinoff=0

::  To disable changing Windows/Office edition if current edition doesn't support K-M-S activation, run the script with "/K-NoEditionChange" parameter or change 0 to 1 in below line
set _NoEditionChange=0

::  To NOT auto-install renewal task with activation, run the script with "/K-NoRenewalTask" parameter or change 0 to 1 in below line
set _norentsk=0

::  To uninstall K-M-S, run the script with "/K-Uninstall" parameter or change 0 to 1 in below line. It'll take preference over any other parameter.
set _uni=0

::  Advanced options:
::  Don't use renewal task option if you are going to use a specific server name instead of public servers used in the script

::  To specify a server address for activation, run the script with "/K-Server-YOURKMSSERVERNAME" parameter or add it in below line after = sign
set _server=

::  To specify a port for activation, run the script with "/K-Port-YOURPORTNAME" parameter or add it in below line after = sign
set _port=

::  Script will run in unattended mode if parameters are used OR value is changed in above lines FOR activation or uninstall.

::========================================================================================================================================

cls
color 07
set KS=K%blank%MS
title  Online %KS% Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args for %%A in (%_args%) do (
if /i "%%A"=="-el"                        (set _elev=1)
if /i "%%A"=="/K-Windows"                 (set _actwin=1)
if /i "%%A"=="/K-Office"                  (set _actoff=1)
if /i "%%A"=="/K-ProjectVisio"            (set _actprojvis=1)
if /i "%%A"=="/K-WindowsOffice"           (set _actwinoff=1)
if /i "%%A"=="/K-NoEditionChange"         (set _NoEditionChange=1)
if /i "%%A"=="/K-NoRenewalTask"           (set _norentsk=1)
if /i "%%A"=="/K-Uninstall"               (set _uni=1)
echo "%%A" | find /i "/K-Port-"   >nul && (set "_port=%%A"   & call set "_port=%%_port:~8%%")
echo "%%A" | find /i "/K-Server-" >nul && (set "_server=%%A" & call set "_server=%%_server:~10%%")
)

for %%A in (%_actwin% %_actoff% %_actprojvis% %_actwinoff% %_uni%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

if %_uni%==1 goto :ks_uninstall

:ks_menu

if defined _server set _norentsk=1
if not defined _server set _port=

if %_unattended%==0 (
cls
if not defined terminal mode 76, 30
title  Online %KS% Activation %masver%

echo:
echo:
echo:
echo:
if exist "%ProgramFiles%\Activation-Renewal\Activation_task.cmd" (
find /i "Ver:2.7" "%ProgramFiles%\Activation-Renewal\Activation_task.cmd" %nul% || (
call :dk_color %_Yellow% "              Old renewal task found, run activation to update it."
)
)
echo        ______________________________________________________________
echo: 
echo               [1] Activate - Windows
echo               [2] Activate - Office [All]
echo               [3] Activate - Office [Project/Visio]
echo               [4] Activate - All
echo               _______________________________________________  
echo: 
if %_norentsk%==0 (
echo               [5] Renewal Task With Activation       [Yes]
) else (
call :dk_color2 %_White% "              [5] Renewal Task With Activation        " %_Yellow% "[No]"
)
if %_NoEditionChange%==0 (
echo               [6] Change Edition If Needed           [Yes]
) else (
call :dk_color2 %_White% "              [6] Change Edition If Needed            " %_Yellow% "[No]"
)
echo               [7] Uninstall Online %KS%
echo               _______________________________________________       
echo:
if defined _server (
echo               [8] Set %KS% Server/Port [%_server%] [%_port%]
) else (
echo               [8] Set %KS% Server/Port
)
echo               [9] Download Office
echo               [0] %_exitmsg%
echo        ______________________________________________________________
echo:
call :dk_color2 %_White% "       " %_Green% "Choose a menu option using your keyboard [1,2,3,4,5,6,7,8,9,0]"
choice /C:1234567890 /N
set _el=!errorlevel!

if !_el!==10 exit /b
if !_el!==9 start %mas%genuine-installation-media & goto :ks_menu
if !_el!==8 goto :ks_ip
if !_el!==7 cls & call :ks_uninstall & cls & goto :ks_menu
if !_el!==6 (if %_NoEditionChange%==0 (set _NoEditionChange=1) else (set _NoEditionChange=0)) & goto :ks_menu
if !_el!==5 (if %_norentsk%==0 (set _norentsk=1) else (set _norentsk=0)) & goto :ks_menu
if !_el!==4 cls & setlocal & set "_actwin=1" & set "_actoff=1" & set "_actprojvis=0" & call :ks_start & endlocal & cls & goto :ks_menu
if !_el!==3 cls & setlocal & set "_actwin=0" & set "_actoff=0" & set "_actprojvis=1" & call :ks_start & endlocal & cls & goto :ks_menu
if !_el!==2 cls & setlocal & set "_actwin=0" & set "_actoff=1" & set "_actprojvis=0" & call :ks_start & endlocal & cls & goto :ks_menu
if !_el!==1 cls & setlocal & set "_actwin=1" & set "_actoff=0" & set "_actprojvis=0" & call :ks_start & endlocal & cls & goto :ks_menu
goto :ks_menu
)

::========================================================================================================================================

:ks_start

cls
if not defined terminal (
mode 115, 32
if exist "%SysPath%\spp\store_test\" mode 135, 32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=32;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)
title  Online %KS% Activation %masver%

echo:
echo Initializing...
call :dk_chkmal

if not exist %SysPath%\sppsvc.exe (
%eline%
echo [%SysPath%\sppsvc.exe] file is missing, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

if %_actprojvis%==1 (set "_actoff=1")
if %_actwinoff%==1 (set "_actwin=1" & set "_actoff=1")

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_ckeckwmic
call :dk_checksku
call :dk_product
call :dk_sppissue

::========================================================================================================================================

set error=

cls
echo:
call :dk_showosinfo

::  Check Internet connection

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not "%%#"=="" set _int=1)
)

if not defined _int (
%psc% "If([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet){Exit 0}Else{Exit 1}"
if !errorlevel!==0 (set _int=1&set ping_f= But Ping Failed)
)

if defined _int (
echo Checking Internet Connection            [Connected%ping_f%]
) else (
set error=1
call :dk_color %Red% "Checking Internet Connection            [Not Connected]"
call :dk_color %Blue% "Internet is required for Online %KS% Activation."
)

::========================================================================================================================================

echo Initiating Diagnostic Tests...

set "_serv=sppsvc Winmgmt"

::  Software Protection
::  Windows Management Instrumentation

if %_actwin%==0 set notwinact=1
call :dk_errorcheck

::========================================================================================================================================

call :_taskclear-cache
call :_tasksetserv

if not %_actwin%==1 goto :ks_office

::  Process Windows
::  Check if system is permanently activated or not

echo:
echo Processing Windows...
call :dk_checkperm
if defined _perm (
call :dk_color %Gray% "Checking OS Activation                  [Windows is already permanently activated]"
goto :ks_office
)

::  Check Evaluation version

set _eval=
set _evalserv=

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _eval=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" set _evalserv=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" set _eval=1 & set _evalserv=1

if defined _eval (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% && (
call :dk_color %Red% "Checking Evaluation Edition             [Evaluation editions cannot be activated outside of evaluation period.]"

if defined _evalserv (
call :dk_color %Blue% "Go back to main menu and use [Change Edition] option."
) else (
call :dk_color %Blue% "Use TSforge activation option from the main menu to reset evaluation period."
set fixes=%fixes% %mas%evaluation_editions
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%evaluation_editions"
)

goto :ks_office
)
)

::========================================================================================================================================

::  Check if GVLK is already installed or not

call :k_channel

::  Detect Key

set key=
set pkey=
set altkey=
set changekey=
set altedition=

call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f
if defined allapps call :ksdata winkey
if not defined key call :k_gvlk %nul%
if defined allapps if not defined key call :kms38fallback

if defined altkey (set key=%altkey%&set changekey=1)

set /a UBR=0
if %osSKU%==191 if defined altkey if defined altedition (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR %nul6%') do if not errorlevel 1 set /a UBR=%%b
if %winbuild% LSS 22598 if !UBR! LSS 2788 (
call :dk_color %Blue% "Windows must be updated to build 19044.2788 or higher for IotEnterpriseS %KS% activation."
)
)

if not defined key if defined notfoundaltactID (
call :dk_color %Red% "Checking Alternate Edition For %KS%      [%altedition% Activation ID Not Found]"
)

if not defined key if not defined _gvlk (
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]

if %winbuild% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%osedition%\*GVLK*.xrm-ms" set sppks=1
if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%osedition%\*VLKMS*.xrm-ms" set sppks=1
if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%osedition%\*VL-BYPASS*.xrm-ms" set sppks=1

if defined skunotfound (
call :dk_color %Red% "Required license files not found in %SysPath%\spp\tokens\skus\"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

if defined sppks (
call :dk_color %Red% "%KS% activation is supported but failed to find the %KS% key."
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

if not defined skunotfound if not defined sppks (
call :dk_color %Red% "This product does not support %KS% activation."
call :dk_color %Blue% "Use TSforge activation option from the main menu."
)
echo:
goto :ks_office
)

::========================================================================================================================================

::  Install key

if defined changekey (
call :dk_color %Blue% "[%altedition%] edition product key will be used to enable %KS% activation."
echo:
)

if defined winsub (
call :dk_color %Blue% "Windows Subscription [SKU ID-%slcSKU%] found. Script will activate base edition [SKU ID-%regSKU%]."
echo:
)

set _partial=
if not defined key (
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path %spp% where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null AND LicenseDependsOn is NULL" Get PartialProductKey /value %nul6%') do set "_partial=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT PartialProductKey FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').Get()).PartialProductKey | %% {echo ('PartialProductKey='+$_)}" %nul6%') do set "_partial=%%#"
call echo Checking Installed Product Key          [Partial Key - %%_partial%%] [Volume:GVLK]
)

if defined key (
call :dk_inskey "[%key%]"
)

::========================================================================================================================================

:ks_office

if not %_actoff%==1 goto :ks_activate

call :ks_setspp

::  Check ohook install

set ohook=
for %%# in (15 16) do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" set ohook=1
)
)

for %%# in (System SystemX86) do (
for %%G in ("Office 15" "Office") do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" set ohook=1
)
)
)

if defined ohook (
echo:
call :dk_color %Gray% "Checking Ohook                          [Ohook activation is already installed for Office]"
)

::  Check unsupported office versions

set o14c2r=
set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office
%nul% reg query %_68%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 
%nul% reg query %_86%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 

if not "%o14c2r%"=="" (
echo:
call :dk_color %Red% "Checking Unsupported Office Install     [ %o14c2r%]"
)

if %winbuild% GEQ 10240 %psc% "Get-AppxPackage -name "Microsoft.MicrosoftOfficeHub"" | find /i "Office" %nul1% && (
set ohub=1
)

::========================================================================================================================================

::  Check supported office versions

call :ks_getpath

set o16uwp=
set o16uwp_path=

if %winbuild% GEQ 10240 (
for /f "delims=" %%a in ('%psc% "(Get-AppxPackage -name 'Microsoft.Office.Desktop' | Select-Object -ExpandProperty InstallLocation)" %nul6%') do (if exist "%%a\Integration\Integrator.exe" (set o16uwp=1&set "o16uwp_path=%%a"))
)

sc query ClickToRunSvc %nul%
set error1=%errorlevel%

if defined o16c2r if %error1% EQU 1060 (
echo:
call :dk_color %Red% "Checking ClickToRun Service             [Not found, Office 16.0 files found]"
set o16c2r=
set error=1
)

sc query OfficeSvc %nul%
set error2=%errorlevel%

if defined o15c2r if %error1% EQU 1060 if %error2% EQU 1060 (
echo:
call :dk_color %Red% "Checking ClickToRun Service             [Not found, Office 15.0 files found]"
set o15c2r=
set error=1
)

if "%o16uwp%%o16c2r%%o15c2r%%o16msi%%o15msi%%o14msi%"=="" (
set error=1
echo:
if not "%o14c2r%"=="" (
call :dk_color %Red% "Checking Supported Office Install       [Not Found]"
) else (
call :dk_color %Red% "Checking Installed Office               [Not Found]"
)

if defined ohub (
echo:
echo You have only Office dashboard app installed, you need to install full Office version.
)
call :dk_color %Blue% "Download and install Office from below URL and try again."
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto :ks_activate
)

set multioffice=
if not "%o16uwp%%o16c2r%%o15c2r%%o16msi%%o15msi%%o14msi%"=="1" set multioffice=1
if not "%o14c2r%"=="" set multioffice=1

if defined multioffice (
echo:
call :dk_color %Gray% "Checking Multiple Office Install        [Found. Recommended to install one version only]"
)

::========================================================================================================================================

::  Process Office UWP

if not defined o16uwp goto :ks_starto15c2r

call :ks_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=16
set "_oLPath=%o16uwp_path%\Licenses16"
for /f "delims=" %%a in ('%psc% "(Get-AppxPackage -name 'Microsoft.Office.Desktop' | Select-Object -ExpandProperty Dependencies) | Select-Object PackageFullName" %nul6%') do (set "o16uwpapplist=!o16uwpapplist! %%a")

echo "%o16uwpapplist%" | findstr /i "Access Excel OneNote Outlook PowerPoint Publisher SkypeForBusiness Word" %nul% && set "_oIds=O365HomePremRetail"

for %%# in (Project Visio) do (
echo "%o16uwpapplist%" | findstr /i "%%#" %nul% && (
set _lat=
if exist "%_oLPath%\%%#Pro2024VL*.xrm-ms" set "_oIds= !_oIds! %%#Pro2024Retail " & set _lat=1
if not defined _lat if exist "%_oLPath%\%%#Pro2021VL*.xrm-ms" set "_oIds= !_oIds! %%#Pro2021Retail " & set _lat=1
if not defined _lat if exist "%_oLPath%\%%#Pro2019VL*.xrm-ms" set "_oIds= !_oIds! %%#Pro2019Retail " & set _lat=1
if not defined _lat set "_oIds= !_oIds! %%#ProRetail "
)
)

set uwpinfo=%o16uwp_path:C:\Program Files\WindowsApps\Microsoft.Office.Desktop_=%

echo:
echo Processing Office...                    [UWP ^| %uwpinfo%]

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :ks_starto15c2r
)

call :ks_process

::========================================================================================================================================

:ks_starto15c2r

::  Process Office 15.0 C2R

if not defined o15c2r goto :ks_starto16c2r

call :ks_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=15
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v ProductReleaseIds" %nul6%') do (set "_prids=%o15c2r_reg%\Configuration /v ProductReleaseIds" & set "_config=%o15c2r_reg%\Configuration")
if not defined _oArch   for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v Platform" %nul6%') do (set "_oArch=%%b")
if not defined _version for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v version" %nul6%') do (set "_version=%%b")
if not defined _prids   for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v ProductReleaseId" %nul6%') do (set "_prids=%o15c2r_reg%\propertyBag /v ProductReleaseId" & set "_config=%o15c2r_reg%\propertyBag")

echo "%o15c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=10) || (set _tok=9)
for /f "tokens=%_tok% delims=\" %%a in ('reg query %o15c2r_reg%\ProductReleaseIDs\Active %nul6% ^| findstr /i "Retail Volume"') do (
echo "!_oIds!" | find /i " %%a " %nul1% || (set "_oIds= !_oIds! %%a ")
)

set "_oLPath=%_oRoot%\Licenses"
set "_oIntegrator=%_oRoot%\integration\integrator.exe"

echo:
echo Processing Office...                    [C2R ^| %_version% ^| %_oArch%]

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :ks_starto16c2r
)

if "%_actprojvis%"=="0" call :oh_fixprids
call :ks_process

::========================================================================================================================================

:ks_starto16c2r

::  Process Office 16.0 C2R

if not defined o16c2r goto :ks_startmsi

call :ks_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=16
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v AudienceData" %nul6%') do (set "_AudienceData=^| %%b ")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ProductReleaseIds" %nul6%') do (set "_prids=%o16c2r_reg%\Configuration /v ProductReleaseIds" & set "_config=%o16c2r_reg%\Configuration")

echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=9) || (set _tok=8)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs" /s /f ".16" /k %nul6% ^| findstr /i "Retail Volume"') do (
echo "!_oIds!" | find /i " %%a " %nul1% || (set "_oIds= !_oIds! %%a ")
)
set _oIds=%_oIds:.16=%
set _o16c2rIds=%_oIds%

set "_oLPath=%_oRoot%\Licenses16"
set "_oIntegrator=%_oRoot%\integration\integrator.exe"

echo:
echo Processing Office...                    [C2R ^| %_version% %_AudienceData%^| %_oArch%]

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :ks_startmsi
)

if "%_actprojvis%"=="0" call :oh_fixprids
call :ks_process

::========================================================================================================================================

:ks_startmsi

if defined o14msi call :ks_setspp 14
if defined o14msi call :ks_processmsi 14 %o14msi_reg%
call :ks_setspp
if defined o15msi call :ks_processmsi 15 %o15msi_reg%
if defined o16msi call :ks_processmsi 16 %o16msi_reg%

::========================================================================================================================================

echo:
call :oh_clearblock
if "%o16msi%%o15msi%"=="" if not "%o16uwp%%o16c2r%%o15c2r%"=="" if "%keyerror%"=="0" if %_NoEditionChange%==0 call :oh_uninstkey
call :oh_licrefresh

::========================================================================================================================================

:ks_activate

::  Opt out of sending KMSclient activation data to Microsoft
::  https://learn.microsoft.com/windows/privacy/manage-connections-from-windows-operating-system-components-to-microsoft-services#19-software-protection-platform

if %winbuild% GEQ 9600 (
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f %nul%
if %winbuild% EQU 14393 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoAcquireGT /t REG_DWORD /d 1 /f %nul%
echo Turn off %KS% AVS Validation             [Successful]
)

set "slp=SoftwareLicensingProduct"
set "ospp=OfficeSoftwareProtectionProduct"

echo:
echo Activating Volume Products...
if %_actwin%==1 call :_taskgetids sppwid %slp% windows
if %_actoff%==1 call :_taskgetids sppoid %slp% office
if %_actoff%==1 call :_taskgetids osppid %ospp% office

if not defined sppwid if not defined sppoid if not defined osppid (
if not defined keyerror (
echo No installed Volume Windows / Office products found.
) else (
call :dk_color %Red% "Failed to get installed Volume Windows / Office products."
)
call :_taskgetserv
call :_taskregserv
)

call :_taskact
if not defined showfix if defined _tserror (call :dk_color %Blue% "%_fixmsg%" & set showfix=1)

::  Don't create renewal task if Windows/Office volume IDs are not found, even if script is set to create it by default
::  Don't create renewal task if only Windows volume ID is found and OEM BIOS error is present on Windows 7, even if script is set to create it by default

set _deltask=
if not %_norentsk%==0 set _deltask=1
if not defined _deltask (
if %_actwin%==0 call :_taskgetids sppwid %slp% windows
if %_actoff%==0 call :_taskgetids sppoid %slp% office
if %_actoff%==0 call :_taskgetids osppid %ospp% office
)

if not defined sppwid if not defined sppoid if not defined osppid (set _deltask=1)
if defined oemerr if not defined sppoid if not defined osppid (set _deltask=1)

if not defined _deltask (
call :ks_renewal
) else (
if exist "%ProgramFiles%\Activation-Renewal\Activation_task.cmd" call :dk_color %Gray% "Deleting activation renewal task..."
call :dk_color %Gray% "Skipping creation of activation renewal task..."
call :ks_clearstuff %nul%
if not defined _server (
if %winbuild% GEQ 9200 (
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath /reg:32" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
if defined _C16R (
REM  mass grave[.]dev/office-license-is-not-genuine
set _server=10.0.0.10
call :_taskregserv
echo Keeping the non-existent IP address 10.0.0.10 as %KS% Server.
)
)
if not defined _C16R (
call :_taskclear-cache
echo Cleared %KS% Server from the registry.
)
)
)

::  https://learn.microsoft.com/azure/virtual-desktop/windows-10-multisession-faq

if %_actwin%==1 for %%# in (407) do if %osSKU%==%%# (
call :dk_color %Red% "%winos% does not support activation on non-azure platforms."
)

if %_actoff%==1 if defined sppoid if not defined _tserror if %_NoEditionChange%==0 if defined ohub (
echo:
call :dk_color %Gray% "Office apps such as Word, Excel are activated, use them directly. Ignore 'Buy' button in Office dashboard app."
)

::  Trigger reevaluation of SPP's Scheduled Tasks

call :dk_reeval %nul%
goto :dk_done

::========================================================================================================================================

:ks_ip

cls
set _server=
echo:
echo Enter / Paste the %KS% Server address, or just press Enter to return:
echo:
set /p _server=
if not defined _server goto :ks_menu
set "_server=%_server: =%"

echo:
echo Enter / Paste the %KS% Port address, or just press Enter to use default:
echo:
set /p _port=
if not defined _port goto :ks_menu
set "_port=%_port: =%"

goto :ks_menu

::========================================================================================================================================

:ks_reset

set key=
set _oRoot=
set _oArch=
set _oIds=
set _oLPath=
set _actid=
set _prod=
set _lic=
set _arr=
set _prids=
set _config=
set _version=
set _License=
set _oBranding=
exit /b

::========================================================================================================================================

:ks_getpath

set o16c2r=
set o15c2r=
set o16msi=
set o15msi=
set o14msi=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set o16c2r=1&set o16c2r_reg=%_86%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set o16c2r=1&set o16c2r_reg=%_68%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set o15c2r=1&set o15c2r_reg=%_86%\15.0\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set o15c2r=1&set o15c2r_reg=%_68%\15.0\ClickToRun)

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o16msi=1&set o16msi_reg=%_86%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o16msi=1&set o16msi_reg=%_68%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o15msi=1&set o15msi_reg=%_86%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o15msi=1&set o15msi_reg=%_68%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o14msi=1&set o14msi_reg=%_86%\14.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o14msi=1&set o14msi_reg=%_68%\14.0)

exit /b

::========================================================================================================================================

::  After retail to volume conversion, new product ID needs .OSPPReady key in registry, otherwise product info may not fully reflect 

:ks_osppready

if not defined _config exit /b

echo: %_config% | find /i "propertyBag" %nul1% && (
set "_osppt=REG_DWORD"
set "_osppready=%o15c2r_reg%"
) || (
set "_osppt=REG_SZ"
set "_osppready=%_config%"
)

reg add %_osppready% /f /v %_altoffid%.OSPPReady /t %_osppt% /d 1 %nul1%

::  Office builds before 16.0.10730.20102 need the Installed license product ID in ProductReleaseIds, otherwise product info may not fully reflect 

if exist "%_oLPath%\Word2019VL_KMS_Client_AE*.xrm-ms" exit /b

reg query %_prids% | findstr /I "%_altoffid%" %nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_prids%') do reg add %_prids% /t REG_SZ /d "%%b,%_altoffid%" /f %nul1%
)
exit /b

::========================================================================================================================================

:ks_setspp

if %winbuild% GEQ 9200 (
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
) else (
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
)
if "%1"=="14" (
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
)
exit /b

::========================================================================================================================================

:ks_process

for %%# in (%_oIds%) do (

set skipprocess=
if %_NoEditionChange%==1 if not defined _oBranding (
set foundprod=
call :ksdata chkprod %%#
if not defined foundprod (
set skipprocess=1
call :dk_color %Gray% "Skipping Because NoEditionChange Mode   [%%#]"
)
)


if "%_actprojvis%"=="1" if not defined skipprocess (
echo %%# | findstr /i "Project Visio" %nul% || (
set skipprocess=1
call :dk_color %Gray% "Skipping Because Project/Visio Mode     [%%#]"
)
)

if "%_actprojvis%"=="0" if not defined skipprocess echo %_oIds% | findstr /i "O365" %nul% && (
echo %%# | findstr /i "Project Visio" %nul% && (
set skipprocess=1
echo Skipping Because Mondo Is Available     [%%#]
)
)

if not defined skipprocess (
set key=
set _actid=
set _preview=
set _License=%%#
set _altoffid=

echo %%# | find /i "2024" %nul% && (
if exist "!_oLPath!\ProPlus2024PreviewVL_*.xrm-ms" if not exist "!_oLPath!\ProPlus2024VL_*.xrm-ms" set _preview=-Preview
)
set _prod=%%#!_preview!

call :ksdata getinfo !_prod!

if defined _altoffid (
set _License=!_altoffid!
echo Converting Retail To Volume             [!_prod! To !_altoffid!]
set _prod=!_altoffid!
call :ks_osppready
)

if not "!key!"=="" (
echo "!allapps!" | find /i "!_actid!" %nul1% || call :oh_installlic
call :dk_inskey "[!key!] [!_prod!]"
) else (
if not defined _oBranding (
set error=1
call :dk_color %Red% "Checking Product In Script              [Office %oVer%.0 !_prod! not found in script]"
call :dk_color %Blue% "Make sure you are using Latest MAS script."
) else (
call :dk_color %Red% "Checking Product In Script              [!_prod! MSI Retail is not supported]"
call :dk_color %Blue% "Uninstall this and Install C2R or MSI VL version of Office."
)
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
)
)
)

exit /b

::========================================================================================================================================

:ks_processmsi

::  Process Office MSI Version

call :ks_reset

if "%1"=="14" (
call :dk_actids 59a52881-a989-479d-af46-f275c6370663
) else (
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663
)

set oVer=%1
for /f "skip=2 tokens=2*" %%a in ('"reg query %2\Common\InstallRoot /v Path" %nul6%') do (set "_oRoot=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %2\Common\ProductVersion /v LastProduct" %nul6%') do (set "_version=%%b")
if "%_oRoot:~-1%"=="\" set "_oRoot=%_oRoot:~0,-1%"

echo "%2" | find /i "Wow6432Node" %nul1% && set _oArch=x86
if not "%osarch%"=="x86" if not defined _oArch set _oArch=x64
if "%osarch%"=="x86" set _oArch=x86

set "_common=%CommonProgramFiles%"
if defined PROCESSOR_ARCHITEW6432 set "_common=%CommonProgramW6432%"
set "_common2=%CommonProgramFiles(x86)%"

for /r "%_common%\Microsoft Shared\OFFICE%oVer%\" %%f in (BRANDING.XML) do if exist "%%f" set "_oBranding=%%f"
if not defined _oBranding for /r "%_common2%\Microsoft Shared\OFFICE%oVer%\" %%f in (BRANDING.XML) do if exist "%%f" set "_oBranding=%%f"

call :ksdata getmsiprod %2
call :ks_msiretaildata getmsiret %2

echo:
echo Processing Office...                    [MSI ^| %_version% ^| %_oArch%]

if not defined _oBranding (
set error=1
call :dk_color %Red% "Checking BRANDING.XML                   [Not Found. Aborting activation...]"
exit /b
)

if not defined _oIds (
set error=1
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
exit /b
)

call :ks_process
exit /b

::========================================================================================================================================

:ks_uninstall

cls
if not defined terminal mode 91, 30
title  Online %KS% Complete Uninstall %masver%

set "uline=__________________________________________________________________________________________"

set "_C16R="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath /reg:32" 2^>nul') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
if %winbuild% GEQ 9200 if defined _C16R (
echo:
call :dk_color %Gray% "Notice-"
echo:
echo To make sure Office programs do not show a non-genuine banner,
echo please run the activation option once, and don't uninstall afterward.
echo %uline%
)

echo:
set error_=
call :_taskclear-cache
call :ks_clearstuff

:: check KMS38 lock

%nul% reg query "HKLM\%SPPk%\%_wApp%" && (
set error_=9
echo Failed to completely clear %KS% Cache.
reg query "HKLM\%SPPk%\%_wApp%" /s %nul2% | findstr /i "127.0.0.2" %nul1% && echo KMS38 activation is locked.
) || (
echo Cleared %KS% Cache successfully.
)

if defined error_ (
if "%error_%"=="1" (
echo %uline%
%eline%
echo Try Again / Restart the System
echo %uline%
)
) else (
echo %uline%
echo:
call :dk_color %Green% "Online %KS% has been successfully uninstalled."
echo %uline%
)

goto :dk_done

:ks_clearstuff

set "key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks"

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" %nul1% && (
echo Deleting [Task] Activation-Renewal
schtasks /delete /tn Activation-Renewal /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" %nul1% && (
echo Deleting [Task] Activation-Run_Once
schtasks /delete /tn Activation-Run_Once /f %nul%
)

If exist "%ProgramFiles%\Activation-Renewal\" (
echo Deleting [Folder] %ProgramFiles%\Activation-Renewal\
rmdir /s /q "%ProgramFiles%\Activation-Renewal\" %nul%
)

::  Stuff from old MAS versions

schtasks /delete /tn Online_%KS%_Activation_Script-Renewal /f %nul%
schtasks /delete /tn Online_%KS%_Activation_Script-Run_Once /f %nul%
del /f /q "%ProgramData%\Online_%KS%_Activation.cmd" %nul%
rmdir /s /q "%ProgramData%\Activation-Renewal\" %nul%
rmdir /s /q "%ProgramData%\Online_%KS%_Activation\" %nul%
rmdir /s /q "%windir%\Online_%KS%_Activation_Script\" %nul%
reg delete "HKCR\DesktopBackground\shell\Activate Windows - Office" /f %nul%

::  Check if all is removed

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" %nul1% && (set error_=1)
reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" %nul1% && (set error_=1)
reg query "%key%" /f Path /s | find /i "\Online_%KS%_Activation_Script" %nul1% && (set error_=1)
If exist "%windir%\Online_%KS%_Activation_Script\" (set error_=1)
reg query "HKCR\DesktopBackground\shell\Activate Windows - Office" %nul% && (set error_=1)
if exist "%ProgramData%\Online_%KS%_Activation.cmd" (set error_=1)
if exist "%ProgramData%\Online_%KS%_Activation\" (set error_=1)
if exist "%ProgramData%\Activation-Renewal\" (set error_=1)
if exist "%ProgramFiles%\Activation-Renewal\" (set error_=1)
exit /b

::========================================================================================================================================

:_extracttask:
@echo off

::   Renew K-M-S activation with Online servers via scheduled task

::============================================================================
::
::   Homepage: mass grave[.]dev
::      Email: mas.help@outlook.com
::
::============================================================================


if not "%~1"=="Task" (
echo:
echo ====== Error ======
echo:
echo This file is supposed to be run only by the scheduled task.
echo:
echo Press any key to exit
pause >nul
exit /b
)

::  Set Environment variables, it helps if they are misconfigured in the system

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

>nul fltmc || exit /b

::========================================================================================================================================

set _tserror=
set winbuild=1
set "nul=>nul 2>&1"
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set psc=powershell.exe

set run_once=
set t_name=Renewal Task
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks" /f Path /s | find /i "\Activation-Run_Once" >nul && (
set run_once=1
set t_name=Run Once Task
)

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
cmd /c "wmic path Win32_ComputerSystem get CreationClassName /value" 2>nul | find /i "computersystem" 1>nul && set _wmic=1
)

setlocal EnableDelayedExpansion
if exist "%ProgramFiles%\Activation-Renewal\" call :_taskstart>>"%ProgramFiles%\Activation-Renewal\Logs.txt"
exit

::========================================================================================================================================

:_taskstart

echo:
echo %date%, %time%

set /a loop=1
set /a max_loop=4

call :_tasksetserv

:_intrepeat

::  Check Internet connection. Works even if ICMP echo is disabled.

for %%a in (%srvlist%) do (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (
if not "%%#"=="" goto _taskIntConnected
)
)

nslookup dns.msftncsi.com 2>nul | find "131.107.255.255" 1>nul
if "%errorlevel%"=="0" goto _taskIntConnected

if %loop%==%max_loop% (
set _tserror=1
goto _taskend
)

echo:
echo Error: Internet is not connected
echo Waiting 30 seconds

timeout /t 30 >nul
set /a loop=%loop%+1
goto _intrepeat

:_taskIntConnected

::========================================================================================================================================

call :_taskclear-cache

::  Check WMI and sppsvc Errors

set applist=
net start sppsvc /y %nul%
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%_wApp%') get ID /VALUE" 2^>nul')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %slp% WHERE ApplicationID=''%_wApp%''').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))

if not defined applist (
set _tserror=1
if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" 2>nul | find /i "computersystem" 1>nul
if !errorlevel! NEQ 0 (set e_wmispp=WMI, SPP) else (set e_wmispp=SPP)
echo:
echo Error: Not Respoding- !e_wmispp!
echo:
)

::========================================================================================================================================

::  Check installed volume products activation ID's

call :_taskgetids sppwid %slp% windows
call :_taskgetids sppoid %slp% office
call :_taskgetids osppid %ospp% office

::========================================================================================================================================

echo:
echo Renewing K-M-S activation for all installed Volume products

if not defined sppwid if not defined sppoid if not defined osppid (
echo:
echo No installed Volume Windows / Office product found
echo:
echo Renewing K-M-S server
call :_taskgetserv
call :_taskregserv
goto :_skipact
)

::========================================================================================================================================

call :_taskact

:_skipact

::========================================================================================================================================

if defined run_once (
echo:
echo Deleting Scheduled Task Activation-Run_Once
schtasks /delete /tn Activation-Run_Once /f %nul%
)

::========================================================================================================================================

:_taskend

echo:
echo Exiting
echo ______________________________________________________________________

if defined _tserror (exit /b 123456789) else (exit /b 0)

::========================================================================================================================================

:_act

set prodname=
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('"wmic path !_path! where ID='!_actid!' get LicenseFamily /VALUE" 2^>nul') do (call set "prodname=%%#")
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM !_path! WHERE ID=''!_actid!''').Get()).LicenseFamily | %% {echo ('LicenseFamily='+$_)}" 2^>nul') do (call set "prodname=%%#")
for /f "tokens=1 delims=-_" %%a in ("%prodname%") do set "prodname=%%a"

set _taskskip=
if "%_actprojvis%"=="1" (
echo: %prodname% | find /i "Office" %nul% && (
echo: %prodname% | findstr /i "Project Visio" %nul% || (set _taskskip=1& exit /b)
)
)

if defined t_name Activating: %prodname%

set errorcode=12345
set /a act_attempt=0

:_act2

if %act_attempt% GTR 4 exit /b

if not "%act_ok%"=="1" (
if not defined _server call :_taskgetserv
call :_taskregserv
)

if not !server_num! GTR %max_servers% (

if "%1"=="act_win" if %_kms38% EQU 1 (
set act_ok=1
exit /b
)

if %_wmic% EQU 1 wmic path !_path! where ID='!_actid!' call Activate %nul%
if %_wmic% EQU 0 %psc% "try {$null=(([WMISEARCHER]'SELECT ID FROM !_path! where ID=''!_actid!''').Get()).Activate(); exit 0} catch { exit $_.Exception.InnerException.HResult }"

call set errorcode=!errorlevel!

if !errorcode! EQU 0 (
set act_ok=1
exit /b
)
if "%1"=="act_win" if !errorcode! EQU -1073418187 if %winbuild% LSS 9200 (
set act_ok=1
exit /b
)

set act_ok=0
set /a act_attempt+=1
if not defined _server goto _act2
)
exit /b

::========================================================================================================================================

:_actinfo

if "%1"=="act_win" if not defined t_name (set prodname=%winos%)

if "%1"=="act_win" if %_kms38% EQU 1 (
if defined t_name (
echo %prodname% is already activated with KMS38.
) else (
call :dk_color %Green% "%prodname% is already activated with KMS38."
)
exit /b
)

if %errorcode% EQU 12345 (
if defined t_name (
echo %prodname% activation failed due to restricted or no Internet.
) else (
call :dk_color %Red% "%prodname% activation failed due to restricted or no Internet."
)
set showfix=1
set _tserror=1
exit /b
)

if %errorcode% EQU -1073418187 if "%1"=="act_win" if %winbuild% LSS 9200 (
if defined t_name (
echo %prodname% cannot be KMS-activated on this computer due to unqualified OEM BIOS [0xC004F035].
) else (
call :dk_color %Red% "%prodname% cannot be KMS-activated on this computer due to unqualified OEM BIOS [0xC004F035]."
call :dk_color %Blue% "Use TSforge activation option from the main menu."
)
set oemerr=1
set showfix=1
exit /b
)

if %errorcode% EQU -1073418124 (
if defined t_name (
echo %prodname% activation failed due to Internet issue [0xC004F074].
) else (
call :dk_color %Red% "%prodname% activation failed due to Internet issue [0xC004F074]."
if not defined _tserror (
call :dk_color %Blue% "Make sure that system files are not blocked by firewall."
call :dk_color %Blue% "If the issue persists, try another Internet connection or VPN such as https://1.1.1.1"
)
)
set showfix=1
set _tserror=1
exit /b
)


set gpr=0
set gpr2=0
call :_taskgetgrace
set /a "gpr2=(%gpr%+1440-1)/1440"

if %errorcode% EQU 0 if %gpr% EQU 0 (
if defined t_name (
echo %prodname% activation succeeded, but Remaining Period failed to increase.
) else (
call :dk_color %Red% "%prodname% activation succeeded, but Remaining Period failed to increase."
)
set _tserror=1
exit /b
)

set _actpass=1
if %gpr% EQU 43200  if "%1"=="act_win" if %winbuild% GEQ 9200 set _actpass=0
if %gpr% EQU 64800  set _actpass=0
if %gpr% GTR 259200 if "%1"=="act_win" call :_taskchkEnterpriseG _actpass
if %gpr% EQU 259200 set _actpass=0

if %errorcode% EQU 0 if %_actpass% EQU 0 (
if defined t_name (
echo %prodname% is successfully activated for %gpr2% days.
) else (
call :dk_color %Green% "%prodname% is successfully activated for %gpr2% days."
)
exit /b
)

cmd /c exit /b %errorcode%
if defined t_name (
echo %prodname% has failed to activate [0x!=ExitCode!]. Remaining Period: %gpr2% days [%gpr% minutes].
) else (
call :dk_color %Red% "%prodname% has failed to activate [0x!=ExitCode!]. Remaining Period: %gpr2% days [%gpr% minutes]."
)
set _tserror=1
exit /b

::========================================================================================================================================

:_taskact

:: Check KMS38 activation

set gpr=0
set _kms38=0
if defined sppwid if %winbuild% GEQ 14393 (
set _path=%slp%
set _actid=%sppwid%
call :_taskgetgrace
)

if %gpr% NEQ 0 if %gpr% GTR 259200 (
set _kms38=1
call :_taskchkEnterpriseG _kms38
)

:: Set specific K-M-S host to Local Host so that global K-M-S IP can not replace KMS38 activation but can be used with Office and other Windows Editions.

if %_kms38% EQU 1 (
%nul% reg add "HKLM\%SPPk%\%_wApp%\%sppwid%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2"
%nul% reg add "HKLM\%SPPk%\%_wApp%\%sppwid%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)

echo:
if defined sppwid (
set _path=%slp%
set _actid=%sppwid%
call :_act act_win
call :_actinfo act_win
) else (
if defined t_name echo Checking: Volume version of Windows is not installed
)

if defined sppoid (
set _path=%slp%
for %%# in (%sppoid%) do (
set _actid=%%#
call :_act
if not defined _taskskip call :_actinfo
)
)

if defined osppid (
set _path=%ospp%
for %%# in (%osppid%) do (
set _actid=%%#
call :_act
if not defined _taskskip call :_actinfo
)
)

if not defined sppoid if not defined osppid if defined t_name (
echo:
echo Checking: Volume version of Office is not installed
)

exit /b

::========================================================================================================================================

:_taskgetids

set %1=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %2 where (Name like '%%%3%%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL AND LicenseDependsOn is NULL) get ID /VALUE" 2^>nul')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %2 WHERE Name like ''%%%3%%'' and Description like ''%%KMSCLIENT%%'' and PartialProductKey is not NULL AND LicenseDependsOn is NULL').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
%chkapp% do (if defined %1 (call set "%1=!%1! %%a") else (call set "%1=%%a"))
exit /b

:_taskgetgrace

set gpr=0
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('"wmic path !_path! where ID='!_actid!' get GracePeriodRemaining /VALUE" 2^>nul') do call set "gpr=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT GracePeriodRemaining FROM !_path! where ID=''!_actid!''').Get()).GracePeriodRemaining | %% {echo ('GracePeriodRemaining='+$_)}" 2^>nul') do call set "gpr=%%#"
exit /b

:_taskchkEnterpriseG

for %%# in (e0b2d383-d112-413f-8a80-97f373a5820c e38454fb-41a4-4f59-a5dc-25080e354730) do (if %sppwid%==%%# set %1=0)
exit /b

::========================================================================================================================================

::  Clean existing K-M-S cache from the registry

:_taskclear-cache

set w=
for %%# in (SppE%w%xtComObj.exe sppsvc.exe) do (
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%w%ge File Execu%w%tion Options\%%#" /f %nul%
)

set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"

set "slp=SoftwareLicensingProduct"
set "ospp=OfficeSoftwareProtectionProduct"

set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"

%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName /reg:32
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort /reg:32
%nul% reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
%nul% reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
%nul% reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
)
if %winbuild% GEQ 9600 (
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
%nul% reg delete "HKLM\%OPPk%" /f /v KeyManagementServiceName
%nul% reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
%nul% reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing
%nul% reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching
%nul% reg delete "HKLM\%OPPk%\%_oA14%" /f
%nul% reg delete "HKLM\%OPPk%\%_oApp%" /f

exit /b

::========================================================================================================================================

:_taskregserv

if defined _server (set KMS_IP=%_server%)
if not defined _port set _port=1688

%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%_port%"
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%_port%" /reg:32

%nul% reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"
%nul% reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%_port%"

if %winbuild% GEQ 9200 (
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "%_port%"
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "%_port%" /reg:32
)
exit /b

::========================================================================================================================================

:_tasksetserv

::  Multi K-M-S servers integration and servers randomization

set srvlist=
set -=

set "srvlist=kms.03%-%k.org kms-default.cangs%-%hui.net kms.six%-%yin.com kms.moe%-%club.org kms.cgt%-%soft.com"
set "srvlist=%srvlist% kms.id%-%ina.cn kms.moe%-%yuuko.com xinch%-%eng213618.cn kms.lol%-%i.best kms.my%-%ds.cloud"
set "srvlist=%srvlist% kms.0%-%t.net.cn win.k%-%ms.pub kms.wx%-%lost.com kms.moe%-%yuuko.top kms.gh%-%pym.com"

set n=1
for %%a in (%srvlist%) do (set %%a=&set server!n!=%%a&set /a n+=1)
set max_servers=15
set /a server_num=0
exit /b

:_taskgetserv

if %server_num% geq %max_servers% (set /a server_num+=1&set KMS_IP=222.184.9.98&exit /b)
set /a rand=%Random%%%(15+1-1)+1
if defined !server%rand%! goto :_taskgetserv
set KMS_IP=!server%rand%!
set !server%rand%!=1

::  Get IPv4 address of K-M-S server to use for the activation, works even if ICMP echo is disabled.
::  Microsoft and Antivirus's may flag the issue if public K-M-S server host name is directly used for the activation.

set /a server_num+=1
(for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%a"
if "%KMS_IP%"=="!KMS_IP!" for /f "delims=[] tokens=2" %%# in ('pathping -4 -h 1 -n -p 1 -q 1 -w 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%#"
if not "%KMS_IP%"=="!KMS_IP!" exit /b
goto :_taskgetserv
)
::Ver:2.7
:_extracttask:

::========================================================================================================================================

:ks_renewal

set error_=
set "_dest=%ProgramFiles%\Activation-Renewal"
set "key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks"

call :ks_clearstuff %nul%

if defined error_ (
set error=1
call :dk_color %Red% "Failed to remove previous Renewal Task. Restart system / Try again."
exit /b
)

if not exist "%_dest%\" md "%_dest%\" %nul%
for /f %%G in ('%psc% "[Guid]::NewGuid().Guid"') do set "randguid=%%G"
set "_temp=%SystemRoot%\Temp\%Random%%randguid%"

set nil=
if exist "%_temp%\.*" rmdir /s /q "%_temp%\" %nul%
md "%_temp%\" %nul%
call :ks_RenExport renewal "%_temp%\Renewal.xml" Unicode
if not defined _int (call :ks_RenExport run_once "%_temp%\Run_Once.xml" Unicode)
s%nil%cht%nil%asks /cre%nil%ate /tn "Activation-Renewal" /ru "SYS%nil%TEM" /xml "%_temp%\Renewal.xml" %nul%
if not defined _int (s%nil%cht%nil%asks /cre%nil%ate /tn "Activation-Run_Once" /ru "SYS%nil%TEM" /xml "%_temp%\Run_Once.xml" %nul%)
if exist "%_temp%\.*" rmdir /s /q "%_temp%\" %nul%

call :ks_createInfo.txt
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":_extracttask\:.*`r`n\"; [io.file]::WriteAllText('%_dest%\Activation_task.cmd', '@::%randguid%' + [Environment]::NewLine + $f[1].Trim(), [System.Text.Encoding]::ASCII)"

::========================================================================================================================================

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul || (set error_=1)
if not defined _int reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul || (set error_=1)

If not exist "%_dest%\Activation_task.cmd" (set error_=1)
If not exist "%_dest%\Info.txt" (set error_=1)

if defined error_ (
schtasks /delete /tn Activation-Renewal /f %nul%
schtasks /delete /tn Activation-Run_Once /f %nul%
rmdir /s /q "%_dest%\" %nul%
set error=1
call :dk_color %Red% "Failed to install Renewal Task. Restart system / Try again."
exit /b
)

if "%keyerror%"=="0" if not defined _tserror (
call :dk_color %Green% "Renewal Task for lifetime activation is successfully installed in %_dest%"
exit /b
)
echo Renewal Task for lifetime activation is successfully installed in %_dest%
exit /b

::  Extract the text from batch script without character and file encoding issue

:ks_RenExport

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('%~2',$f[1].Trim(),[System.Text.Encoding]::%~3);"
exit /b

::========================================================================================================================================

:ks_createInfo.txt

(
echo   The use of this script is to renew your Windows/Office license using online K-M-S.
echo:
echo   If renewal/activation Scheduled tasks were created then following would exist,
echo:
echo   - Scheduled tasks
echo     Activation-Renewal    [Renewal / Weekly]
echo     Activation-Run_Once   [Activation Task - deletes itself once activated]
echo     The scheduled tasks runs only if the system is connected to the Internet.
echo:
echo   - Files
echo     C:\Program Files\Activation-Renewal\Activation_task.cmd
echo     C:\Program Files\Activation-Renewal\Info.txt
echo     C:\Program Files\Activation-Renewal\Logs.txt
echo ______________________________________________________________________________________________
echo:
echo   This Script is a part of MAS project.
echo:   
echo   Homepage: mass grave[.]dev
echo      Email: mas.help@outlook.com
)>"%_dest%\Info.txt"
exit /b

::========================================================================================================================================

:renewal:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Source>Microsoft Corporation</Source>
    <Date>1999-01-01T12:00:00.34375</Date>
    <Author>WindowsAddict</Author>
    <Version>1.0</Version>
    <Description>Online K-M-S Activation-Renewal - Weekly Task</Description>
    <URI>\Activation-Renewal</URI>
    <SecurityDescriptor>D:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FRFX;;;LS)(A;;FRFW;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)(A;;FR;;;S-1-5-4)</SecurityDescriptor>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>1999-01-01T12:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Sunday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="LocalSystem">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT2M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="LocalSystem">
    <Exec>
      <Command>%ProgramFiles%\Activation-Renewal\Activation_task.cmd</Command>
    <Arguments>Task</Arguments>
    </Exec>
  </Actions>
</Task>
:renewal:

:run_once:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Source>Microsoft Corporation</Source>
    <Date>1999-01-01T12:00:00.34375</Date>
    <Author>WindowsAddict</Author>
    <Version>1.0</Version>
    <Description>Online K-M-S Activation Run Once - Run and Delete itself on first Internet Contact</Description>
    <URI>\Activation-Run_Once</URI>
    <SecurityDescriptor>D:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FRFX;;;LS)(A;;FRFW;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)(A;;FR;;;S-1-5-4)</SecurityDescriptor>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="LocalSystem">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT2M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="LocalSystem">
    <Exec>
      <Command>%ProgramFiles%\Activation-Renewal\Activation_task.cmd</Command>
    <Arguments>Task</Arguments>
    </Exec>
  </Actions>
</Task>
:run_once:

::========================================================================================================================================

::  1st column = Office version number
::  2nd column = Activation ID
::  3rd column = Edition
::  4th column = Other Edition IDs if they are part of the same primary product (For reference only)
::  Separator  = "_"

:ks_msiretaildata

for %%# in (
:: Office 2010
14_4d463c2c-0505-4626-8cdb-a4da82e2d8ed_AccessR
14_745fb377-0a59-4ca9-b9a9-c359557a2c4e_AccessRuntimeR
14_4eaff0d0-c6cb-4187-94f3-c7656d49a0aa_ExcelR
14_7004b7f0-6407-4f45-8eac-966e5f868bde_GrooveR
14_7b7d1f17-fdcb-4820-9789-9bec6e377821_HomeBusinessR_[HomeBusinessDemoR]
14_19316117-30a8-4773-8fd9-7f7231f4e060_HomeBusinessSubR
14_09e2d37e-474b-4121-8626-58ad9be5776f_HomeStudentR_[HomeStudentDemoR]
14_c3ae020c-5a71-4cc5-a27a-2a97c2d46860_HSExcelR
14_25fe4611-b44d-49cc-ae87-2143d299194e_HSOneNoteR
14_d652ad8d-da5c-4358-b928-7fb1b4de7a7c_HSPowerPointR
14_a963d7ae-7a88-41a7-94da-8bb5635a8af9_HSWordR
14_ef1da464-01c8-43a6-91af-e4e5713744f9_InfoPathR
14_14f5946a-debc-4716-babc-7e2c240fec08_MondoR
14_c1ceda8b-c578-4d5d-a4aa-23626be4e234_OEM
14_3f7aa693-9a7e-44fc-9309-bb3d8e604925_OneNoteR
14_fbf4ac36-31c8-4340-8666-79873129cf40_OutlookR
14_acb51361-c0db-4895-9497-1831c41f31a6_PersonalR_[PersonalDemoR,PersonalPrepaidR]
14_133c8359-4e93-4241-8118-30bb18737ea0_PowerPointR
14_8b559c37-0117-413e-921b-b853aeb6e210_ProfessionalR_[ProfessionalAcadR,ProfessionalDemoR]
14_725714d7-d58f-4d12-9fa8-35873c6f7215_ProjectProR_[ProjectProMSDNR]
14_4d06f72e-fd50-4bc2-a24b-d448d7f17ef2_ProjectProSubR
14_688f6589-2bd9-424e-a152-b13f36aa6de1_ProjectStdR
14_71af7e84-93e6-4363-9b69-699e04e74071_ProPlusR_[ProPlusAcadR,ProPlusMSDNR,Sub4R]
14_e98ef0c0-71c4-42ce-8305-287d8721e26c_ProPlusSubR
14_98677603-a668-4fa4-9980-3f1f05f78f69_PublisherR
14_dbe3aee0-5183-4ff7-8142-66050173cb01_SmallBusBasicsR_[SmallBusBasicsMSDNR]
14_b78df69e-0966-40b1-ae85-30a5134dedd0_SPDR
14_d3422cfb-8d8b-4ead-99f9-eab0ccd990d7_StandardR
14_2745e581-565a-4670-ae90-6bf7c57ffe43_StarterR
14_66cad568-c2dc-459d-93ec-2f3cb967ee34_VisioSIR_Prem[Pro,Std]
14_db3bbc9c-ce52-41d1-a46f-1a1d68059119_WordR
:: Office 2013
15_ab4d047b-97cf-4126-a69f-34df08e2f254_AccessRetail
15_259de5be-492b-44b3-9d78-9645f848f7b0_AccessRuntimeRetail
15_1b1d9bd5-12ea-4063-964c-16e7e87d6e08_ExcelRetail
15_cfaf5356-49e3-48a8-ab3c-e729ab791250_GrooveRetail
15_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_HomeBusinessPipcRetail
15_a2b90e7a-a797-4713-af90-f0becf52a1dd_HomeBusinessRetail
15_1fdfb4e4-f9c9-41c4-b055-c80daf00697d_HomeStudentARMRetail
15_ebef9f05-5273-404a-9253-c5e252f50555_HomeStudentPlusARMRetail
15_f2de350d-3028-410a-bfae-283e00b44d0e_HomeStudentRetail
15_44984381-406e-4a35-b1c3-e54f499556e2_InfoPathRetail
15_9103f3ce-1084-447a-827e-d6097f68c895_LyncAcademicRetail
15_ff693bf4-0276-4ddb-bb42-74ef1a0c9f4d_LyncEntryRetail
15_fada6658-bfc6-4c4e-825a-59a89822cda8_LyncRetail
15_69ec9152-153b-471a-bf35-77ec88683eae_MondoRetail
15_3391e125-f6e4-4b1e-899c-a25e6092d40d_OneNoteFreeRetail
15_8b524bcc-67ea-4876-a509-45e46f6347e8_OneNoteRetail
15_12004b48-e6c8-4ffa-ad5a-ac8d4467765a_OutlookRetail
15_5aab8561-1686-43f7-9ff5-2c861da58d17_PersonalPipcRetail
15_17e9df2d-ed91-4382-904b-4fed6a12caf0_PersonalRetail
15_31743b82-bfbc-44b6-aa12-85d42e644d5b_PowerPointRetail
15_064383fa-1538-491c-859b-0ecab169a0ab_ProPlusRetail
15_4e26cac1-e15a-4467-9069-cb47b67fe191_ProfessionalPipcRetail
15_44bc70e2-fb83-4b09-9082-e5557e0c2ede_ProfessionalRetail
15_2f72340c-b555-418d-8b46-355944fe66b8_ProjectProRetail
15_58d95b09-6af6-453d-a976-8ef0ae0316b1_ProjectStdRetail
15_c3a0814a-70a4-471f-af37-2313a6331111_PublisherRetail
15_ba3e3833-6a7e-445a-89d0-7802a9a68588_SPDRetail
15_32255c0a-16b4-4ce2-b388-8a4267e219eb_StandardRetail
15_a56a3b37-3a35-4bbb-a036-eee5f1898eee_VisioProRetail
15_980f9e3e-f5a8-41c8-8596-61404addf677_VisioStdRetail
15_191509f2-6977-456f-ab30-cf0492b1e93a_WordRetail
:: Office 365 - 15.0 version
15_6337137e-7c07-4197-8986-bece6a76fc33_O365BusinessRetail
15_537ea5b5-7d50-4876-bd38-a53a77caca32_O365HomePremRetail
15_149dbce7-a48e-44db-8364-a53386cd4580_O365ProPlusRetail
15_bacd4614-5bef-4a5e-bafc-de4c788037a2_O365SmallBusPremRetail
:: Office 365 - 16.0 version
16_6337137e-7c07-4197-8986-bece6a76fc33_O365BusinessRetail
16_2f5c71b4-5b7a-4005-bb68-f9fac26f2ea3_O365EduCloudRetail
16_537ea5b5-7d50-4876-bd38-a53a77caca32_O365HomePremRetail
16_149dbce7-a48e-44db-8364-a53386cd4580_O365ProPlusRetail
16_bacd4614-5bef-4a5e-bafc-de4c788037a2_O365SmallBusPremRetail
:: Office 2016
16_bfa358b0-98f1-4125-842e-585fa13032e6_AccessRetail
16_9d9faf9e-d345-4b49-afce-68cb0a539c7c_AccessRuntimeRetail
16_424d52ff-7ad2-4bc7-8ac6-748d767b455d_ExcelRetail
16_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_HomeBusinessPipcRetail
16_86834d00-7896-4a38-8fae-32f20b86fa2b_HomeBusinessRetail
16_c28acdb8-d8b3-4199-baa4-024d09e97c99_HomeStudentRetail
16_090896a0-ea98-48ac-b545-ba5da0eb0c9c_HomeStudentARMRetail
16_6bbe2077-01a4-4269-bf15-5bf4d8efc0b2_HomeStudentPlusARMRetail
16_e2127526-b60c-43e0-bed1-3c9dc3d5a468_HomeStudentVNextRetail
16_69ec9152-153b-471a-bf35-77ec88683eae_MondoRetail
16_436366de-5579-4f24-96db-3893e4400030_OneNoteFreeRetail
16_83ac4dd9-1b93-40ed-aa55-ede25bb6af38_OneNoteRetail
16_5a670809-0983-4c2d-8aad-d3c2c5b7d5d1_OutlookRetail
16_5aab8561-1686-43f7-9ff5-2c861da58d17_PersonalPipcRetail
16_a9f645a1-0d6a-4978-926a-abcb363b72a6_PersonalRetail
16_f32d1284-0792-49da-9ac6-deb2bc9c80b6_PowerPointRetail
16_de52bd50-9564-4adc-8fcb-a345c17f84f9_ProPlusRetail
16_4e26cac1-e15a-4467-9069-cb47b67fe191_ProfessionalPipcRetail
16_d64edc00-7453-4301-8428-197343fafb16_ProfessionalRetail
16_2f72340c-b555-418d-8b46-355944fe66b8_ProjectProRetail
16_58d95b09-6af6-453d-a976-8ef0ae0316b1_ProjectStdRetail
16_6e0c1d99-c72e-4968-bcb7-ab79e03e201e_PublisherRetail
16_9103f3ce-1084-447a-827e-d6097f68c895_SkypeServiceBypassRetail
16_971cd368-f2e1-49c1-aedd-330909ce18b6_SkypeforBusinessEntryRetail
16_418d2b9f-b491-4d7f-84f1-49e27cc66597_SkypeforBusinessRetail
16_4a31c291-3a12-4c64-b8ab-cd79212be45e_StandardRetail
16_a56a3b37-3a35-4bbb-a036-eee5f1898eee_VisioProRetail
16_980f9e3e-f5a8-41c8-8596-61404addf677_VisioStdRetail
16_cacaa1bf-da53-4c3b-9700-11738ef1c2a5_WordRetail
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do (

if %1==getmsiret if "%oVer%"=="%%A" (
for /f "tokens=*" %%x in ('findstr /i /c:"%%B" "%_oBranding%"') do set "prodId=%%x"
set prodId=!prodId:"/>=!
set prodId=!prodId:~-4!
if "%oVer%"=="14" (
REM Exception case for Visio because wrong primary product ID is mentioned in Branding.xml
echo %%C | find /i "Visio" %nul% && set prodId=0057
)
reg query "%2\Registration\{%%B}" /v ProductCode %nul2% | find /i "-!prodId!-" %nul% && (
reg query "%2\Common\InstalledPackages" %nul2% | find /i "-!prodId!-" %nul% && (
if defined _oIds (set _oIds=!_oIds! %%C) else (set _oIds=%%C)
)
)
)

)
)
exit /b

::========================================================================================================================================

::  1st column = Activation ID
::  2nd column = GVLK / Free Office products keys
::  3rd column = In case of Windows, its SKU ID. In case of Office, its Office version
::  4th column = Edition ID
::  5th column = In case of Windows, its Build Branch name in case same Edition ID is used in different OS versions with different key (For reference only)
::               In case of Office, its either a key type if its a free Office product or Retail product names that needs to be converted to the Edition ID mentioned in 4th column
::               In Office 2010, one highest VL edition from each primary product ID is selected, that's why Visio Prem key is mentioned but not for Visio Pro, Std
::  Separator  = "_"

:ksdata

set f=
for %%# in (
:: Windows 10/11
73111121-5638-40f6-bc11-f1d7b0d64300_NPPR9-FWDCX-D2C8J-H872K-2Y%f%T43___4_Enterprise
e272e3e2-732f-4c65-a8f0-484747d0d947_DPH2V-TTNVB-4X9Q3-TJR4H-KH%f%JW4__27_EnterpriseN
2de67392-b7a7-462a-b1ca-108dd189f588_W269N-WFGWX-YVC9B-4J6C9-T8%f%3GX__48_Professional
a80b5abf-76ad-428b-b05d-a47d2dffeebf_MH37W-N47XK-V7XM9-C7227-GC%f%QG9__49_ProfessionalN
7b9e1751-a8da-4f75-9560-5fadfe3d8e38_3KHY7-WNT83-DGQKR-F7HPR-84%f%4BM__98_CoreN
a9107544-f4a0-4053-a96a-1479abdef912_PVMJN-6DFY6-9CCP6-7BKTT-D3%f%WVR__99_CoreCountrySpecific
cd918a57-a41b-4c82-8dce-1a538e221a83_7HNRX-D7KGG-3K4RQ-4WPJ4-YT%f%DFH_100_CoreSingleLanguage
58e97c99-f377-4ef1-81d5-4ad5522b5fd8_TX9XD-98N7V-6WMQ6-BX7FG-H8%f%Q99_101_Core
e0c42288-980c-4788-a014-c080d2e1926e_NW6C2-QMPVW-D7KKK-3GKT6-VC%f%FB2_121_Education
3c102355-d027-42c6-ad23-2e7ef8a02585_2WH4N-8QGBV-H22JP-CT43Q-MD%f%WWJ_122_EducationN
32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7XTQ-FN8P6-TTKYV-9D4CC-J4%f%62D_125_EnterpriseS_RS5,VB,Ge
2d5a5a60-3040-48bf-beb0-fcd770c20ce0_DCPHK-NFMTC-H88MJ-PFHPY-QJ%f%4BJ_125_EnterpriseS_RS1
7b51a46c-0c04-4e8f-9af4-8496cca90d5e_WNMTR-4C88C-JK8YV-HQ7T2-76%f%DF9_125_EnterpriseS_TH1
7103a333-b8c8-49cc-93ce-d37c09687f92_92NFX-8DJQP-P6BBQ-THF9C-7C%f%G2H_126_EnterpriseSN_RS5,VB,Ge
9f776d83-7156-45b2-8a5c-359b9c9f22a3_QFFDN-GRT3P-VKWWX-X7T3R-8B%f%639_126_EnterpriseSN_RS1
87b838b7-41b6-4590-8318-5797951d8529_2F77B-TNFGY-69QQF-B8YKP-D6%f%9TJ_126_EnterpriseSN_TH1
82bbc092-bc50-4e16-8e18-b74fc486aec3_NRG8B-VKK3Q-CXVCJ-9G2XF-6Q%f%84J_161_ProfessionalWorkstation
4b1571d3-bafb-4b40-8087-a961be2caf65_9FNHH-K3HBT-3W4TD-6383H-6X%f%YWF_162_ProfessionalWorkstationN
3f1afc82-f8ac-4f6c-8005-1d233e606eee_6TP4R-GNPTD-KYYHQ-7B7DP-J4%f%47Y_164_ProfessionalEducation
5300b18c-2e33-4dc2-8291-47ffcec746dd_YVWGF-BXNMC-HTQYQ-CPQ99-66%f%QFC_165_ProfessionalEducationN
e0b2d383-d112-413f-8a80-97f373a5820c_YYVX9-NTFWV-6MDM3-9PT4T-4M%f%68B_171_EnterpriseG
e38454fb-41a4-4f59-a5dc-25080e354730_44RPN-FTY23-9VTTB-MP9BX-T8%f%4FV_172_EnterpriseGN
ec868e65-fadf-4759-b23e-93fe37f2cc29_CPWHC-NT2C7-VYW78-DHDB2-PG%f%3GK_175_ServerRdsh_RS5
e4db50ea-bda1-4566-b047-0ca50abc6f07_7NBT4-WGBQX-MP4H7-QXFF8-YP%f%3KX_175_ServerRdsh_RS3
0df4f814-3f57-4b8b-9a9d-fddadcd69fac_NBTWJ-3DR69-3C4V8-C26MC-GQ%f%9M6_183_CloudE
59eb965c-9150-42b7-a0ec-22151b9897c5_KBN8V-HFGQ4-MGXVD-347P6-PD%f%QGT_191_IoTEnterpriseS_VB,NI
d30136fc-cb4b-416e-a23d-87207abc44a9_6XN7V-PCBDC-BDBRH-8DQY7-G6%f%R44_202_CloudEditionN
ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69_37D7F-N49CB-WQR8W-TBJ73-FM%f%8RX_203_CloudEdition
:: Windows 2016/19/22/25 LTSC/SAC
7dc26449-db21-4e09-ba37-28f2958506a6_TVRH6-WHNXV-R9WG3-9XRFY-MY%f%832___7_ServerStandard_Ge
9bd77860-9b31-4b7b-96ad-2564017315bf_VDYBN-27WPP-V4HQT-9VMD4-VM%f%K7H___7_ServerStandard_FE
de32eafd-aaee-4662-9444-c1befb41bde2_N69G4-B89J2-4G8F4-WWYCC-J4%f%64C___7_ServerStandard_RS5
8c1c5410-9f39-4805-8c9d-63a07706358f_WC2BQ-8NRM3-FDDYY-2BFGV-KH%f%KQY___7_ServerStandard_RS1
c052f164-cdf6-409a-a0cb-853ba0f0f55a_D764K-2NDRG-47T6Q-P8T8W-YP%f%6DF___8_ServerDatacenter_Ge
ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03_WX4NM-KYWYW-QJJR4-XV3QB-6V%f%M33___8_ServerDatacenter_FE
34e1ae55-27f8-4950-8877-7a03be5fb181_WMDGN-G9PQG-XVVXX-R3X43-63%f%DFG___8_ServerDatacenter_RS5
21c56779-b449-4d20-adfc-eece0e1ad74b_CB7KF-BWN84-R7R2Y-793K2-8X%f%DDG___8_ServerDatacenter_RS1
034d3cbb-5d4b-4245-b3f8-f84571314078_WVDHN-86M7X-466P6-VHXV7-YY%f%726__50_ServerSolution_RS5
2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283_JCKRF-N37P4-C2D82-9YXRT-4M%f%63B__50_ServerSolution_RS1
7b4433f4-b1e7-4788-895a-c45378d38253_QN4C6-GBJD2-FB422-GHWJK-GJ%f%G2R_110_ServerCloudStorage
8de8eb62-bbe0-40ac-ac17-f75595071ea3_GRFBW-QNDC4-6QBHG-CCK3B-2P%f%R88_120_ServerARM64_RS5
43d9af6e-5e86-4be8-a797-d072a046896c_K9FYF-G6NCK-73M32-XMVPY-F9%f%DRR_120_ServerARM64_RS4
39e69c41-42b4-4a0a-abad-8e3c10a797cc_QFND9-D3Y9C-J3KKY-6RPVP-2D%f%PYV_145_ServerDatacenterACor_FE
90c362e5-0da1-4bfd-b53b-b87d309ade43_6NMRW-2C8FM-D24W7-TQWMY-CW%f%H2D_145_ServerDatacenterACor_RS5
e49c08e7-da82-42f8-bde2-b570fbcae76c_2HXDN-KRXHB-GPYC7-YCKFJ-7F%f%VDG_145_ServerDatacenterACor_RS3
f5e9429c-f50b-4b98-b15c-ef92eb5cff39_67KN8-4FYJW-2487Q-MQ2J7-4C%f%4RG_146_ServerStandardACor_FE
73e3957c-fc0c-400d-9184-5f7b6f2eb409_N2KJX-J94YW-TQVFB-DG9YT-72%f%4CC_146_ServerStandardACor_RS5
61c5ef22-f14f-4553-a824-c4b31e84b100_PTXN8-JFHJM-4WC78-MPCBR-9W%f%4KR_146_ServerStandardACor_RS3
45b5aff2-60a0-42f2-bc4b-ec6e5f7b527e_FCNV3-279Q9-BQB46-FTKXX-9H%f%PRH_168_ServerAzureCor_Ge
8c8f0ad3-9a43-4e05-b840-93b8d1475cbc_6N379-GGTMK-23C6M-XVVTC-CK%f%FRQ_168_ServerAzureCor_FE
a99cc1f0-7719-4306-9645-294102fbff95_FDNH6-VW9RW-BXPJ7-4XTYG-23%f%9TB_168_ServerAzureCor_RS5
3dbf341b-5f6c-4fa7-b936-699dce9e263f_VP34G-4NPPG-79JTQ-864T4-R3%f%MQX_168_ServerAzureCor_RS1
c2e946d1-cfa2-4523-8c87-30bc696ee584_XGN3F-F394H-FD2MY-PP6FD-8M%f%CRC_407_ServerTurbine_Ge
19b5e0fb-4431-46bc-bac1-2f1873e4ae73_NTBV8-9K7Q8-V27C6-M2BTV-KH%f%MXV_407_ServerTurbine_RS5
:: Windows 8.1
81671aaf-79d1-4eb1-b004-8cbbe173afea_MHF9N-XY6XB-WVXMC-BTDCT-MK%f%KG7___4_Enterprise
113e705c-fa49-48a4-beea-7dd879b46b14_TT4HM-HN7YT-62K67-RGRQJ-JF%f%FXW__27_EnterpriseN
c06b6981-d7fd-4a35-b7b4-054742b7af67_GCRJD-8NW9H-F2CDX-CCM8D-9D%f%6T9__48_Professional
7476d79f-8e48-49b4-ab63-4d0b813a16e4_HMCNV-VVBFX-7HMBH-CTY9B-B4%f%FXY__49_ProfessionalN
f7e88590-dfc7-4c78-bccb-6f3865b99d1a_VHXM3-NR6FT-RY6RT-CK882-KW%f%2CJ__86_EmbeddedIndustryA
0ab82d54-47f4-4acb-818c-cc5bf0ecb649_NMMPB-38DD4-R2823-62W8D-VX%f%KJB__89_EmbeddedIndustry
cd4e2d9f-5059-4a50-a92d-05d5bb1267c7_FNFKF-PWTVT-9RC8H-32HB2-JB%f%34X__91_EmbeddedIndustryE
ffee456a-cd87-4390-8e07-16146c672fd0_XYTND-K6QKT-K2MRH-66RTM-43%f%JKP__97_CoreARM
78558a64-dc19-43fe-a0d0-8075b2a370a3_7B9N3-D94CG-YTVHR-QBPX3-RJ%f%P64__98_CoreN
db78b74f-ef1c-4892-abfe-1e66b8231df6_NCTT7-2RGK8-WMHRF-RY7YQ-JT%f%XG3__99_CoreCountrySpecific
c72c6a1d-f252-4e7e-bdd1-3fca342acb35_BB6NG-PQ82V-VRDPW-8XVD2-V8%f%P66_100_CoreSingleLanguage
fe1c3238-432a-43a1-8e25-97e7d1ef10f3_M9Q9P-WNJJT-6PXPY-DWX8H-6X%f%WKK_101_Core
096ce63d-4fac-48a9-82a9-61ae9e800e5f_789NJ-TQK6T-6XTH8-J39CJ-J8%f%D3P_103_ProfessionalWMC
e9942b32-2e55-4197-b0bd-5ff58cba8860_3PY8R-QHNP9-W7XQD-G6DPH-3J%f%2C9_111_CoreConnected
c6ddecd6-2354-4c19-909b-306a3058484e_Q6HTR-N24GM-PMJFP-69CD8-2G%f%XKR_113_CoreConnectedN
b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c_KF37N-VDV38-GRRTV-XH8X6-6F%f%3BB_115_CoreConnectedSingleLanguage
ba998212-460a-44db-bfb5-71bf09d1c68b_R962J-37N87-9VVK2-WJ74P-XT%f%MHR_116_CoreConnectedCountrySpecific
e58d87b5-8126-4580-80fb-861b22f79296_MX3RK-9HNGX-K3QKC-6PJ3F-W8%f%D7B_112_ProfessionalStudent
cab491c7-a918-4f60-b502-dab75e334f40_TNFGH-2R6PB-8XM3K-QYHX2-J4%f%296_114_ProfessionalStudentN
:: Windows Server 2012 R2
b3ca044e-a358-4d68-9883-aaa2941aca99_D2N9P-3P6X9-2R39C-7RTCD-MD%f%VJX___7_ServerStandard
00091344-1ea4-4f37-b789-01750ba6988c_W3GGN-FT8W3-Y4M27-J84CP-Q3%f%VJ9___8_ServerDatacenter
21db6ba4-9a7b-4a14-9e29-64a60c59301d_KNC87-3J2TX-XB4WP-VCPJV-M4%f%FWM__50_ServerSolution
b743a2be-68d4-4dd3-af32-92425b7bb623_3NPTF-33KPT-GGBPR-YX76B-39%f%KDD_110_ServerCloudStorage
:: Windows 8
458e1bec-837a-45f6-b9d5-925ed5d299de_32JNW-9KQ84-P47T8-D8GGY-CW%f%CK7___4_Enterprise
e14997e7-800a-4cf7-ad10-de4b45b578db_JMNMF-RHW7P-DMY6X-RF3DR-X2%f%BQT__27_EnterpriseN
a98bcd6d-5343-4603-8afe-5908e4611112_NG4HW-VH26C-733KW-K6F98-J8%f%CK4__48_Professional
ebf245c1-29a8-4daf-9cb1-38dfc608a8c8_XCVCF-2NXM9-723PB-MHCB7-2R%f%YQQ__49_ProfessionalN
10018baf-ce21-4060-80bd-47fe74ed4dab_RYXVT-BNQG7-VD29F-DBMRY-HT%f%73M__89_EmbeddedIndustry
18db1848-12e0-4167-b9d7-da7fcda507db_NKB3R-R2F8T-3XCDP-7Q2KW-XW%f%YQ2__91_EmbeddedIndustryE
af35d7b7-5035-4b63-8972-f0b747b9f4dc_DXHJF-N9KQX-MFPVR-GHGQK-Y7%f%RKV__97_CoreARM
197390a0-65f6-4a95-bdc4-55d58a3b0253_8N2M2-HWPGY-7PGT9-HGDD8-GV%f%GGY__98_CoreN
9d5584a2-2d85-419a-982c-a00888bb9ddf_4K36P-JN4VD-GDC6V-KDT89-DY%f%FKP__99_CoreCountrySpecific
8860fcd4-a77b-4a20-9045-a150ff11d609_2WN2H-YGCQR-KFX6K-CD6TF-84%f%YXQ_100_CoreSingleLanguage
c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60_BN3D2-R7TKB-3YPBD-8DRP2-27%f%GG4_101_Core
a00018a3-f20f-4632-bf7c-8daa5351c914_GNBB8-YVD74-QJHX6-27H4K-8Q%f%HDG_103_ProfessionalWMC
:: Windows Server 2012
f0f5ec41-0d55-4732-af02-440a44a3cf0f_XC9B7-NBPP2-83J2H-RHMBY-92%f%BT4___7_ServerStandard
d3643d60-0c42-412d-a7d6-52e6635327f6_48HP8-DN98B-MYWDG-T2DCC-8W%f%83P___8_ServerDatacenter
8f365ba6-c1b9-4223-98fc-282a0756a3ed_HTDQM-NBMMG-KGYDT-2DTKT-J2%f%MPV__50_ServerSolution
7d5486c7-e120-4771-b7f1-7b56c6d3170c_HM7DN-YVMH3-46JC3-XYTG7-CY%f%QJJ__76_ServerMultiPointStandard
95fd1c83-7df5-494a-be8b-1300e1c9d1cd_XNH6W-2V9GX-RGJ4K-Y8X6F-QG%f%J2G__77_ServerMultiPointPremium
:: Windows 7
ae2ee509-1b34-41c0-acb7-6d4650168915_33PXH-7Y6KF-2VJC9-XBBR8-HV%f%THH___4_Enterprise
1cb6d605-11b3-4e14-bb30-da91c8e3983a_YDRBP-3D83W-TY26F-D46B2-XC%f%KRJ__27_EnterpriseN
b92e9980-b9d5-4821-9c94-140f632f6312_FJ82H-XT6CR-J8D7P-XQJJ2-GP%f%DD4__48_Professional
54a09a0d-d57b-4c10-8b69-a842d6590ad5_MRPKT-YTG23-K7D7T-X2JMM-QY%f%7MG__49_ProfessionalN
db537896-376f-48ae-a492-53d0547773d0_YBYF6-BHCR3-JPKRB-CDW7B-F9%f%BK4__65_Embedded_POSReady
aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395_73KQT-CD9G6-K7TQG-66MRP-CQ%f%22C__65_Embedded_ThinPC
5a041529-fef8-4d07-b06f-b59b573b32d2_W82YF-2Q76Y-63HXB-FGJG9-GF%f%7QX__69_ProfessionalE
46bbed08-9c7b-48fc-a614-95250573f4ea_C29WB-22CC8-VJ326-GHFJW-H9%f%DH4__70_EnterpriseE
:: Windows Server 2008 R2
68531fb9-5511-4989-97be-d11a0f55633f_YC6KT-GKW9T-YTKYR-T4X34-R7%f%VHC___7_ServerStandard
7482e61b-c589-4b7f-8ecc-46d455ac3b87_74YFP-3QFB3-KQT8W-PMXWJ-7M%f%648___8_ServerDatacenter
620e2b3d-09e7-42fd-802a-17a13652fe7a_489J6-VHDMP-X63PK-3K798-CP%f%X3Y__10_ServerEnterprise
7482e61b-c589-4b7f-8ecc-46d455ac3b87_74YFP-3QFB3-KQT8W-PMXWJ-7M%f%648__12_ServerDatacenterCore
68531fb9-5511-4989-97be-d11a0f55633f_YC6KT-GKW9T-YTKYR-T4X34-R7%f%VHC__13_ServerStandardCore
620e2b3d-09e7-42fd-802a-17a13652fe7a_489J6-VHDMP-X63PK-3K798-CP%f%X3Y__14_ServerEnterpriseCore
8a26851c-1c7e-48d3-a687-fbca9b9ac16b_GT63C-RJFQ3-4GMB6-BRFB9-CB%f%83V__15_ServerEnterpriseIA64
a78b8bd9-8017-4df5-b86a-09f756affa7c_6TPJF-RBVHG-WBW2R-86QPH-6R%f%TM4__17_ServerWeb
cda18cf3-c196-46ad-b289-60c072869994_TT8MH-CG224-D3D7Q-498W2-9Q%f%CTX__18_ServerHPC
a78b8bd9-8017-4df5-b86a-09f756affa7c_6TPJF-RBVHG-WBW2R-86QPH-6R%f%TM4__29_ServerWebCore
f772515c-0e87-48d5-a676-e6962c3e1195_736RG-XDKJK-V34PF-BHK87-J6%f%X3K__56_ServerEmbeddedSolution
::========================================================================================================================================
:: Office 2010
8ce7e872-188c-4b98-9d90-f8f90b7aad02_V7Y44-9T38C-R2VJK-666HK-T7%f%DDX__14_AccessVL
cee5d470-6e3b-4fcc-8c2b-d17428568a9f_H62QG-HXVKF-PP4HP-66KMR-CW%f%9BM__14_ExcelVL
8947d0b8-c33b-43e1-8c56-9b674c052832_QYYW6-QP4CB-MBV6G-HYMCJ-4T%f%3J4__14_GrooveVL
ca6b6639-4ad6-40ae-a575-14dee07f6430_K96W8-67RPQ-62T9Y-J8FQJ-BT%f%37T__14_InfoPathVL
09ed9640-f020-400a-acd8-d7d867dfd9c2_YBJTT-JG6MD-V9Q7P-DBKXJ-38%f%W9R__14_MondoVL
ab586f5c-5256-4632-962f-fefd8b49e6f4_Q4Y4M-RHWJM-PY37F-MTKWH-D3%f%XHX__14_OneNoteVL
ecb7c192-73ab-4ded-acf4-2399b095d0cc_7YDC2-CWM8M-RRTJC-8MDVC-X3%f%DWQ__14_OutlookVL
45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a_RC8FX-88JRY-3PF7C-X8P67-P4%f%VTT__14_PowerPointVL
df133ff7-bf14-4f95-afe3-7b48e7e331ef_YGX6F-PGV49-PGW3J-9BTGG-VH%f%KC6__14_ProjectProVL
5dc7bf61-5ec9-4996-9ccb-df806a2d0efe_4HP3K-88W3F-W2K3D-6677X-F9%f%PGB__14_ProjectStdVL
6f327760-8c5c-417c-9b61-836a98287e0c_VYBBJ-TRJPB-QFQRF-QFT4D-H3%f%GVB__14_ProPlusVL
b50c4f75-599b-43e8-8dcd-1081a7967241_BFK7F-9MYHM-V68C7-DRQ66-83%f%YTP__14_PublisherVL
ea509e87-07a1-4a45-9edc-eba5a39f36af_D6QFG-VBYP2-XQHM7-J97RH-VV%f%RCK__14_SmallBusBasicsVL
9da2a678-fb6b-4e67-ab84-60dd6a9c819a_V7QKV-4XVVR-XYV4D-F7DFM-8R%f%6BM__14_StandardVL
92236105-bb67-494f-94c7-7f7a607929bd_D9DWC-HPYVV-JGF4P-BTWQB-WX%f%8BJ__14_VisioSIVL
2d0882e7-a4e7-423b-8ccc-70d91e0158b1_HVHB3-C6FV7-KQX9W-YQG79-CR%f%Y7T__14_WordVL
:: Office 2013
6ee7622c-18d8-4005-9fb7-92db644a279b_NG2JY-H4JBT-HQXYP-78QH9-4J%f%M2D__15_AccessVolume_-AccessRetail-
259de5be-492b-44b3-9d78-9645f848f7b0_X3XNB-HJB7K-66THH-8DWQ3-XH%f%GJP__15_AccessRuntimeRetail_[Bypass]
f7461d52-7c2b-43b2-8744-ea958e0bd09a_VGPNG-Y7HQW-9RHP7-TKPV3-BG%f%7GB__15_ExcelVolume_-ExcelRetail-
fb4875ec-0c6b-450f-b82b-ab57d8d1677f_H7R7V-WPNXQ-WCYYC-76BGV-VT%f%7GH__15_GrooveVolume_-GrooveRetail-
a30b8040-d68a-423f-b0b5-9ce292ea5a8f_DKT8B-N7VXH-D963P-Q4PHY-F8%f%894__15_InfoPathVolume_-InfoPathRetail-
9103f3ce-1084-447a-827e-d6097f68c895_6MDN4-WF3FV-4WH3Q-W699V-RG%f%CMY__15_LyncAcademicRetail_[PrepidBypass]
ff693bf4-0276-4ddb-bb42-74ef1a0c9f4d_N42BF-CBY9F-W2C7R-X397X-DY%f%FQW__15_LyncEntryRetail_[PrepidBypass]
1b9f11e3-c85c-4e1b-bb29-879ad2c909e3_2MG3G-3BNTT-3MFW9-KDQW3-TC%f%K7R__15_LyncVolume_-LyncRetail-
1dc00701-03af-4680-b2af-007ffc758a1f_CWH2Y-NPYJW-3C7HD-BJQWB-G2%f%8JJ__15_MondoRetail
dc981c6b-fc8e-420f-aa43-f8f33e5c0923_42QTK-RN8M7-J3C4G-BBGYM-88%f%CYV__15_MondoVolume_-O365BusinessRetail-O365HomePremRetail-O365ProPlusRetail-O365SmallBusPremRetail-
3391e125-f6e4-4b1e-899c-a25e6092d40d_4TGWV-6N9P6-G2H8Y-2HWKB-B4%f%FF4__15_OneNoteFreeRetail_[Bypass]
efe1f3e6-aea2-4144-a208-32aa872b6545_TGN6P-8MMBC-37P2F-XHXXK-P3%f%4VW__15_OneNoteVolume_-OneNoteRetail-
771c3afa-50c5-443f-b151-ff2546d863a0_QPN8Q-BJBTJ-334K3-93TGY-2P%f%MBT__15_OutlookVolume_-OutlookRetail-
8c762649-97d1-4953-ad27-b7e2c25b972e_4NT99-8RJFH-Q2VDH-KYG2C-4R%f%D4F__15_PowerPointVolume_-PowerPointRetail-
4a5d124a-e620-44ba-b6ff-658961b33b9a_FN8TT-7WMH6-2D4X9-M337T-23%f%42K__15_ProjectProVolume_-ProjectProRetail-
427a28d1-d17c-4abf-b717-32c780ba6f07_6NTH3-CW976-3G3Y2-JK3TX-8Q%f%HTT__15_ProjectStdVolume_-ProjectStdRetail-
b322da9c-a2e2-4058-9e4e-f59a6970bd69_YC7DK-G2NP3-2QQC3-J6H88-GV%f%GXT__15_ProPlusVolume_-ProPlusRetail-ProfessionalPipcRetail-ProfessionalRetail-
00c79ff1-6850-443d-bf61-71cde0de305f_PN2WF-29XG2-T9HJ7-JQPJR-FC%f%XK4__15_PublisherVolume_-PublisherRetail-
ba3e3833-6a7e-445a-89d0-7802a9a68588_3NY6J-WHT3F-47BDV-JHF36-23%f%43W__15_SPDRetail_[PrepidBypass]
b13afb38-cd79-4ae5-9f7f-eed058d750ca_KBKQT-2NMXY-JJWGP-M62JB-92%f%CD4__15_StandardVolume_-StandardRetail-HomeBusinessPipcRetail-HomeBusinessRetail-HomeStudentARMRetail-HomeStudentPlusARMRetail-HomeStudentRetail-PersonalPipcRetail-PersonalRetail-
e13ac10e-75d0-4aff-a0cd-764982cf541c_C2FG9-N6J68-H8BTJ-BW3QX-RM%f%3B3__15_VisioProVolume_-VisioProRetail-
ac4efaf0-f81f-4f61-bdf7-ea32b02ab117_J484Y-4NKBF-W2HMG-DBMJC-PG%f%WR7__15_VisioStdVolume_-VisioStdRetail-
d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3_6Q7VD-NX8JD-WJ2VH-88V73-4G%f%BJ7__15_WordVolume_-WordRetail-
:: Office 2016
9d9faf9e-d345-4b49-afce-68cb0a539c7c_RNB7V-P48F4-3FYY6-2P3R3-63%f%BQV__16_AccessRuntimeRetail_[PrepidBypass]
67c0fc0c-deba-401b-bf8b-9c8ad8395804_GNH9Y-D2J4T-FJHGG-QRVH7-QP%f%FDW__16_AccessVolume_-AccessRetail-
c3e65d36-141f-4d2f-a303-a842ee756a29_9C2PK-NWTVB-JMPW8-BFT28-7F%f%TBF__16_ExcelVolume_-ExcelRetail-
e914ea6e-a5fa-4439-a394-a9bb3293ca09_DMTCJ-KNRKX-26982-JYCKT-P7%f%KB6__16_MondoRetail
9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce_HFTND-W9MK4-8B7MJ-B6C4G-XQ%f%BR2__16_MondoVolume_-O365BusinessRetail-O365EduCloudRetail-O365HomePremRetail-O365ProPlusRetail-O365SmallBusPremRetail-
436366de-5579-4f24-96db-3893e4400030_XYNTG-R96FY-369HX-YFPHY-F9%f%CPM__16_OneNoteFreeRetail_[Bypass]
d8cace59-33d2-4ac7-9b1b-9b72339c51c8_DR92N-9HTF2-97XKM-XW2WJ-XW%f%3J6__16_OneNoteVolume_-OneNoteRetail-OneNote2021Retail-
ec9d9265-9d1e-4ed0-838a-cdc20f2551a1_R69KK-NTPKF-7M3Q4-QYBHW-6M%f%T9B__16_OutlookVolume_-OutlookRetail-
d70b1bba-b893-4544-96e2-b7a318091c33_J7MQP-HNJ4Y-WJ7YM-PFYGF-BY%f%6C6__16_PowerPointVolume_-PowerPointRetail-
4f414197-0fc2-4c01-b68a-86cbb9ac254c_YG9NW-3K39V-2T3HJ-93F3Q-G8%f%3KT__16_ProjectProVolume_-ProjectProRetail-
829b8110-0e6f-4349-bca4-42803577788d_WGT24-HCNMF-FQ7XH-6M8K7-DR%f%TW9__16_ProjectProXVolume
da7ddabc-3fbe-4447-9e01-6ab7440b4cd4_GNFHQ-F6YQM-KQDGJ-327XX-KQ%f%BVC__16_ProjectStdVolume_-ProjectStdRetail-
cbbaca45-556a-4416-ad03-bda598eaa7c8_D8NRQ-JTYM3-7J2DX-646CT-68%f%36M__16_ProjectStdXVolume
d450596f-894d-49e0-966a-fd39ed4c4c64_XQNVK-8JYDB-WJ9W3-YJ8YR-WF%f%G99__16_ProPlusVolume_-ProPlusRetail-ProfessionalPipcRetail-ProfessionalRetail-
041a06cb-c5b8-4772-809f-416d03d16654_F47MM-N3XJP-TQXJ9-BP99D-8K%f%837__16_PublisherVolume_-PublisherRetail-
9103f3ce-1084-447a-827e-d6097f68c895_6MDN4-WF3FV-4WH3Q-W699V-RG%f%CMY__16_SkypeServiceBypassRetail_[PrepidBypass]
971cd368-f2e1-49c1-aedd-330909ce18b6_4N4D8-3J7Y3-YYW7C-73HD2-V8%f%RHY__16_SkypeforBusinessEntryRetail_[PrepidBypass]
83e04ee1-fa8d-436d-8994-d31a862cab77_869NQ-FJ69K-466HW-QYCP2-DD%f%BV6__16_SkypeforBusinessVolume_-SkypeforBusinessRetail-
dedfa23d-6ed1-45a6-85dc-63cae0546de6_JNRGM-WHDWX-FJJG3-K47QV-DR%f%TFM__16_StandardVolume_-StandardRetail-HomeBusinessPipcRetail-HomeBusinessRetail-HomeStudentARMRetail-HomeStudentPlusARMRetail-HomeStudentRetail-HomeStudentVNextRetail-PersonalPipcRetail-PersonalRetail-
6bf301c1-b94a-43e9-ba31-d494598c47fb_PD3PC-RHNGV-FXJ29-8JK7D-RJ%f%RJK__16_VisioProVolume_-VisioProRetail-
b234abe3-0857-4f9c-b05a-4dc314f85557_69WXN-MBYV6-22PQG-3WGHK-RM%f%6XC__16_VisioProXVolume
aa2a7821-1827-4c2c-8f1d-4513a34dda97_7WHWN-4T7MP-G96JF-G33KR-W8%f%GF4__16_VisioStdVolume_-VisioStdRetail-
361fe620-64f4-41b5-ba77-84f8e079b1f7_NY48V-PPYYH-3F4PX-XJRKJ-W4%f%423__16_VisioStdXVolume
bb11badf-d8aa-470e-9311-20eaf80fe5cc_WXY84-JN2Q9-RBCCQ-3Q3J3-3P%f%FJ6__16_WordVolume_-WordRetail-
:: Office 2019
22e6b96c-1011-4cd5-8b35-3c8fb6366b86_FGQNJ-JWJCG-7Q8MG-RMRGJ-9T%f%QVF__16_AccessRuntime2019Retail_[PrepidBypass]
9e9bceeb-e736-4f26-88de-763f87dcc485_9N9PT-27V4Y-VJ2PD-YXFMF-YT%f%FQT__16_Access2019Volume_-Access2019Retail-
237854e9-79fc-4497-a0c1-a70969691c6b_TMJWT-YYNMB-3BKTF-644FC-RV%f%XBD__16_Excel2019Volume_-Excel2019Retail-
c8f8a301-19f5-4132-96ce-2de9d4adbd33_7HD7K-N4PVK-BHBCQ-YWQRW-XW%f%4VK__16_Outlook2019Volume_-Outlook2019Retail-
3131fd61-5e4f-4308-8d6d-62be1987c92c_RRNCX-C64HY-W2MM7-MCH9G-TJ%f%HMQ__16_PowerPoint2019Volume_-PowerPoint2019Retail-
2ca2bf3f-949e-446a-82c7-e25a15ec78c4_B4NPR-3FKK7-T2MBV-FRQ4W-PK%f%D2B__16_ProjectPro2019Volume_-ProjectPro2019Retail-
1777f0e3-7392-4198-97ea-8ae4de6f6381_C4F7P-NCP8C-6CQPT-MQHV9-JX%f%D2M__16_ProjectStd2019Volume_-ProjectStd2019Retail-
85dd8b5f-eaa4-4af3-a628-cce9e77c9a03_NMMKJ-6RK4F-KMJVX-8D9MJ-6M%f%WKP__16_ProPlus2019Volume_-ProPlus2019Retail-Professional2019Retail-
9d3e4cca-e172-46f1-a2f4-1d2107051444_G2KWX-3NW6P-PY93R-JXK2T-C9%f%Y9V__16_Publisher2019Volume_-Publisher2019Retail-
734c6c6e-b0ba-4298-a891-671772b2bd1b_NCJ33-JHBBY-HTK98-MYCV8-HM%f%KHJ__16_SkypeforBusiness2019Volume_-SkypeforBusiness2019Retail-
f88cfdec-94ce-4463-a969-037be92bc0e7_N9722-BV9H6-WTJTT-FPB93-97%f%8MK__16_SkypeforBusinessEntry2019Retail_[PrepidBypass]
6912a74b-a5fb-401a-bfdb-2e3ab46f4b02_6NWWJ-YQWMR-QKGCB-6TMB3-9D%f%9HK__16_Standard2019Volume_-Standard2019Retail-HomeBusiness2019Retail-HomeStudentARM2019Retail-HomeStudentPlusARM2019Retail-HomeStudent2019Retail-Personal2019Retail-
5b5cf08f-b81a-431d-b080-3450d8620565_9BGNQ-K37YR-RQHF2-38RQ3-7V%f%CBB__16_VisioPro2019Volume_-VisioPro2019Retail-
e06d7df3-aad0-419d-8dfb-0ac37e2bdf39_7TQNQ-K3YQQ-3PFH7-CCPPM-X4%f%VQ2__16_VisioStd2019Volume_-VisioStd2019Retail-
059834fe-a8ea-4bff-b67b-4d006b5447d3_PBX3G-NWMT6-Q7XBW-PYJGG-WX%f%D33__16_Word2019Volume_-Word2019Retail-
:: Office 2021
:: OneNote2021Volume KMS license is not available
844c36cb-851c-49e7-9079-12e62a049e2a_MNX9D-PB834-VCGY2-K2RW2-2D%f%P3D__16_AccessRuntime2021Retail_[Bypass]
1fe429d8-3fa7-4a39-b6f0-03dded42fe14_WM8YG-YNGDD-4JHDC-PG3F4-FC%f%4T4__16_Access2021Volume_-Access2021Retail-
ea71effc-69f1-4925-9991-2f5e319bbc24_NWG3X-87C9K-TC7YY-BC2G7-G6%f%RVC__16_Excel2021Volume_-Excel2021Retail-
a5799e4c-f83c-4c6e-9516-dfe9b696150b_C9FM6-3N72F-HFJXB-TM3V9-T8%f%6R9__16_Outlook2021Volume_-Outlook2021Retail-
778ccb9a-2f6a-44e5-853c-eb22b7609643_CNM3W-V94GB-QJQHH-BDQ3J-33%f%Y8H__16_OneNoteFree2021Retail_[Bypass]
6e166cc3-495d-438a-89e7-d7c9e6fd4dea_TY7XF-NFRBR-KJ44C-G83KF-GX%f%27K__16_PowerPoint2021Volume_-PowerPoint2021Retail-
76881159-155c-43e0-9db7-2d70a9a3a4ca_FTNWT-C6WBT-8HMGF-K9PRX-QV%f%9H8__16_ProjectPro2021Volume_-ProjectPro2021Retail-
6dd72704-f752-4b71-94c7-11cec6bfc355_J2JDC-NJCYY-9RGQ4-YXWMH-T3%f%D4T__16_ProjectStd2021Volume_-ProjectStd2021Retail-
fbdb3e18-a8ef-4fb3-9183-dffd60bd0984_FXYTK-NJJ8C-GB6DW-3DYQT-6F%f%7TH__16_ProPlus2021Volume_-ProPlus2021Retail-Professional2021Retail-
aa66521f-2370-4ad8-a2bb-c095e3e4338f_2MW9D-N4BXM-9VBPG-Q7W6M-KF%f%BGQ__16_Publisher2021Volume_-Publisher2021Retail-
1f32a9af-1274-48bd-ba1e-1ab7508a23e8_HWCXN-K3WBT-WJBKY-R8BD9-XK%f%29P__16_SkypeforBusiness2021Volume_-SkypeforBusiness2021Retail-
080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3_KDX7X-BNVR8-TXXGX-4Q7Y8-78%f%VT3__16_Standard2021Volume_-Standard2021Retail-HomeBusiness2021Retail-HomeStudent2021Retail-Personal2021Retail-
fb61ac9a-1688-45d2-8f6b-0674dbffa33c_KNH8D-FGHT4-T8RK3-CTDYJ-K2%f%HT4__16_VisioPro2021Volume_-VisioPro2021Retail-
72fce797-1884-48dd-a860-b2f6a5efd3ca_MJVNY-BYWPY-CWV6J-2RKRT-4M%f%8QG__16_VisioStd2021Volume_-VisioStd2021Retail-
abe28aea-625a-43b1-8e30-225eb8fbd9e5_TN8H9-M34D3-Y64V9-TR72V-X7%f%9KV__16_Word2021Volume_-Word2021Retail-
:: Office 2024
fceda083-1203-402a-8ec4-3d7ed9f3648c_2TDPW-NDQ7G-FMG99-DXQ7M-TX%f%3T2__16_ProPlus2024Volume-Preview
aaea0dc8-78e1-4343-9f25-b69b83dd1bce_D9GTG-NP7DV-T6JP3-B6B62-JB%f%89R__16_ProjectPro2024Volume-Preview
4ab4d849-aabc-43fb-87ee-3aed02518891_YW66X-NH62M-G6YFP-B7KCT-WX%f%GKQ__16_VisioPro2024Volume-Preview
72e9faa7-ead1-4f3d-9f6e-3abc090a81d7_82FTR-NCHR7-W3944-MGRHM-JM%f%CWD__16_Access2024Volume_-Access2024Retail-
cbbba2c3-0ff5-4558-846a-043ef9d78559_F4DYN-89BP2-WQTWJ-GR8YC-CK%f%GJG__16_Excel2024Volume_-Excel2024Retail-
bef3152a-8a04-40f2-a065-340c3f23516d_D2F8D-N3Q3B-J28PV-X27HD-RJ%f%WB9__16_Outlook2024Volume_-Outlook2024Retail-
b63626a4-5f05-4ced-9639-31ba730a127e_CW94N-K6GJH-9CTXY-MG2VC-FY%f%CWP__16_PowerPoint2024Volume_-PowerPoint2024Retail-
f510af75-8ab7-4426-a236-1bfb95c34ff8_FQQ23-N4YCY-73HQ3-FM9WC-76%f%HF4__16_ProjectPro2024Volume_-ProjectPro2024Retail-
9f144f27-2ac5-40b9-899d-898c2b8b4f81_PD3TT-NTHQQ-VC7CY-MFXK3-G8%f%7F8__16_ProjectStd2024Volume_-ProjectStd2024Retail-
8d368fc1-9470-4be2-8d66-90e836cbb051_XJ2XN-FW8RK-P4HMP-DKDBV-GC%f%VGB__16_ProPlus2024Volume_-ProPlus2024Retail-
0002290a-2091-4324-9e53-3cfe28884cde_4NKHF-9HBQF-Q3B6C-7YV34-F6%f%4P3__16_SkypeforBusiness2024Volume
bbac904f-6a7e-418a-bb4b-24c85da06187_V28N4-JG22K-W66P8-VTMGK-H6%f%HGR__16_Standard2024Volume_-Home2024Retail-HomeBusiness2024Retail-
fa187091-8246-47b1-964f-80a0b1e5d69a_B7TN8-FJ8V3-7QYCP-HQPMV-YY%f%89G__16_VisioPro2024Volume_-VisioPro2024Retail-
923fa470-aa71-4b8b-b35c-36b79bf9f44b_JMMVY-XFNQC-KK4HK-9H7R3-WQ%f%QTV__16_VisioStd2024Volume_-VisioStd2024Retail-
d0eded01-0881-4b37-9738-190400095098_MQ84N-7VYDM-FXV7C-6K7CC-VF%f%W9J__16_Word2024Volume_-Word2024Retail-
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do (

if %1==winkey if %osSKU%==%%C if not defined key (
echo "!allapps!" | find /i "%%A" %nul1% && set key=%%B
)

if %1==chkprod if "%oVer%"=="%%C" if not defined foundprod (
echo "%%D" | findstr /I "\<%2.*" %nul% && set foundprod=1
)

if %1==getinfo if not defined key if "%oVer%"=="%%C" (
if /i "%2"=="%%D" (
set key=%%B
set _actid=%%A
set _allactid=!_allactid! %%A
) else if not defined _oBranding if %_NoEditionChange%==0 (
echo: %%E | find /i "-%2-" %nul% && (
set key=%%B
set _altoffid=%%D
set _actid=%%A
set _allactid=!_allactid! %%A
)
)
)

if %1==getmsiprod if "%oVer%"=="%%C" (
for /f "tokens=*" %%x in ('findstr /i /c:"%%A" "%_oBranding%"') do set "prodId=%%x"
set prodId=!prodId:"/>=!
set prodId=!prodId:~-4!
if "%oVer%"=="14" (
REM Exception case for Visio because wrong primary product ID is mentioned in Branding.xml
echo %%D | find /i "Visio" %nul% && set prodId=0057
)
reg query "%2\Registration\{%%A}" /v ProductCode %nul2% | find /i "-!prodId!-" %nul% && (
reg query "%2\Common\InstalledPackages" %nul2% | find /i "-!prodId!-" %nul% && (
if defined _oIds (set _oIds=!_oIds! %%D) else (set _oIds=%%D)
)
)
)

)
)
exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:check_actstatus

cls
if not defined terminal (
mode 100, 36
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=35;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)

%psc% "$f=[IO.File]::ReadAllText('!_batp!') -split ':sppmgr\:.*';iex ($f[1])"
goto dk_done

:sppmgr:
param (
    [Parameter()]
    [switch]
    $All,
    [Parameter()]
    [switch]
    $Dlv,
    [Parameter()]
    [switch]
    $IID,
    [Parameter()]
    [switch]
    $Pass
)

function CONOUT($strObj)
{
	Out-Host -Input $strObj
}

function ExitScript($ExitCode = 0)
{
	Exit $ExitCode
}

if (-Not $PSVersionTable) {
	"==== ERROR ====`r`n"
	"Windows PowerShell 1.0 is not supported by this script."
	ExitScript 1
}

if ($ExecutionContext.SessionState.LanguageMode.value__ -NE 0) {
	"==== ERROR ====`r`n"
	"Windows PowerShell is not running in Full Language Mode."
	ExitScript 1
}

$winbuild = 1
try {
	$winbuild = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$env:SystemRoot\System32\kernel32.dll").FileBuildPart
} catch {
	$winbuild = [int]([wmi]'Win32_OperatingSystem=@').BuildNumber
}

if ($winbuild -EQ 1) {
	"==== ERROR ====`r`n"
	"Could not detect Windows build."
	ExitScript 1
}

if ($winbuild -LT 2600) {
	"==== ERROR ====`r`n"
	"This build of Windows is not supported by this script."
	ExitScript 1
}

if ($All.IsPresent)
{
	$isAll = {CONOUT "`r"}
	$noAll = {$null}
}
else
{
	$isAll = {$null}
	$noAll = {CONOUT "`r"}
}
$Dlv = $Dlv.IsPresent
$IID = $IID.IsPresent -Or $Dlv.IsPresent

$NT6 = $winbuild -GE 6000
$NT7 = $winbuild -GE 7600
$NT9 = $winbuild -GE 9600

$Admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$line2 = "============================================================"
$line3 = "____________________________________________________________"

function echoWindows
{
	CONOUT "$line2"
	CONOUT "===                   Windows Status                     ==="
	CONOUT "$line2"
	& $noAll
}

function echoOffice
{
	if ($doMSG -EQ 0) {
		return
	}

	& $isAll
	CONOUT "$line2"
	CONOUT "===                   Office Status                      ==="
	CONOUT "$line2"
	& $noAll

	$script:doMSG = 0
}

function strGetRegistry($strKey, $strName)
{
	try {
		return [Microsoft.Win32.Registry]::GetValue($strKey, $strName, $null)
	} catch {
		return $null
	}
}

function CheckOhook
{
	$ohook = 0
	$paths = "${env:ProgramFiles}", "${env:ProgramW6432}", "${env:ProgramFiles(x86)}"

	15, 16 | foreach `
	{
		$A = $_; $paths | foreach `
		{
			if (Test-Path "$($_)$('\Microsoft Office\Office')$($A)$('\sppc*dll')") {$ohook = 1}
		}
	}

	"System", "SystemX86" | foreach `
	{
		$A = $_; "Office 15", "Office" | foreach `
		{
			$B = $_; $paths | foreach `
			{
				if (Test-Path "$($_)$('\Microsoft ')$($B)$('\root\vfs\')$($A)$('\sppc*dll')") {$ohook = 1}
			}
		}
	}

	if ($ohook -EQ 0) {
		return
	}

	& $isAll
	CONOUT "$line2"
	CONOUT "===                Office Ohook Status                   ==="
	CONOUT "$line2"
	$host.UI.WriteLine('Yellow', 'Black', "`r`nOhook for permanent Office activation is installed.`r`nYou can ignore the below mentioned Office activation status.")
	& $noAll
}

#region WMI
function DetectID($strSLP, $strAppId)
{
	$ppk = (" AND PartialProductKey <> NULL)", ")")[$All.IsPresent]
	$fltr = "SELECT ID FROM $strSLP WHERE (ApplicationID='$strAppId'"
	$clause = $fltr + $ppk
	$sWmi = [wmisearcher]$clause
	$sWmi.Options.Rewindable = $false
	return ($sWmi.Get().Count -GT 0)
}

function GetID($strSLP, $strAppId)
{
	$NT5 = ($strSLP -EQ $wslp -And $winbuild -LT 6001)
	$IDs = [Collections.ArrayList]@()
	$isAdd = (" AND LicenseDependsOn <> NULL)", ")")[$NT5]
	$noAdd = " AND LicenseDependsOn IS NULL)"
	$query = "SELECT ID FROM $strSLP WHERE (ApplicationID='$strAppId' AND PartialProductKey"

	if ($All.IsPresent) {
		$fltr = $query + " IS NULL"
		$clause = $fltr + $isAdd
		$sWmi = [wmisearcher]$clause
		$sWmi.Options.Rewindable = $false
		try {$sWmi.Get() | select -Expand Properties -EA 0 | foreach {$IDs += $_.Value}} catch {}
		if (-Not $NT5) {
		$clause = $fltr + $noAdd
		$sWmi = [wmisearcher]$clause
		$sWmi.Options.Rewindable = $false
		try {$sWmi.Get() | select -Expand Properties -EA 0 | foreach {$IDs += $_.Value}} catch {}
		}
	}

	$fltr = $query + " <> NULL"
	$clause = $fltr + $isAdd
	$sWmi = [wmisearcher]$clause
	$sWmi.Options.Rewindable = $false
	try {$sWmi.Get() | select -Expand Properties -EA 0 | foreach {$IDs += $_.Value}} catch {}
	if (-Not $NT5) {
	$clause = $fltr + $noAdd
	$sWmi = [wmisearcher]$clause
	$sWmi.Options.Rewindable = $false
	try {$sWmi.Get() | select -Expand Properties -EA 0 | foreach {$IDs += $_.Value}} catch {}
	}

	return $IDs
}

function DetectSubscription {
	if ($null -EQ $objSvc.SubscriptionType -Or $objSvc.SubscriptionType -EQ 120) {
		return
	}

	if ($objSvc.SubscriptionType -EQ 1) {
		$SubMsgType = "Device based"
	} else {
		$SubMsgType = "User based"
	}

	if ($objSvc.SubscriptionStatus -EQ 120) {
		$SubMsgStatus = "Expired"
	} elseif ($objSvc.SubscriptionStatus -EQ 100) {
		$SubMsgStatus = "Disabled"
	} elseif ($objSvc.SubscriptionStatus -EQ 1) {
		$SubMsgStatus = "Active"
	} else {
		$SubMsgStatus = "Not active"
	}

	$SubMsgExpiry = "Unknown"
	if ($objSvc.SubscriptionExpiry) {
		if ($objSvc.SubscriptionExpiry.Contains("unspecified") -EQ $false) {$SubMsgExpiry = $objSvc.SubscriptionExpiry}
	}

	$SubMsgEdition = "Unknown"
	if ($objSvc.SubscriptionEdition) {
		if ($objSvc.SubscriptionEdition.Contains("UNKNOWN") -EQ $false) {$SubMsgEdition = $objSvc.SubscriptionEdition}
	}

	CONOUT "`nSubscription information:"
	CONOUT "    Edition: $SubMsgEdition"
	CONOUT "    Type   : $SubMsgType"
	CONOUT "    Status : $SubMsgStatus"
	CONOUT "    Expiry : $SubMsgExpiry"
}

function DetectAdbaClient
{
	CONOUT "`nAD Activation client information:"
	CONOUT "    Object Name: $ADActivationObjectName"
	CONOUT "    Domain Name: $ADActivationObjectDN"
	CONOUT "    CSVLK Extended PID: $ADActivationCsvlkPid"
	CONOUT "    CSVLK Activation ID: $ADActivationCsvlkSkuId"
}

function DetectAvmClient
{
	CONOUT "`nAutomatic VM Activation client information:"
	if (-Not [String]::IsNullOrEmpty($IAID)) {
		CONOUT "    Guest IAID: $IAID"
	} else {
		CONOUT "    Guest IAID: Not Available"
	}
	if (-Not [String]::IsNullOrEmpty($AutomaticVMActivationHostMachineName)) {
		CONOUT "    Host machine name: $AutomaticVMActivationHostMachineName"
	} else {
		CONOUT "    Host machine name: Not Available"
	}
	if ($AutomaticVMActivationLastActivationTime.Substring(0,4) -NE "1601") {
		$EED = [DateTime]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($AutomaticVMActivationLastActivationTime),$null,48).ToString('yyyy-MM-dd hh:mm:ss tt')
		CONOUT "    Activation time: $EED UTC"
	} else {
		CONOUT "    Activation time: Not Available"
	}
	if (-Not [String]::IsNullOrEmpty($AutomaticVMActivationHostDigitalPid2)) {
		CONOUT "    Host Digital PID2: $AutomaticVMActivationHostDigitalPid2"
	} else {
		CONOUT "    Host Digital PID2: Not Available"
	}
}

function DetectKmsHost
{
	if ($Vista -Or $NT5) {
		$KeyManagementServiceListeningPort = strGetRegistry $SLKeyPath "KeyManagementServiceListeningPort"
		$KeyManagementServiceDnsPublishing = strGetRegistry $SLKeyPath "DisableDnsPublishing"
		$KeyManagementServiceLowPriority = strGetRegistry $SLKeyPath "EnableKmsLowPriority"
		if (-Not $KeyManagementServiceDnsPublishing) {$KeyManagementServiceDnsPublishing = "TRUE"}
		if (-Not $KeyManagementServiceLowPriority) {$KeyManagementServiceLowPriority = "FALSE"}
	} else {
		$KeyManagementServiceListeningPort = $objSvc.KeyManagementServiceListeningPort
		$KeyManagementServiceDnsPublishing = $objSvc.KeyManagementServiceDnsPublishing
		$KeyManagementServiceLowPriority = $objSvc.KeyManagementServiceLowPriority
	}

	if (-Not $KeyManagementServiceListeningPort) {$KeyManagementServiceListeningPort = 1688}
	if ($KeyManagementServiceDnsPublishing -EQ "TRUE") {
		$KeyManagementServiceDnsPublishing = "Enabled"
	} else {
		$KeyManagementServiceDnsPublishing = "Disabled"
	}
	if ($KeyManagementServiceLowPriority -EQ "TRUE") {
		$KeyManagementServiceLowPriority = "Low"
	} else {
		$KeyManagementServiceLowPriority = "Normal"
	}

	CONOUT "`nKey Management Service host information:"
	CONOUT "    Current count: $KeyManagementServiceCurrentCount"
	CONOUT "    Listening on Port: $KeyManagementServiceListeningPort"
	CONOUT "    DNS publishing: $KeyManagementServiceDnsPublishing"
	CONOUT "    KMS priority: $KeyManagementServiceLowPriority"
	if (-Not [String]::IsNullOrEmpty($KeyManagementServiceTotalRequests)) {
		CONOUT "`nKey Management Service cumulative requests received from clients:"
		CONOUT "    Total: $KeyManagementServiceTotalRequests"
		CONOUT "    Failed: $KeyManagementServiceFailedRequests"
		CONOUT "    Unlicensed: $KeyManagementServiceUnlicensedRequests"
		CONOUT "    Licensed: $KeyManagementServiceLicensedRequests"
		CONOUT "    Initial grace period: $KeyManagementServiceOOBGraceRequests"
		CONOUT "    Expired or Hardware out of tolerance: $KeyManagementServiceOOTGraceRequests"
		CONOUT "    Non-genuine grace period: $KeyManagementServiceNonGenuineGraceRequests"
		if ($null -NE $KeyManagementServiceNotificationRequests) {CONOUT "    Notification: $KeyManagementServiceNotificationRequests"}
	}
}

function DetectKmsClient
{
	if ($null -NE $VLActivationTypeEnabled) {CONOUT "Configured Activation Type: $($VLActTypes[$VLActivationTypeEnabled])"}
	CONOUT "`r"
	if ($LicenseStatus -NE 1) {
		CONOUT "Please activate the product in order to update KMS client information values."
		return
	}

	if ($Vista) {
		$KeyManagementServicePort = strGetRegistry $SLKeyPath "KeyManagementServicePort"
		$DiscoveredKeyManagementServiceMachineName = strGetRegistry $NSKeyPath "DiscoveredKeyManagementServiceName"
		$DiscoveredKeyManagementServiceMachinePort = strGetRegistry $NSKeyPath "DiscoveredKeyManagementServicePort"
	}

	if ([String]::IsNullOrEmpty($KeyManagementServiceMachine)) {
		$KmsReg = $null
	} else {
		if (-Not $KeyManagementServicePort) {$KeyManagementServicePort = 1688}
		$KmsReg = "Registered KMS machine name: ${KeyManagementServiceMachine}:${KeyManagementServicePort}"
	}

	if ([String]::IsNullOrEmpty($DiscoveredKeyManagementServiceMachineName)) {
		$KmsDns = "DNS auto-discovery: KMS name not available"
		if ($Vista -And -Not $Admin) {$KmsDns = "DNS auto-discovery: Run the script as administrator to retrieve info"}
	} else {
		if (-Not $DiscoveredKeyManagementServiceMachinePort) {$DiscoveredKeyManagementServiceMachinePort = 1688}
		$KmsDns = "KMS machine name from DNS: ${DiscoveredKeyManagementServiceMachineName}:${DiscoveredKeyManagementServiceMachinePort}"
	}

	if ($null -NE $objSvc.KeyManagementServiceHostCaching) {
		if ($objSvc.KeyManagementServiceHostCaching -EQ "TRUE") {
			$KeyManagementServiceHostCaching = "Enabled"
		} else {
			$KeyManagementServiceHostCaching = "Disabled"
		}
	}

	CONOUT "Key Management Service client information:"
	CONOUT "    Client Machine ID (CMID): $($objSvc.ClientMachineID)"
	if ($null -EQ $KmsReg) {
		CONOUT "    $KmsDns"
		CONOUT "    Registered KMS machine name: KMS name not available"
	} else {
		CONOUT "    $KmsReg"
	}
	if ($null -NE $DiscoveredKeyManagementServiceMachineIpAddress) {CONOUT "    KMS machine IP address: $DiscoveredKeyManagementServiceMachineIpAddress"}
	CONOUT "    KMS machine extended PID: $KeyManagementServiceProductKeyID"
	CONOUT "    Activation interval: $VLActivationInterval minutes"
	CONOUT "    Renewal interval: $VLRenewalInterval minutes"
	if ($null -NE $KeyManagementServiceHostCaching) {CONOUT "    KMS host caching: $KeyManagementServiceHostCaching"}
	if (-Not [String]::IsNullOrEmpty($KeyManagementServiceLookupDomain)) {CONOUT "    KMS SRV record lookup domain: $KeyManagementServiceLookupDomain"}
}

function GetResult($strSLP, $strSLS, $strID)
{
	try
	{
		$objPrd = [wmisearcher]"SELECT * FROM $strSLP WHERE ID='$strID'"
		$objPrd.Options.Rewindable = $false
		$objPrd.Get() | select -Expand Properties -EA 0 | foreach { if (-Not [String]::IsNullOrEmpty($_.Value)) {set $_.Name $_.Value} }
		$objPrd.Dispose()
	}
	catch
	{
		return
	}

	$winID = ($ApplicationID -EQ $winApp)
	$winPR = ($winID -And -Not $LicenseIsAddon)
	$Vista = ($winID -And $NT6 -And -Not $NT7)
	$NT5 = ($strSLP -EQ $wslp -And $winbuild -LT 6001)
	$reapp = ("Windows", "App")[!$winID]
	$prmnt = ("machine", "product")[!$winPR]

	if ($Description | Select-String "VOLUME_KMSCLIENT") {$cKmsClient = 1; $_mTag = "Volume"}
	if ($Description | Select-String "TIMEBASED_") {$cTblClient = 1; $_mTag = "Timebased"}
	if ($Description | Select-String "VIRTUAL_MACHINE_ACTIVATION") {$cAvmClient = 1; $_mTag = "Automatic VM"}
	if ($null -EQ $cKmsClient) {
		if ($Description | Select-String "VOLUME_KMS") {$cKmsHost = 1}
	}

	$_gpr = [Math]::Round($GracePeriodRemaining/1440)
	if ($_gpr -GT 0) {
		$_xpr = [DateTime]::Now.addMinutes($GracePeriodRemaining).ToString('yyyy-MM-dd hh:mm:ss tt')
	}

	if ($null -EQ $LicenseStatusReason) {$LicenseStatusReason = -1}
	$LicenseReason = '0x{0:X}' -f $LicenseStatusReason
	$LicenseMsg = "Time remaining: $GracePeriodRemaining minute(s) ($_gpr day(s))"
	if ($LicenseStatus -EQ 0) {
		$LicenseInf = "Unlicensed"
		$LicenseMsg = $null
	}
	if ($LicenseStatus -EQ 1) {
		$LicenseInf = "Licensed"
		$LicenseMsg = $null
		if ($GracePeriodRemaining -EQ 0) {
			$ExpireMsg = "The $prmnt is permanently activated."
		} else {
			$LicenseMsg = "$_mTag activation expiration: $GracePeriodRemaining minute(s) ($_gpr day(s))"
			if ($null -NE $_xpr) {$ExpireMsg = "$_mTag activation will expire $_xpr"}
		}
	}
	if ($LicenseStatus -EQ 2) {
		$LicenseInf = "Initial grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Initial grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 3) {
		$LicenseInf = "Additional grace period (KMS license expired or hardware out of tolerance)"
		if ($null -NE $_xpr) {$ExpireMsg = "Additional grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 4) {
		$LicenseInf = "Non-genuine grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Non-genuine grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 5 -And -Not $NT5) {
		$LicenseInf = "Notification"
		$LicenseMsg = "Notification Reason: $LicenseReason"
		if ($LicenseReason -EQ "0xC004F00F") {if ($null -NE $cKmsClient) {$LicenseMsg = $LicenseMsg + " (KMS license expired)."} else {$LicenseMsg = $LicenseMsg + " (hardware out of tolerance)."}}
		if ($LicenseReason -EQ "0xC004F200") {$LicenseMsg = $LicenseMsg + " (non-genuine)."}
		if ($LicenseReason -EQ "0xC004F009" -Or $LicenseReason -EQ "0xC004F064") {$LicenseMsg = $LicenseMsg + " (grace time expired)."}
	}
	if ($LicenseStatus -GT 5 -Or ($LicenseStatus -GT 4 -And $NT5)) {
		$LicenseInf = "Unknown"
		$LicenseMsg = $null
	}
	if ($LicenseStatus -EQ 6 -And -Not $Vista -And -Not $NT5) {
		$LicenseInf = "Extended grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Extended grace period ends $_xpr"}
	}

	if ($winPR -And $PartialProductKey -And -Not $NT9) {
		$dp4 = strGetRegistry "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" "DigitalProductId4"
		if ($null -NE $dp4) {
			$ProductKeyChannel = ([System.Text.Encoding]::Unicode.GetString($dp4, 1016, 128)).Trim([char]$null)
		}
	}

	if ($winPR -And $Dlv -And $NT7 -And $null -EQ $RemainingAppReArmCount) {
		try
		{
			$tmp = [wmisearcher]"SELECT RemainingWindowsReArmCount FROM $strSLS"
			$tmp.Options.Rewindable = $false
			$tmp.Get() | select -Expand Properties -EA 0 | foreach {set $_.Name $_.Value}
			$tmp.Dispose()
		}
		catch
		{
		}
	}

	$add_on = $Name.IndexOf("add-on for", 5)

	& $isAll
	if ($add_on -EQ -1) {CONOUT "Name: $Name"} else {CONOUT "Name: $($Name.Substring(0, $add_on + 7))"}
	CONOUT "Description: $Description"
	CONOUT "Activation ID: $ID"
	if ($null -NE $ProductKeyID) {CONOUT "Extended PID: $ProductKeyID"}
	if ($null -NE $ProductKeyID2 -And $Dlv) {CONOUT "Product ID: $ProductKeyID2"}
	if ($null -NE $OfflineInstallationId -And $IID) {CONOUT "Installation ID: $OfflineInstallationId"}
	if ($null -NE $ProductKeyChannel) {CONOUT "Product Key Channel: $ProductKeyChannel"}
	if ($null -NE $PartialProductKey) {CONOUT "Partial Product Key: $PartialProductKey"}
	CONOUT "License Status: $LicenseInf"
	if ($null -NE $LicenseMsg) {CONOUT "$LicenseMsg"}
	if ($LicenseStatus -NE 0 -And $EvaluationEndDate.Substring(0,4) -NE "1601") {
		$EED = [DateTime]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($EvaluationEndDate),$null,48).ToString('yyyy-MM-dd hh:mm:ss tt')
		CONOUT "Evaluation End Date: $EED UTC"
	}
	if ($Dlv) {
		if ($null -NE $RemainingWindowsReArmCount) {
			CONOUT "Remaining Windows rearm count: $RemainingWindowsReArmCount"
		}
		if ($null -NE $RemainingSkuReArmCount -And $RemainingSkuReArmCount -NE 4294967295) {
			CONOUT "Remaining $reapp rearm count: $RemainingAppReArmCount"
			CONOUT "Remaining SKU rearm count: $RemainingSkuReArmCount"
		}
		if ($null -NE $TrustedTime -And $LicenseStatus -NE 0) {
			$TTD = [DateTime]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($TrustedTime),$null,32).ToString('yyyy-MM-dd hh:mm:ss tt')
			CONOUT "Trusted time: $TTD"
		}
	}
	if ($LicenseStatus -EQ 0) {
		return
	}

	if ($strSLP -EQ $wslp -And $null -NE $PartialProductKey -And $null -NE $ADActivationObjectName -And $VLActivationType -EQ 1) {
		DetectAdbaClient
	}

	if ($winID -And $null -NE $cAvmClient -And $null -NE $PartialProductKey) {
		DetectAvmClient
	}

	$chkSub = ($winPR -And $cSub)

	$chkSLS = ($null -NE $PartialProductKey) -And ($null -NE $cKmsClient -Or $null -NE $cKmsHost -Or $chkSub)

	if (!$chkSLS) {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
		return
	}

	try
	{
		$objSvc = New-Object PSObject
		$wmiSvc = [wmisearcher]"SELECT * FROM $strSLS"
		$wmiSvc.Options.Rewindable = $false
		$wmiSvc.Get() | select -Expand Properties -EA 0 | foreach { if (-Not [String]::IsNullOrEmpty($_.Value)) {$objSvc | Add-Member 8 $_.Name $_.Value} }
		$wmiSvc.Dispose()
		if ($null -EQ $IsKeyManagementServiceMachine) {$objSvc.PSObject.Properties | foreach {set $_.Name $_.Value}}
	}
	catch
	{
		return
	}

	if ($strSLS -EQ $wsls -And $NT9) {
		if ([String]::IsNullOrEmpty($DiscoveredKeyManagementServiceMachineIpAddress)) {
			$DiscoveredKeyManagementServiceMachineIpAddress = "not available"
		}
	}

	if ($null -NE $cKmsHost -And $IsKeyManagementServiceMachine -GT 0) {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
		DetectKmsHost
	}

	if ($null -NE $cKmsClient) {
		DetectKmsClient
	}

	if ($null -EQ $cKmsHost) {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
	}

	if ($chkSub) {
		DetectSubscription
	}

}
#endregion

#region vNextDiag
if ($PSVersionTable.PSVersion.Major -Lt 3)
{
	function ConvertFrom-Json
	{
		[CmdletBinding()]
		Param(
			[Parameter(ValueFromPipeline=$true)][Object]$item
		)
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
		$psjs = New-Object System.Web.Script.Serialization.JavaScriptSerializer
		Return ,$psjs.DeserializeObject($item)
	}
	function ConvertTo-Json
	{
		[CmdletBinding()]
		Param(
			[Parameter(ValueFromPipeline=$true)][Object]$item
		)
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
		$psjs = New-Object System.Web.Script.Serialization.JavaScriptSerializer
		Return $psjs.Serialize($item)
	}
}

function PrintModePerPridFromRegistry
{
	$vNextRegkey = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext"
	$vNextPrids = Get-Item -Path $vNextRegkey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'property' -ErrorAction SilentlyContinue | Where-Object -FilterScript {$_.ToLower() -like "*retail" -or $_.ToLower() -like "*volume"}
	If ($null -Eq $vNextPrids)
	{
		CONOUT "`nNo registry keys found."
		Return
	}
	CONOUT "`r"
	$vNextPrids | ForEach `
	{
		$mode = (Get-ItemProperty -Path $vNextRegkey -Name $_).$_
		Switch ($mode)
		{
			2 { $mode = "vNext"; Break }
			3 { $mode = "Device"; Break }
			Default { $mode = "Legacy"; Break }
		}
		CONOUT "$_ = $mode"
	}
}

function PrintSharedComputerLicensing
{
	$scaRegKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	$scaValue = Get-ItemProperty -Path $scaRegKey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	$scaRegKey2 = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing"
	$scaValue2 = Get-ItemProperty -Path $scaRegKey2 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	$scaPolicyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Licensing"
	$scaPolicyValue = Get-ItemProperty -Path $scaPolicyKey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	If ($null -Eq $scaValue -And $null -Eq $scaValue2 -And $null -Eq $scaPolicyValue)
	{
		CONOUT "`nNo registry keys found."
		Return
	}
	$scaModeValue = $scaValue -Or $scaValue2 -Or $scaPolicyValue
	If ($scaModeValue -Eq 0)
	{
		$scaMode = "Disabled"
	}
	If ($scaModeValue -Eq 1)
	{
		$scaMode = "Enabled"
	}
	CONOUT "`nStatus: $scaMode"
	CONOUT "`r"
	$tokenFiles = $null
	$tokenPath = "${env:LOCALAPPDATA}\Microsoft\Office\16.0\Licensing"
	If (Test-Path $tokenPath)
	{
		$tokenFiles = Get-ChildItem -Path $tokenPath -Filter "*authString*" -Recurse | Where-Object { !$_.PSIsContainer }
	}
	If ($null -Eq $tokenFiles -Or $tokenFiles.Length -Eq 0)
	{
		CONOUT "No tokens found."
		Return
	}
	$tokenFiles | ForEach `
	{
		$tokenParts = (Get-Content -Encoding Unicode -Path $_.FullName).Split('_')
		$output = New-Object PSObject
		$output | Add-Member 8 'ACID' $tokenParts[0];
		$output | Add-Member 8 'User' $tokenParts[3];
		$output | Add-Member 8 'NotBefore' $tokenParts[4];
		$output | Add-Member 8 'NotAfter' $tokenParts[5];
		Write-Output $output
	}
}

function PrintLicensesInformation
{
	Param(
		[ValidateSet("NUL", "Device")]
		[String]$mode
	)
	If ($mode -Eq "NUL")
	{
		$licensePath = "${env:LOCALAPPDATA}\Microsoft\Office\Licenses"
	}
	ElseIf ($mode -Eq "Device")
	{
		$licensePath = "${env:PROGRAMDATA}\Microsoft\Office\Licenses"
	}
	$licenseFiles = $null
	If (Test-Path $licensePath)
	{
		$licenseFiles = Get-ChildItem -Path $licensePath -Recurse | Where-Object { !$_.PSIsContainer }
	}
	If ($null -Eq $licenseFiles -Or $licenseFiles.Length -Eq 0)
	{
		CONOUT "`nNo licenses found."
		Return
	}
	$licenseFiles | ForEach `
	{
		$license = (Get-Content -Encoding Unicode $_.FullName | ConvertFrom-Json).License
		$decodedLicense = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($license)) | ConvertFrom-Json
		$licenseType = $decodedLicense.LicenseType
		If ($null -Ne $decodedLicense.ExpiresOn)
		{
			$expiry = [System.DateTime]::Parse($decodedLicense.ExpiresOn, $null, 'AdjustToUniversal')
		}
		Else
		{
			$expiry = New-Object System.DateTime
		}
		$licenseState = "Grace"
		If ((Get-Date) -Gt (Get-Date $decodedLicense.Metadata.NotAfter))
		{
			$licenseState = "RFM"
		}
		ElseIf ((Get-Date) -Lt (Get-Date $expiry))
		{
			$licenseState = "Licensed"
		}
		$output = New-Object PSObject
		$output | Add-Member 8 'File' $_.PSChildName;
		$output | Add-Member 8 'Version' $_.Directory.Name;
		$output | Add-Member 8 'Type' "User|${licenseType}";
		$output | Add-Member 8 'Product' $decodedLicense.ProductReleaseId;
		$output | Add-Member 8 'Acid' $decodedLicense.Acid;
		If ($mode -Eq "Device") { $output | Add-Member 8 'DeviceId' $decodedLicense.Metadata.DeviceId; }
		$output | Add-Member 8 'LicenseState' $licenseState;
		$output | Add-Member 8 'EntitlementStatus' $decodedLicense.Status;
		$output | Add-Member 8 'EntitlementExpiration' ("N/A", $decodedLicense.ExpiresOn)[!($null -eq $decodedLicense.ExpiresOn)];
		$output | Add-Member 8 'ReasonCode' ("N/A", $decodedLicense.ReasonCode)[!($null -eq $decodedLicense.ReasonCode)];
		$output | Add-Member 8 'NotBefore' $decodedLicense.Metadata.NotBefore;
		$output | Add-Member 8 'NotAfter' $decodedLicense.Metadata.NotAfter;
		$output | Add-Member 8 'NextRenewal' $decodedLicense.Metadata.RenewAfter;
		$output | Add-Member 8 'TenantId' ("N/A", $decodedLicense.Metadata.TenantId)[!($null -eq $decodedLicense.Metadata.TenantId)];
		#$output.PSObject.Properties | foreach { $ht = @{} } { $ht[$_.Name] = $_.Value } { $output = $ht | ConvertTo-Json }
		Write-Output $output
	}
}

function vNextDiagRun
{
	$fNUL = ([IO.Directory]::Exists("${env:LOCALAPPDATA}\Microsoft\Office\Licenses")) -and ([IO.Directory]::GetFiles("${env:LOCALAPPDATA}\Microsoft\Office\Licenses", "*", 1).Length -GE 0)
	$fDev = ([IO.Directory]::Exists("${env:PROGRAMDATA}\Microsoft\Office\Licenses")) -and ([IO.Directory]::GetFiles("${env:PROGRAMDATA}\Microsoft\Office\Licenses", "*", 1).Length -GE 0)
	$rPID = $null -NE (GP "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext" -EA 0 | select -Expand 'property' -EA 0 | where -Filter {$_.ToLower() -like "*retail" -or $_.ToLower() -like "*volume"})
	$rSCA = $null -NE (GP "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -EA 0 | select -Expand "SharedComputerLicensing" -EA 0)
	$rSCL = $null -NE (GP "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing" -EA 0 | select -Expand "SharedComputerLicensing" -EA 0)

	if (($fNUL -Or $fDev -Or $rPID -Or $rSCA -Or $rSCL) -EQ $false) {
		Return
	}

	& $isAll
	CONOUT "$line2"
	CONOUT "===                  Office vNext Status                 ==="
	CONOUT "$line2"
	CONOUT "`n========== Mode per ProductReleaseId =========="
	PrintModePerPridFromRegistry
	CONOUT "`n========== Shared Computer Licensing =========="
	PrintSharedComputerLicensing
	CONOUT "`n========== vNext licenses ==========="
	PrintLicensesInformation -Mode "NUL"
	CONOUT "`n========== Device licenses =========="
	PrintLicensesInformation -Mode "Device"
	CONOUT "$line3"
	CONOUT "`r"
}
#endregion

#region clic

<#
;;; Source: https://github.com/asdcorp/clic
;;; Powershell port: abbodi1406

Copyright 2023 asdcorp

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

function BoolToWStr($bVal) {
	("TRUE", "FALSE")[!$bVal]
}

function InitializePInvoke {
	$Marshal = [System.Runtime.InteropServices.Marshal]
	$Module = [AppDomain]::CurrentDomain.DefineDynamicAssembly((Get-Random), 'Run').DefineDynamicModule((Get-Random))

	$Class = $Module.DefineType('NativeMethods', 'Public, Abstract, Sealed, BeforeFieldInit', [Object], 0)
	$Class.DefinePInvokeMethod('SLIsWindowsGenuineLocal', 'slc.dll', 'Public, Static', 'Standard', [Int32], @([UInt32].MakeByRefType()), 'Winapi', 'Unicode').SetImplementationFlags('PreserveSig')
	$Class.DefinePInvokeMethod('SLGetWindowsInformationDWORD', 'slc.dll', 22, 1, [Int32], @([String], [UInt32].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGetWindowsInformation', 'slc.dll', 22, 1, [Int32], @([String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)

	if ($DllSubscription) {
		$Class.DefinePInvokeMethod('ClipGetSubscriptionStatus', 'Clipc.dll', 22, 1, [Int32], @([IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
		$Struct = $Class.DefineNestedType('SubStatus', 'NestedPublic, SequentialLayout, Sealed, BeforeFieldInit', [ValueType], 0)
		[void]$Struct.DefineField('dwEnabled', [UInt32], 'Public')
		[void]$Struct.DefineField('dwSku', [UInt32], 6)
		[void]$Struct.DefineField('dwState', [UInt32], 6)
		$SubStatus = $Struct.CreateType()
	}

	$Win32 = $Class.CreateType()
}

function InitializeDigitalLicenseCheck {
	$CAB = [System.Reflection.Emit.CustomAttributeBuilder]

	$ICom = $Module.DefineType('EUM.IEUM', 'Public, Interface, Abstract, Import')
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.ComImportAttribute].GetConstructor(@()), @()))
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.GuidAttribute].GetConstructor(@([String])), @('F2DCB80D-0670-44BC-9002-CD18688730AF')))
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.InterfaceTypeAttribute].GetConstructor(@([Int16])), @([Int16]1)))

	1..4 | % { [void]$ICom.DefineMethod('VF'+$_, 'Public, Virtual, HideBySig, NewSlot, Abstract', 'Standard, HasThis', [Void], @()) }
	[void]$ICom.DefineMethod('AcquireModernLicenseForWindows', 1478, 33, [Int32], @([Int32], [Int32].MakeByRefType()))

	$IEUM = $ICom.CreateType()
}

function PrintStateData {
	$pwszStateData = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-Action-StateData",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pwszStateData
	)) {
		return $FALSE
	}

	[string[]]$pwszStateString = $Marshal::PtrToStringUni($pwszStateData) -replace ";", "`n    "
	CONOUT ("    $pwszStateString")

	$Marshal::FreeHGlobal($pwszStateData)
	return $TRUE
}

function PrintLastActivationHResult {
	$pdwLastHResult = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-LastWindowsActivationHResult",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pdwLastHResult
	)) {
		return $FALSE
	}

	CONOUT ("    LastActivationHResult=0x{0:x8}" -f $Marshal::ReadInt32($pdwLastHResult))

	$Marshal::FreeHGlobal($pdwLastHResult)
	return $TRUE
}

function PrintLastActivationTime {
	$pdwLastTime = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-LastWindowsActivationTime",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pdwLastTime
	)) {
		return $FALSE
	}

	$actTime = $Marshal::ReadInt64($pdwLastTime)
	if ($actTime -ne 0) {
		CONOUT ("    LastActivationTime={0}" -f [DateTime]::FromFileTimeUtc($actTime).ToString("yyyy/MM/dd:HH:mm:ss"))
	}

	$Marshal::FreeHGlobal($pdwLastTime)
	return $TRUE
}

function PrintIsWindowsGenuine {
	$dwGenuine = 0
	$ppwszGenuineStates = @(
		"SL_GEN_STATE_IS_GENUINE",
		"SL_GEN_STATE_INVALID_LICENSE",
		"SL_GEN_STATE_TAMPERED",
		"SL_GEN_STATE_OFFLINE",
		"SL_GEN_STATE_LAST"
	)

	if ($Win32::SLIsWindowsGenuineLocal([ref]$dwGenuine)) {
		return $FALSE
	}

	if ($dwGenuine -lt 5) {
		CONOUT ("    IsWindowsGenuine={0}" -f $ppwszGenuineStates[$dwGenuine])
	} else {
		CONOUT ("    IsWindowsGenuine={0}" -f $dwGenuine)
	}

	return $TRUE
}

function PrintDigitalLicenseStatus {
	try {
		. InitializeDigitalLicenseCheck
		$ComObj = New-Object -Com EditionUpgradeManagerObj.EditionUpgradeManager
	} catch {
		return $FALSE
	}

	$parameters = 1, $null

	if ([EUM.IEUM].GetMethod("AcquireModernLicenseForWindows").Invoke($ComObj, $parameters)) {
		return $FALSE
	}

	$dwReturnCode = $parameters[1]
	[bool]$bDigitalLicense = $FALSE

	$bDigitalLicense = (($dwReturnCode -ge 0) -and ($dwReturnCode -ne 1))
	CONOUT ("    IsDigitalLicense={0}" -f (BoolToWStr $bDigitalLicense))

	return $TRUE
}

function PrintSubscriptionStatus {
	$dwSupported = 0

	if ($winbuild -ge 15063) {
		$pwszPolicy = "ConsumeAddonPolicySet"
	} else {
		$pwszPolicy = "Allow-WindowsSubscription"
	}

	if ($Win32::SLGetWindowsInformationDWORD($pwszPolicy, [ref]$dwSupported)) {
		return $FALSE
	}

	CONOUT ("    SubscriptionSupportedEdition={0}" -f (BoolToWStr $dwSupported))

	$pStatus = $Marshal::AllocHGlobal($Marshal::SizeOf([Type]$SubStatus))
	if ($Win32::ClipGetSubscriptionStatus([ref]$pStatus)) {
		return $FALSE
	}

	$sStatus = [Activator]::CreateInstance($SubStatus)
	$sStatus = $Marshal::PtrToStructure($pStatus, [Type]$SubStatus)
	$Marshal::FreeHGlobal($pStatus)

	CONOUT ("    SubscriptionEnabled={0}" -f (BoolToWStr $sStatus.dwEnabled))

	if ($sStatus.dwEnabled -eq 0) {
		return $TRUE
	}

	CONOUT ("    SubscriptionSku={0}" -f $sStatus.dwSku)
	CONOUT ("    SubscriptionState={0}" -f $sStatus.dwState)

	return $TRUE
}

function ClicRun
{
	& $isAll
	CONOUT "Client Licensing Check information:"

	$null = PrintStateData
	$null = PrintLastActivationHResult
	$null = PrintLastActivationTime
	$null = PrintIsWindowsGenuine

	if ($DllDigital) {
		$null = PrintDigitalLicenseStatus
	}

	if ($DllSubscription) {
		$null = PrintSubscriptionStatus
	}

	CONOUT "$line3"
	& $noAll
}
#endregion

$Host.UI.RawUI.WindowTitle = "Check Activation Status"
if ($All.IsPresent) {
	$B=$Host.UI.RawUI.BufferSize;$B.Height=3000;$Host.UI.RawUI.BufferSize=$B;
	if (!$Pass.IsPresent) {clear;}
}

$SysPath = "$env:SystemRoot\System32"
if (Test-Path "$env:SystemRoot\Sysnative\reg.exe") {
	$SysPath = "$env:SystemRoot\Sysnative"
}

$wslp = "SoftwareLicensingProduct"
$wsls = "SoftwareLicensingService"
$oslp = "OfficeSoftwareProtectionProduct"
$osls = "OfficeSoftwareProtectionService"
$winApp = "55c92734-d682-4d71-983e-d6ec3f16059f"
$o14App = "59a52881-a989-479d-af46-f275c6370663"
$o15App = "0ff1ce15-a989-479d-af46-f275c6370663"
$cSub = ($winbuild -GE 19041) -And (Select-String -Path "$SysPath\wbem\sppwmi.mof" -Encoding unicode -Pattern "SubscriptionType")
$DllDigital = ($winbuild -GE 14393) -And (Test-Path "$SysPath\EditionUpgradeManagerObj.dll")
$DllSubscription = ($winbuild -GE 14393) -And (Test-Path "$SysPath\Clipc.dll")
$VLActTypes = @("All", "AD", "KMS", "Token")
$SLKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"
$NSKeyPath = "HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"

'cW1nd0ws', 'c0ff1ce15', 'c0ff1ce14', 'ospp14', 'ospp15' | foreach {set $_ $false}

$offsvc = "osppsvc"
if ($NT7 -Or -Not $NT6) {$winsvc = "sppsvc"} else {$winsvc = "slsvc"}

try {gsv $winsvc -EA 1 | Out-Null; $WsppHook = 1} catch {$WsppHook = 0}
try {gsv $offsvc -EA 1 | Out-Null; $OsppHook = 1} catch {$OsppHook = 0}

if ($WsppHook -NE 0) {
	try {sasv $winsvc -EA 1} catch {}
	$cW1nd0ws  = DetectID $wslp $winApp
	$c0ff1ce15 = DetectID $wslp $o15App
	$c0ff1ce14 = DetectID $wslp $o14App
}

if ($OsppHook -NE 0) {
	try {sasv $offsvc -EA 1} catch {}
	$ospp15 = DetectID $oslp $o15App
	$ospp14 = DetectID $oslp $o14App
}

if ($cW1nd0ws)
{
	echoWindows
	GetID $wslp $winApp | foreach -EA 1 {
	GetResult $wslp $wsls $_
	CONOUT "$line3"
	& $noAll
	}
}
elseif ($NT6)
{
	echoWindows
	CONOUT "`nError: product key not found."
}

if ($winbuild -GE 9200) {
	. InitializePInvoke
	ClicRun
}

if ($c0ff1ce15 -Or $ospp15) {
	CheckOhook
}

$doMSG = 1

if ($c0ff1ce15)
{
	echoOffice
	GetID $wslp $o15App | foreach -EA 1 {
	GetResult $wslp $wsls $_
	CONOUT "$line3"
	& $noAll
	}
}

if ($c0ff1ce14)
{
	echoOffice
	GetID $wslp $o14App | foreach -EA 1 {
	GetResult $wslp $wsls $_
	CONOUT "$line3"
	& $noAll
	}
}

if ($ospp15)
{
	echoOffice
	GetID $oslp $o15App | foreach -EA 1 {
	GetResult $oslp $osls $_
	CONOUT "$line3"
	& $noAll
	}
}

if ($ospp14)
{
	echoOffice
	GetID $oslp $o14App | foreach -EA 1 {
	GetResult $oslp $osls $_
	CONOUT "$line3"
	& $noAll
	}
}

if ($NT7) {
	vNextDiagRun
}

ExitScript 0
:sppmgr:

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:troubleshoot

set "line=_________________________________________________________________________________________________"

:at_menu

cls
title  Troubleshoot %masver%
if not defined terminal mode 77, 30

echo:
echo:
echo:
echo:
echo:       _______________________________________________________________
echo:                                                   
call :dk_color2 %_White% "             [1] " %_Green% "Help"
echo:             ___________________________________________________
echo:                                                                      
echo:             [2] Dism RestoreHealth
echo:             [3] SFC Scannow
echo:                                                                      
echo:             [4] Fix WMI
echo:             [5] Fix Licensing
echo:             [6] Fix WPA Registry
echo:             ___________________________________________________
echo:
echo:             [0] %_exitmsg%
echo:       _______________________________________________________________
echo:          
call :dk_color2 %_White% "            " %_Green% "Choose a menu option using your keyboard :"
choice /C:1234560 /N
set _erl=%errorlevel%

if %_erl%==7 exit /b
if %_erl%==6 start %mas%fix-wpa-registry &goto at_menu
if %_erl%==5 goto:retokens
if %_erl%==4 goto:fixwmi
if %_erl%==3 goto:sfcscan
if %_erl%==2 goto:dism_rest
if %_erl%==1 start %mas%troubleshoot.html &goto at_menu
goto :at_menu

::========================================================================================================================================

:dism_rest

cls
if not defined terminal mode 98, 30
title  Dism /English /Online /Cleanup-Image /RestoreHealth

if %winbuild% LSS 9200 (
%eline%
echo Unsupported OS version detected.
echo This command only works on Windows 8/8.1/10/11 and their Server equivalents.
goto :at_back
)

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not "%%#"=="" set _int=1)
)

echo:
if defined _int (
echo      Checking Internet Connection  [Connected]
) else (
call :dk_color2 %_White% "     " %Red% "Checking Internet Connection  [Not connected]"
)

echo %line%
echo:
echo      DISM uses Windows Update to provide replacement files required to fix corruption.
echo      This will take 5-15 minutes or more..
echo %line%
echo:
echo      Notes:
echo:
call :dk_color2 %_White% "     - " %Gray% "Make sure the internet is connected."
call :dk_color2 %_White% "     - " %Gray% "Make sure that Windows update is properly working."
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
if not defined terminal mode 110, 30

for /f %%a in ('%psc% "(Get-Date).ToString('yyyyMMdd-HHmmssfff')"') do set _time=%%a

%psc% Stop-Service TrustedInstaller -force %nul%

copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\backup_cbs_%_time%.log" %nul%
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "%SystemRoot%\logs\DISM\backup_dism_%_time%.log" %nul%
del /f /q "%SystemRoot%\logs\cbs\cbs.log" %nul%
del /f /q "%SystemRoot%\logs\DISM\dism.log" %nul%

echo:
echo Applying the command...
echo dism /english /online /cleanup-image /restorehealth
dism /english /online /cleanup-image /restorehealth

timeout /t 5 %nul1%
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\cbs_%_time%.log" %nul%
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "%SystemRoot%\logs\DISM\dism_%_time%.log" %nul%

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%
call :compresslog cbs\cbs_%_time%.log AT_Logs\RHealth_CBS %nul%
call :compresslog DISM\dism_%_time%.log AT_Logs\RHealth_DISM %nul%

if not exist "!desktop!\AT_Logs\RHealth_CBS_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "!desktop!\AT_Logs\RHealth_CBS_%_time%.log" %nul%
)

if not exist "!desktop!\AT_Logs\RHealth_DISM_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "!desktop!\AT_Logs\RHealth_DISM_%_time%.log" %nul%
)

echo:
call :dk_color %Gray% "CBS and DISM logs are copied to the AT_Logs folder on your desktop."
goto :at_back

::========================================================================================================================================

:sfcscan

cls
if not defined terminal mode 98, 30
title  sfc /scannow

echo:
echo %line%
echo:    
echo      SFC will repair missing or corrupted system files.
echo      It is recommended you run the DISM option first before this one.
echo      This will take 10-15 minutes or more..
echo:
echo      If SFC could not fix something, then run the command again to see if it may be able 
echo      to the next time. Sometimes it may take running the sfc /scannow command 3 times
echo      restarting the PC after each time to completely fix everything that it's able to.
echo:   
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
for /f %%a in ('%psc% "(Get-Date).ToString('yyyyMMdd-HHmmssfff')"') do set _time=%%a

%psc% Stop-Service TrustedInstaller -force %nul%

copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\backup_cbs_%_time%.log" %nul%
del /f /q "%SystemRoot%\logs\cbs\cbs.log" %nul%

echo:
echo Applying the command...
echo sfc /scannow
sfc /scannow

timeout /t 5 %nul1%
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\cbs_%_time%.log" %nul%

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%
call :compresslog cbs\cbs_%_time%.log AT_Logs\SFC_CBS %nul%

if not exist "!desktop!\AT_Logs\SFC_CBS_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "!desktop!\AT_Logs\SFC_CBS_%_time%.log" %nul%
)

echo:
call :dk_color %Gray% "The CBS log was copied to the AT_Logs folder on your Desktop."
goto :at_back

::========================================================================================================================================

:retokens

cls
if not defined terminal (
mode 125, 32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)
title  Fix Licensing ^(ClipSVC ^+ SPP ^+ OSPP^)

echo:
echo %line%
echo:   
echo      Notes:
echo:
echo       - This option helps in troubleshooting activation issues.
echo:
echo       - This option will:
echo            - Deactivate Windows and Office, you may need to reactivate.
echo              If Windows is activated with motherboard / OEM / Digital license
echo              then Windows will activate itself again.
echo:
echo            - Clear ClipSVC, SPP and OSPP licenses.
echo            - Fix permissions of SPP tokens folder and registries.
echo            - Trigger the repair option for Office.
echo:
call :dk_color2 %_White% "      - " %Red% "Apply this option only when it is necessary."
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

::========================================================================================================================================

::  Rebuild ClipSVC Licences

cls
:cleanlicensing

echo:
echo %line%
echo:
call :dk_color %Blue% "Rebuilding ClipSVC Licenses..."
echo:

if %winbuild% LSS 10240 (
echo ClipSVC license rebuilding is supported only on Windows 10/11 and their Server equivalents.
echo Skipping...
goto :rebuildspptok
)

%psc% "(([WMISEARCHER]'SELECT Name FROM SoftwareLicensingProduct WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').Get()).Name" %nul2% | findstr /i "Windows" %nul1% && (
echo Windows is permanently activated.
echo Skipping...
goto :rebuildspptok
)

echo Stopping ClipSVC service...
%psc% Stop-Service ClipSVC -force %nul%
timeout /t 2 %nul%

echo:
echo Applying the command to clean ClipSVC Licenses...
echo rundll32 clipc.dll,ClipCleanUpState

rundll32 clipc.dll,ClipCleanUpState

if %winbuild% LEQ 10240 (
echo [Successful]
) else (
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat" (
call :dk_color %Red% "[Failed]"
) else (
echo [Successful]
)
)

::  Below registry key (Volatile & Protected) gets created after the ClipSVC License cleanup command, and gets automatically deleted after 
::  system restart. It needs to be deleted to activate the system without restart.

set "RegKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState"
set "_ident=HKU\S-1-5-19\SOFTWARE\Microsoft\IdentityCRL"

reg query "%RegKey%" %nul% && %nul% call :regownstart
reg delete "%RegKey%" /f %nul% 

echo:
echo Deleting a Volatile ^& Protected Registry Key...
echo [%RegKey%]
reg query "%RegKey%" %nul% && (
call :dk_color %Red% "[Failed]"
echo Reboot your machine using the restart option, that will delete this registry key automatically.
) || (
echo [Successful]
)

::   Clear HWID token related registry to fix activation incase there is any corruption

echo:
echo Deleting IdentityCRL Registry Key...
echo [%_ident%]
reg delete "%_ident%" /f %nul%
reg query "%_ident%" %nul% && (
call :dk_color %Red% "[Failed]"
) || (
echo [Successful]
)

%psc% Stop-Service ClipSVC -force %nul%

::  Rebuild ClipSVC folder to fix permission issues

echo:
if %winbuild% GTR 10240 (
echo Deleting folder %ProgramData%\Microsoft\Windows\ClipSVC\
rmdir /s /q "C:\ProgramData\Microsoft\Windows\ClipSvc" %nul%

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\" (
call :dk_color %Red% "[Failed]"
) else (
echo [Successful]
)

echo:
echo Rebuilding the %ProgramData%\Microsoft\Windows\ClipSVC\ folder...
%psc% Start-Service ClipSVC %nul%
timeout /t 3 %nul%
if not exist "%ProgramData%\Microsoft\Windows\ClipSVC\" timeout /t 5 %nul%
if not exist "%ProgramData%\Microsoft\Windows\ClipSVC\" (
call :dk_color %Red% "[Failed]"
) else (
echo [Successful]
)
)

echo:
echo Restarting wlidsvc ^& LicenseManager services...
for %%# in (wlidsvc LicenseManager) do (%psc% "Start-Job { Restart-Service %%# } | Wait-Job -Timeout 20 | Out-Null")

::========================================================================================================================================

::  Rebuild SPP Tokens

:rebuildspptok

echo:
echo %line%
echo:
call :dk_color %Blue% "Rebuilding SPP licensing tokens..."
echo:

call :scandat check

if not defined token (
call :dk_color %Red% "tokens.dat file not found."
) else (
echo tokens.dat file: [%token%]
)

set tokenstore=
set badregistry=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore %nul6%') do call set "tokenstore=%%b"
if %winbuild% GEQ 9200 if /i not "%tokenstore%"=="%SysPath%\spp\store" if /i not "%tokenstore%"=="%SysPath%\spp\store\2.0" if /i not "%tokenstore%"=="%SysPath%\spp\store_test\2.0" (
set badregistry=1
echo:
call :dk_color %Red% "Correct path not found in TokenStore Registry [%tokenstore%]"
)

::  Check sppsvc permissions and apply fixes

if %winbuild% GEQ 9200 if not defined badregistry (
echo:
echo Checking SPP permission related issues...
call :checkperms
if defined permerror (
call :dk_color %Red% "[!permerror!]"
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':fixsppperms\:.*';iex ($f[1])" %nul%
call :checkperms
if defined permerror (
call :dk_color %Red% "[!permerror!] [Failed To Fix]"
) else (
call :dk_color %Green% "[Successfully Fixed]"
)
) else (
echo [No Error Found]
)
)

echo:
echo Stopping sppsvc service...
%psc% Stop-Service sppsvc -force %nul%

set w=
set _sppint=
for %%# in (SppEx%w%tComObj.exe sppsvc.exe) do (reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%w%ge File Execu%w%tion Options\%%#" %nul% && (set _sppint=1))
if defined _sppint (
echo:
echo Removing SPP IFEO registry keys...
for %%# in (SppE%w%xtComObj.exe sppsvc.exe) do (reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%w%ge File Execu%w%tion Options\%%#" /f %nul%)
)

if %winbuild% LSS 9200 (
REM Fix issues caused by Update KB971033 in Windows 7
REM https://support.microsoft.com/help/4487266
echo:
echo Checking Update KB971033...
%psc% "if (Get-Hotfix -Id KB971033 -ErrorAction SilentlyContinue) {Exit 3}" %nul%
if !errorlevel!==3 (
echo Found, uninstalling it...
wusa /uninstall /quiet /norestart /kb:971033
) else (
echo [Not Found]
)
%psc% Stop-Service sppuinotify -force %nul%
sc config sppuinotify start= disabled
del /f /q %SysPath%\7B296FB0-376B-497e-B012-9C450E1B7327-*.C7483456-A289-439d-8115-601632D005A0 /ah
)

::  Delete registry keys that are not deleted by activation scripts

echo:
echo Cleaning some licensing-related registry keys...
%nul% reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "ServiceSessionId" /f
%nul% reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "LicStatusArray" /f
%nul% reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "PolicyValuesArray" /f
%nul% reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "actionlist" /f
%nul% reg delete "HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\data" /f

echo:
call :scandat delete
call :scandat check

if defined token (
echo:
call :dk_color %Red% "Failed to delete .dat files."
echo:
)

echo:
echo Reinstalling system licenses...
%psc% "Stop-Service sppsvc -force; $sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
if %errorlevel% NEQ 0 %psc% "$sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
if %errorlevel% EQU 0 (
echo [Successful]
) else (
call :dk_color %Red% "[Failed]"
)

call :scandat check

echo:
if not defined token (
call :dk_color %Red% "Failed to rebuild tokens.dat file."
) else (
echo tokens.dat file was rebuilt successfully.
)

if %winbuild% LSS 9200 (
sc config sppuinotify start= demand
)

::========================================================================================================================================

::  Rebuild OSPP Tokens

echo:
echo %line%
echo:
call :dk_color %Blue% "Rebuilding OSPP licensing tokens..."
echo:

sc qc osppsvc %nul% || (
echo OSPP-based Office is not installed.
echo Skipping rebuilding OSPP tokens...
goto :repairoffice
)

call :scandatospp check

if not defined token (
call :dk_color %Red% "tokens.dat file not found."
) else (
echo tokens.dat file: [%token%]
)

echo:
echo Stopping osppsvc service...
%psc% Stop-Service osppsvc -force %nul%

echo:
call :scandatospp delete
call :scandatospp check

if defined token (
echo:
call :dk_color %Red% "Failed to delete .dat files."
echo:
)

echo:
echo Starting osppsvc service to generate tokens.dat...
%psc% Start-Service osppsvc %nul%
call :scandatospp check
if not defined token (
%psc% Stop-Service osppsvc -force %nul%
%psc% Start-Service osppsvc %nul%
timeout /t 3 %nul%
)

call :scandatospp check

echo:
if not defined token (
call :dk_color %Red% "Failed to rebuild tokens.dat file."
) else (
echo tokens.dat file was rebuilt successfully.
)

::========================================================================================================================================

:repairoffice

echo:
echo %line%
echo:
call :dk_color %Blue% "Repairing Office licenses..."
echo:

for %%# in (68 86) do (
for %%A in (msi14 msi15 msi16 c2r14 c2r15 c2r16) do (set %%A_%%#=&set %%Arepair%%#=)
)

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

reg query %_68%\14.0\CVH /f Click2run /k %nul% && (set "c2r14_68=Office 14.0 C2R x86/x64"  & set "c2r14repair68=")
reg query %_86%\14.0\CVH /f Click2run /k %nul% && (set "c2r14_86=Office 14.0 C2R x86"      & set "c2r14repair86=")

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi14_86=Office 14.0 MSI x86"      & call :getrepairsetup msi14repair86 14)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi14_68=Office 14.0 MSI x86/x64"  & call :getrepairsetup msi14repair68 14)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi15_86=Office 15.0 MSI x86"      & call :getrepairsetup msi15repair86 15)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi15_68=Office 15.0 MSI x86/x64"  & call :getrepairsetup msi15repair68 15)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi16_86=Office 16.0 MSI x86"      & call :getrepairsetup msi16repair86 16)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi16_68=Office 16.0 MSI x86/x64"  & call :getrepairsetup msi16repair68 16)

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set "c2r15_86=Office 15.0 C2R x86"      & call :getc2rrepair c2r15repair86 integratedoffice.exe)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set "c2r15_68=Office 15.0 C2R x86/x64"  & call :getc2rrepair c2r15repair68 integratedoffice.exe)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set "c2r16_86=Office 16.0 C2R x86"      & call :getc2rrepair c2r16repair86 OfficeClickToRun.exe)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set "c2r16_68=Office 16.0 C2R x86/x64"  & call :getc2rrepair c2r16repair68 OfficeClickToRun.exe)

set uwp16=
if %winbuild% GEQ 10240 (
%psc% "Get-AppxPackage -name "Microsoft.Office.Desktop"" | find /i "Office" %nul1% && set uwp16=Office 16.0 UWP
)

set /a counter=0
echo Checking installed Office versions...
echo:

for %%# in (
"%msi14_68%"
"%msi14_86%"
"%msi15_68%"
"%msi15_86%"
"%msi16_68%"
"%msi16_86%"
"%c2r14_68%"
"%c2r14_86%"
"%c2r15_68%"
"%c2r15_86%"
"%c2r16_68%"
"%c2r16_86%"
"%uwp16%"
) do (
if not "%%#"=="""" (
set insoff=%%#
set insoff=!insoff:"=!
echo [!insoff!]
set /a counter+=1
)
)

if %counter% GTR 1 (
%eline%
echo Multiple Office versions found.
echo It is recommended to only install one version of Office.
echo ________________________________________________________________
echo:
)

if %counter% EQU 0 (
echo:
echo Office ^(2010 and later^) is not installed.
goto :repairend
echo:
) else (
echo:
call :dk_color %_Yellow% "A new window will appear, in that window you need to select [Quick Repair] option."
if defined terminal (
call :dk_color %_Yellow% "Press [0] to continue..."
choice /c 0 /n
) else (
call :dk_color %_Yellow% "Press any key to continue..."
pause %nul1%
)
)

if defined uwp16 (
echo:
echo Skipping repair for Office 16.0 UWP... 
echo:
)

set c2r14=
if defined c2r14_68 set c2r14=1
if defined c2r14_86 set c2r14=1

if defined c2r14 (
echo:
echo Skipping repair for Office 14.0 C2R...
echo:
)

if defined msi14_68 if exist "%msi14repair68%" echo Running - "%msi14repair68%"                    & "%msi14repair68%"
if defined msi14_86 if exist "%msi14repair86%" echo Running - "%msi14repair86%"                    & "%msi14repair86%"
if defined msi15_68 if exist "%msi15repair68%" echo Running - "%msi15repair68%"                    & "%msi15repair68%"
if defined msi15_86 if exist "%msi15repair86%" echo Running - "%msi15repair86%"                    & "%msi15repair86%"
if defined msi16_68 if exist "%msi16repair68%" echo Running - "%msi16repair68%"                    & "%msi16repair68%"
if defined msi16_86 if exist "%msi16repair86%" echo Running - "%msi16repair86%"                    & "%msi16repair86%"
if defined c2r15_68 if exist "%c2r15repair68%" echo Running - "%c2r15repair68%" REPAIRUI RERUNMODE & "%c2r15repair68%" REPAIRUI RERUNMODE
if defined c2r15_86 if exist "%c2r15repair86%" echo Running - "%c2r15repair86%" REPAIRUI RERUNMODE & "%c2r15repair86%" REPAIRUI RERUNMODE
if defined c2r16_68 if exist "%c2r16repair68%" echo Running - "%c2r16repair68%" scenario=Repair    & "%c2r16repair68%" scenario=Repair
if defined c2r16_86 if exist "%c2r16repair86%" echo Running - "%c2r16repair86%" scenario=Repair    & "%c2r16repair86%" scenario=Repair

:repairend

echo:
echo %line%
echo:
echo:
call :dk_color %Green% "Finished"
goto :at_back

:getc2rrepair

for %%# in (X86 X64) do (
if exist "%systemdrive%\Program Files\Microsoft Office 15\Client%%#\%2" (
set "%1=%systemdrive%\Program Files\Microsoft Office 15\Client%%#\%2"
)
)
exit /b

:getrepairsetup

set "_common86=%systemdrive%\Program Files (x86)\Common Files\Microsoft Shared\OFFICE%2\Office Setup Controller\setup.exe"
set "_common68=%systemdrive%\Program Files\Common Files\Microsoft Shared\OFFICE%2\Office Setup Controller\setup.exe"

if exist "%_common86%" set "%1=%_common86%"
if exist "%_common68%" set "%1=%_common68%"
exit /b

::========================================================================================================================================

:fixwmi

cls
if not defined terminal mode 98, 34
title  Fix WMI

::  https://techcommunity.microsoft.com/t5/ask-the-performance-team/wmi-repository-corruption-or-not/ba-p/375484

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo Rebuilding WMI is not recommended on Windows Server, aborting...
goto :at_back
)

echo:
echo Checking WMI
call :checkwmi

::  Apply basic fix first and check

if defined error (
%psc% Stop-Service Winmgmt -force %nul%
winmgmt /salvagerepository %nul%
call :checkwmi
)

if not defined error (
echo [Working]
echo No need to apply this option, aborting...
goto :at_back
)

call :dk_color %Red% "[Not Responding]"

set _corrupt=
sc start Winmgmt %nul%
if %errorlevel% EQU 1060 set _corrupt=1
sc query Winmgmt %nul% || set _corrupt=1
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type) do if not defined _corrupt (reg query HKLM\SYSTEM\CurrentControlSet\Services\Winmgmt /v %%G %nul% || set _corrupt=1)

echo:
if defined _corrupt (
%eline%
echo Winmgmt service is corrupted, aborting...
goto :at_back
)

echo Disabling Winmgmt service
sc config Winmgmt start= disabled %nul%
if %errorlevel% EQU 0 (
echo [Successful]
) else (
call :dk_color %Red% "[Failed] Aborting..."
sc config Winmgmt start= auto %nul%
goto :at_back
)

echo:
echo Stopping Winmgmt service
%psc% Stop-Service Winmgmt -force %nul%
%psc% Stop-Service Winmgmt -force %nul%
%psc% Stop-Service Winmgmt -force %nul%
sc query Winmgmt | find /i "STOPPED" %nul% && (
echo [Successful]
) || (
call :dk_color %Red% "[Failed]"
echo:
call :dk_color %Blue% "Its recommended to select [Restart] option and then apply Fix WMI option again."
echo %line%
echo:
choice /C:21 /N /M "> [1] Restart  [2] Revert Back Changes :"
if !errorlevel!==1 (sc config Winmgmt start= auto %nul%&goto :at_back)
echo:
echo Restarting...
shutdown -t 5 -r
exit
)

echo:
echo Deleting WMI repository
rmdir /s /q "%SysPath%\wbem\repository\" %nul%
if exist "%SysPath%\wbem\repository\" (
call :dk_color %Red% "[Failed]"
) else (
echo [Successful]
)

echo:
echo Enabling Winmgmt service
sc config Winmgmt start= auto %nul%
if %errorlevel% EQU 0 (
echo [Successful]
) else (
call :dk_color %Red% "[Failed]"
)

call :checkwmi
if not defined error (
echo:
echo Checking WMI
call :dk_color %Green% "[Working]"
goto :at_back
)

echo:
echo Registering .dll's and Compiling .mof's, .mfl's
call :registerobj %nul%

echo:
echo Checking WMI
call :checkwmi
if defined error (
call :dk_color %Red% "[Not Responding]"
echo:
echo Run [Dism RestoreHealth] and [SFC Scannow] options and make sure there are no errors.
) else (
call :dk_color %Green% "[Working]"
)

goto :at_back

:registerobj

::  https://eskonr.com/2012/01/how-to-fix-wmi-issues-automatically/

%psc% Stop-Service Winmgmt -force %nul%
cd /d %SysPath%\wbem\
regsvr32 /s %SysPath%\scecli.dll
regsvr32 /s %SysPath%\userenv.dll
mofcomp cimwin32.mof
mofcomp cimwin32.mfl
mofcomp rsop.mof
mofcomp rsop.mfl
for /f %%s in ('dir /b /s *.dll') do regsvr32 /s %%s
for /f %%s in ('dir /b *.mof') do mofcomp %%s
for /f %%s in ('dir /b *.mfl') do mofcomp %%s

winmgmt /salvagerepository
winmgmt /resetrepository
exit /b

:checkwmi

::  https://learn.microsoft.com/en-us/windows/win32/wmisdk/wmi-error-constants

set error=
%psc% "Get-WmiObject -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" %nul2% | find /i "computersystem" %nul1%
if %errorlevel% NEQ 0 (set error=1& exit /b)
winmgmt /verifyrepository %nul%
if %errorlevel% NEQ 0 (set error=1& exit /b)

%psc% "try { $null=([WMISEARCHER]'SELECT * FROM SoftwareLicensingService').Get().Version; exit 0 } catch { exit $_.Exception.InnerException.HResult }" %nul%
cmd /c exit /b %errorlevel%
echo "0x%=ExitCode%" | findstr /i "0x800410 0x800440 0x80131501" %nul1%
if %errorlevel% EQU 0 set error=1
exit /b

::========================================================================================================================================

:at_back

echo:
echo %line%
echo:
if defined terminal (
call :dk_color %_Yellow% "Press [0] key to %_exitmsg%..."
choice /c 0 /n
) else (
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
pause %nul1%
)
goto :at_menu

::========================================================================================================================================

:compresslog

::  https://stackoverflow.com/a/46268232

set "ddf="%SystemRoot%\Temp\ddf""
%nul% del /q /f %ddf%
echo/.New Cabinet>%ddf%
echo/.set Cabinet=ON>>%ddf%
echo/.set CabinetFileCountThreshold=0;>>%ddf%
echo/.set Compress=ON>>%ddf%
echo/.set CompressionType=LZX>>%ddf%
echo/.set CompressionLevel=7;>>%ddf%
echo/.set CompressionMemory=21;>>%ddf%
echo/.set FolderFileCountThreshold=0;>>%ddf%
echo/.set FolderSizeThreshold=0;>>%ddf%
echo/.set GenerateInf=OFF>>%ddf%
echo/.set InfFileName=nul>>%ddf%
echo/.set MaxCabinetSize=0;>>%ddf%
echo/.set MaxDiskFileCount=0;>>%ddf%
echo/.set MaxDiskSize=0;>>%ddf%
echo/.set MaxErrors=1;>>%ddf%
echo/.set RptFileName=nul>>%ddf%
echo/.set UniqueFiles=ON>>%ddf%
for /f "tokens=* delims=" %%D in ('dir /a:-D/b/s "%SystemRoot%\logs\%1"') do (
 echo/"%%~fD"  /inf=no;>>%ddf%
)
makecab /F %ddf% /D DiskDirectory1="" /D CabinetNameTemplate="!desktop!\%2_%_time%.cab"
del /q /f %ddf%
exit /b

::========================================================================================================================================

:checkperms

::  This code checks if SPP has permission access to tokens folder and required registry keys. Incorrect permissions are often set by gaming spoofers.

set permerror=
if not exist "%tokenstore%\" set "permerror=Error Found In Token Folder"

if defined ps32onArm exit /b

for %%# in (
"%tokenstore%+FullControl"
"HKLM:\SYSTEM\WPA+QueryValues, EnumerateSubKeys, WriteKey"
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform+SetValue"
) do for /f "tokens=1,2 delims=+" %%A in (%%#) do if not defined permerror (
%psc% "$acl = (Get-Acl '%%A' | fl | Out-String); if (-not ($acl -match 'NT SERVICE\\sppsvc Allow  %%B') -or ($acl -match 'NT SERVICE\\sppsvc Deny')) {Exit 2}" %nul%
if !errorlevel!==2 (
if "%%A"=="%tokenstore%" (
set "permerror=Error Found In Token Folder"
) else (
set "permerror=Error Found In SPP Registries"
)
)
)

REM  https://learn.microsoft.com/office/troubleshoot/activation/license-issue-when-start-office-application

if not defined permerror (
reg query "HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion" %nul% && (
set "pol=HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies"
reg query "!pol!" %nul% || reg add "!pol!" %nul%
%psc% "$netServ = (New-Object Security.Principal.SecurityIdentifier('S-1-5-20')).Translate([Security.Principal.NTAccount]).Value; $aclString = Get-Acl 'Registry::!pol!' | Format-List | Out-String; if (-not ($aclString.Contains($netServ + ' Allow  FullControl') -or $aclString.Contains('NT SERVICE\sppsvc Allow  FullControl')) -or ($aclString.Contains('Deny'))) {Exit 3}" %nul%
if !errorlevel!==3 set "permerror=Error Found In S-1-5-20 SPP"
)
)
exit /b

::========================================================================================================================================

::  Fix SPP related registry and folder permissions

:fixsppperms:
# Fix perms for Token Folder

if ($env:permerror -eq 'Error Found In Token Folder') {
    New-Item -Path $env:tokenstore -ItemType Directory -Force
    $sddl = 'O:BAG:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICIIO;GR;;;BU)(A;;FR;;;BU)(A;OICI;FA;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)'
    $AclObject = New-Object System.Security.AccessControl.DirectorySecurity
    $AclObject.SetSecurityDescriptorSddlForm($sddl)
    Set-Acl -Path $env:tokenstore -AclObject $AclObject
    exit
}

# Fix perms for SPP registries

if ($env:permerror -eq 'Error Found In SPP Registries') {
    $acl = Get-Acl 'HKLM:\SYSTEM\WPA'
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc', 'QueryValues, EnumerateSubKeys, WriteKey', 'ContainerInherit, ObjectInherit', 'None', 'Allow')
    $acl.ResetAccessRule($rule)
    $acl.SetAccessRule($rule)
    Set-Acl -Path 'HKLM:\SYSTEM\WPA' -AclObject $acl
	
    $acl = Get-Acl 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc', 'SetValue', 'ContainerInherit, ObjectInherit', 'None', 'Allow')
    $acl.ResetAccessRule($rule)
    $acl.SetAccessRule($rule)
    Set-Acl -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform' -AclObject $acl
    exit
}

# Fix perms for SPP in HKU\S-1-5-20
# https://learn.microsoft.com/office/troubleshoot/activation/license-issue-when-start-office-application

if ($env:permerror -ne 'Error Found In S-1-5-20 SPP') {
    exit
}
if (-not (Test-Path 'Registry::HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform')) {
    exit
}

# https://stackoverflow.com/a/35843420

function Take-Permissions {
    param($rootKey, $key, [System.Security.Principal.SecurityIdentifier]$sid = 'S-1-5-32-545', $recurse = $true)
    
    switch -regex ($rootKey) {
        'HKCU|HKEY_CURRENT_USER' { $rootKey = 'CurrentUser' }
        'HKLM|HKEY_LOCAL_MACHINE' { $rootKey = 'LocalMachine' }
        'HKCR|HKEY_CLASSES_ROOT' { $rootKey = 'ClassesRoot' }
        'HKCC|HKEY_CURRENT_CONFIG' { $rootKey = 'CurrentConfig' }
        'HKU|HKEY_USERS' { $rootKey = 'Users' }
    }

    ### Step 1 - escalate current process's privilege
    # get SeTakeOwnership, SeBackup and SeRestore privileges before executes next lines, script needs Admin privilege
    $AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
    $TypeBuilder = $ModuleBuilder.DefineType(0)
    $TypeBuilder.DefinePInvokeMethod('RtlAdjustPrivilege', 'ntdll.dll', 'Public, Static', 1, [int], @([int], [bool], [bool], [bool].MakeByRefType()), 1, 3) | Out-Null
    9, 17, 18 | ForEach-Object { $TypeBuilder.CreateType()::RtlAdjustPrivilege($_, $true, $false, [ref]$false) | Out-Null }

    function Take-KeyPermissions {
        param($rootKey, $key, $sid, $recurse, $recurseLevel = 0)

        ### Step 2 - get ownerships of key - it works only for current key
        $regKey = [Microsoft.Win32.Registry]::$rootKey.OpenSubKey($key, 'ReadWriteSubTree', 'TakeOwnership')
        $acl = New-Object System.Security.AccessControl.RegistrySecurity
        $acl.SetOwner($sid)
        $regKey.SetAccessControl($acl)

        ### Step 3 - enable inheritance of permissions (not ownership) for current key from parent
        $acl.SetAccessRuleProtection($false, $false)
        $regKey.SetAccessControl($acl)

        ### Step 4 - only for top-level key, change permissions for current key and propagate it for subkeys
        # to enable propagations for subkeys, it needs to execute Steps 2-3 for each subkey (Step 5)
        if ($recurseLevel -eq 0) {
            $regKey = $regKey.OpenSubKey('', 'ReadWriteSubTree', 'ChangePermissions')
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule($sid, 'FullControl', 'ContainerInherit', 'None', 'Allow')
            $acl.ResetAccessRule($rule)
            $regKey.SetAccessControl($acl)
        }

        ### Step 5 - recursively repeat steps 2-5 for subkeys
        if ($recurse) {
            foreach ($subKey in $regKey.OpenSubKey('').GetSubKeyNames()) {
                Take-KeyPermissions $rootKey ($key + '\' + $subKey) $sid $recurse ($recurseLevel + 1)
            }
        }
    }

    Take-KeyPermissions $rootKey $key $sid $recurse
}

Take-Permissions "Users" "S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" "S-1-5-20"
:fixsppperms:

::========================================================================================================================================

:scandat

set token=
for %%# in (
%SysPath%\spp\store_test\2.0\
%SysPath%\spp\store\
%SysPath%\spp\store\2.0\
%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\
) do (

if %1==check (
if exist %%#tokens.dat set token=%%#tokens.dat
)

if %1==delete (
if exist %%# (
%nul% dir /a-d /s "%%#*.dat" && (
attrib -r -s -h "%%#*.dat" /S
del /S /F /Q "%%#*.dat"
)
)
)
)
exit /b

:scandatospp

set token=
for %%# in (
%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform\
) do (

if %1==check (
if exist %%#tokens.dat set token=%%#tokens.dat
)

if %1==delete (
if exist %%# (
%nul% dir /a-d /s "%%#*.dat" && (
attrib -r -s -h "%%#*.dat" /S
del /S /F /Q "%%#*.dat"
)
)
)
)
exit /b

::========================================================================================================================================

:regownstart

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regown\:.*';iex ($f[1]);"
exit /b

::  Below code takes ownership of a volatile registry key and deletes it
::  HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState

:regown:
$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TypeBuilder = $ModuleBuilder.DefineType(0)

$TypeBuilder.DefinePInvokeMethod('RtlAdjustPrivilege', 'ntdll.dll', 'Public, Static', 1, [int], @([int], [bool], [bool], [bool].MakeByRefType()), 1, 3) | Out-Null
$TypeBuilder.CreateType()::RtlAdjustPrivilege(9, $true, $false, [ref]$false) | Out-Null

$SID = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
$IDN = ($SID.Translate([System.Security.Principal.NTAccount])).Value
$Admin = New-Object System.Security.Principal.NTAccount($IDN)

$path = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState'
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey($path, 'ReadWriteSubTree', 'takeownership')

$acl = $key.GetAccessControl()
$acl.SetOwner($Admin)
$key.SetAccessControl($acl)

$rule = New-Object System.Security.AccessControl.RegistryAccessRule($Admin,"FullControl","Allow")
$acl.SetAccessRule($rule)
$key.SetAccessControl($acl)
:regown:

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:change_winedition

::  To stage current edition while changing edition with CBS Upgrade Method, change 0 to 1 in below line
set _stg=0

set "line=echo ___________________________________________________________________________________________"

cls
if not defined terminal mode 98, 30
title  Change Windows Edition %masver%

echo:
echo Initializing...
echo:

for %%# in (
sppsvc.exe
dism.exe
) do (
if not exist %SysPath%\%%# (
%eline%
echo [%SysPath%\%%#] file is missing, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)
)

::========================================================================================================================================

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_reflection
call :dk_ckeckwmic
call :dk_sppissue

for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if not "%%j"=="" (
set fullbuild=%%i.%%j
) else (
for /f "tokens=3" %%G in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR" %nul6%') do if not errorlevel 1 set /a "UBR=%%G"
for /f "skip=2 tokens=3,4 delims=. " %%G in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildLabEx') do (
if defined UBR (set "fullbuild=%%G.!UBR!") else (set "fullbuild=%%G.%%H")
)
)

::========================================================================================================================================

::  Check Activation IDs

call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f
if not defined allapps (
%eline%
echo Failed to find activation IDs. Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

::  Check Windows Edition and branch

set osedition=
set dismnotworking=

for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition %nul6% ^| find /i "Current Edition :"') do set "osedition=%%a"
if not defined osedition set dismnotworking=1

if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL) get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
if not defined osedition %chkedi% do if not errorlevel 1 (call set "osedition=%%a")

if not defined osedition (
%eline%
echo Failed to detect OS edition, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

for /f "skip=2 tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul6%') do set "regedition=%%a"
if /i not "%osedition%"=="%regedition%" (
set "showeditionerror=call :dk_color %_Yellow% "[%osedition%] [Reg-%regedition%].""
)

set branch=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildBranch %nul6%') do set "branch=%%b"

::========================================================================================================================================

::  Get target editions list

set _target=
set _dtarget=
set _ptarget=
set _ntarget=
set _wtarget=

if %winbuild% GEQ 10240 for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (if defined _dtarget (set "_dtarget= !_dtarget! %%a ") else (set "_dtarget= %%a "))
if %winbuild% LSS 10240 for /f "tokens=4" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':cbsxml\:.*';& ([ScriptBlock]::Create($f[1])) -GetTargetEditions;" ^| findstr /i /c:"Target Edition : "') do (if defined _ptarget (set "_ptarget= !_ptarget! %%a ") else (set "_ptarget= %%a "))

if %winbuild% GEQ 10240 if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
if %winbuild% GEQ 17063 call :ced_edilist
if /i "%osedition:~0,4%"=="Core" set _pro=Professional
if /i "%osedition%"=="CoreN" set _pro=ProfessionalN
set "_dtarget= %_dtarget% !_wtarget! !_pro! "
)

::========================================================================================================================================

for %%# in (CloudEdition CloudEditionN ServerRdsh) do if /i %osedition%==%%# (
cls
echo:
call :dk_color %Red% "==== Note ===="
echo:
echo [EditionID:%osedition% ^| %fullbuild%]
echo:
echo Changing this edition may not remove "%osedition%"-specific features.
echo:
call :dk_color %_Yellow% "Press [7] to continue anyway..."
choice /c 7 /n
cls
)

for %%# in ( %_dtarget% %_ptarget% ) do if /i not "%%#"=="%osedition%" (
echo "!_target!" | find /i " %%# " %nul1% || set "_target= !_target! %%# "
)

if defined _target (
for %%# in (%_target%) do (
echo %%# | findstr /i "CountrySpecific CloudEdition" %nul% || (set "_ntarget=!_ntarget! %%#")
)
)

if not defined _ntarget (
%line%
echo:
if defined dismnotworking call :dk_color %Red% "DISM.exe is not working."
call :dk_color %Gray% "Target editions not found."
echo Current Edition [%osedition% ^| %winbuild%] can not be changed to any other Edition.
%line%
goto dk_done
)

::========================================================================================================================================

:cedmenu2

cls
if not defined terminal mode 98, 30
set inpt=
set counter=0
set verified=0
set targetedition=

%line%
echo:
call :dk_color %Gray% "You can change the edition [%osedition%] [%fullbuild%] to one of the following."
%showeditionerror%
if defined dismnotworking (
call :dk_color %_Yellow% "Note - DISM.exe is not working."
if /i "%osedition:~0,4%"=="Core" call :dk_color %_Yellow% "     - You will see more edition options to choose once its changed to Pro."
)
%line%
echo:

for %%A in (%_ntarget%) do (
set /a counter+=1
echo [!counter!]  %%A
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  %_exitmsg%
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto cedmenu2
if "%inpt%"=="0" exit /b
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto cedmenu2

::========================================================================================================================================

if %winbuild% LSS 10240 goto :cbsmethod
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" goto :ced_change_server

cls
if not defined terminal mode con cols=105 lines=32

if /i "%targetedition%"=="ServerRdsh" (
echo:
call :dk_color %Red% "==== Note ===="
echo:
echo Once the edition is changed to "%targetedition%", 
echo the system may not be able to properly change edition later.
echo:
echo [1] Continue Anyway
echo [0] Go Back
echo:
call :dk_color %_Green% "Choose a menu option using your keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 goto cedmenu2
if !errorlevel!==1 rem
)

cls
set key=
set _chan=
set _dismapi=0

::  Check if DISM API or slmgr.vbs is required for edition upgrade

if not exist "%SysPath%\spp\tokens\skus\%targetedition%\%targetedition%*.xrm-ms" (
echo %_wtarget% | find /i " %targetedition% " || (
set _dismapi=1
)
)

set "keyflow=Retail Volume:GVLK Volume:MAK OEM:NONSLP OEM:DM PGS:TB Retail:TB:Eval"

call :ced_targetSKU %targetedition%
if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%
if not defined key call :changeeditiondata
if not defined key if %_dismapi%==1 if /i "%targetedition%"=="Professional" (
set key=VK7JG-NPHTM-C97JM-9MPGT-3V66T
set _chan=Retail
)

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Failed to get product key from pkeyhelper.dll.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

::  Changing from Core to Non-Core & Changing editions in Windows build older than 17134 requires "changepk /productkey" or DISM Api method and restart
::  In other cases, editions can be changed instantly with "slmgr /ipk"

if %_dismapi%==1 (
if not defined terminal mode con cols=105 lines=40
call :ced_rebootflag
if defined rebootreq goto dk_done
)

cls
%line%
echo:
%showeditionerror%
if defined dismnotworking call :dk_color %_Yellow% "DISM.exe is not working."
echo Changing the current edition [%osedition%] %fullbuild% to [%targetedition%]...
echo:

if %_dismapi%==1 (
call :dk_color %Green% "Notes -"
echo:
echo  - Save your work before continuing, the system will auto-restart.
echo:
echo  - You will need to activate with HWID option once the edition is changed.
%line%
echo:
choice /C:21 /N /M "[1] Continue [2] %_exitmsg% : "
if !errorlevel!==1 exit /b
)

::========================================================================================================================================

if %_dismapi%==0 (
echo Installing %_chan% key [%key%]
echo:
if %_wmic% EQU 1 wmic path %sps% where __CLASS='%sps%' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "try { $null=(([WMISEARCHER]'SELECT Version FROM %sps%').Get()).InstallProductKey('%key%'); exit 0 } catch { exit $_.Exception.InnerException.HResult }" %nul%
set keyerror=!errorlevel!
cmd /c exit /b !keyerror!
if !keyerror! NEQ 0 set "keyerror=[0x!=ExitCode!]"

if !keyerror! EQU 0 (
call :dk_refresh
call :dk_color %Green% "[Successful]"
echo:
call :dk_color %Gray% "Reboot is required to fully change the edition."
) else (
call :dk_color %Red% "[Unsuccessful] [Error Code: !keyerror!]"
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)
)

if %_dismapi%==1 (
echo:
echo Applying the DISM API method with %_chan% key %key%. Please wait...
echo:

call :ced_prep
if defined preperror goto dk_done

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':dismapi\:.*';& ([ScriptBlock]::Create($f[1])) %targetedition% %key%"
call :ced_postprep
)
%line%

goto dk_done

::========================================================================================================================================

:cbsmethod

cls
if not defined terminal (
mode con cols=105 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)

call :ced_rebootflag
if defined rebootreq goto dk_done

echo:
%showeditionerror%
if defined dismnotworking call :dk_color %_Yellow% "Note - DISM.exe is not working."
echo Changing the current edition [%osedition%] %fullbuild% to [%targetedition%]...
echo:
call :dk_color %Blue% "Important - Save your work before continuing, the system will auto-restart."
echo:
choice /C:01 /N /M "[1] Continue [0] %_exitmsg% : "
if %errorlevel%==1 exit /b

echo:
echo Initializing...
echo:

call :ced_prep
if defined preperror goto dk_done

if %_stg%==0 (set stage=) else (set stage=-StageCurrent)
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':cbsxml\:.*';& ([ScriptBlock]::Create($f[1])) -SetEdition %targetedition% %stage%"
call :ced_postprep
%line%

goto dk_done

::========================================================================================================================================

:ced_change_server

cls
if not defined terminal (
mode con cols=105 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)

set key=
set _chan=
set "keyflow=Volume:GVLK Retail Volume:MAK OEM:NONSLP OEM:DM PGS:TB Retail:TB:Eval"

call :ced_targetSKU %targetedition%
if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%
if not defined key call :changeeditiondata

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Failed to get product key from pkeyhelper.dll.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

call :ced_rebootflag
if defined rebootreq goto dk_done

cls
echo:
%showeditionerror%
if defined dismnotworking call :dk_color %_Yellow% "Note - DISM.exe is not working."
echo Changing the current edition [%osedition%] %fullbuild% to [%targetedition%]...
echo:

call :ced_prep
if defined preperror goto dk_done

echo Applying the command with %_chan% key...
echo DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula
DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula

call :ced_postprep
%line%

goto dk_done

::========================================================================================================================================

:ced_prep

set _time=
set preperror=

for /f %%a in ('%psc% "(Get-Date).ToString('yyyyMMdd-HHmmssfff')"') do set _time=%%a

%psc% Stop-Service TrustedInstaller -force %nul%

sc query TrustedInstaller | find /i "RUNNING" %nul% && (
%eline%
echo Failed to stop the TrustedInstaller service.
echo Reboot your machine using the restart option and try again.
set preperror=1
exit /b
)

copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\backup_cbs_%_time%.log" %nul%
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "%SystemRoot%\logs\DISM\backup_dism_%_time%.log" %nul%

del /f /q "%SystemRoot%\logs\cbs\cbs.log" %nul%
del /f /q "%SystemRoot%\logs\DISM\dism.log" %nul%

:: Initiate this to appear in fresh logs

dism /online /english /Get-CurrentEdition %nul%
dism /online /english /Get-TargetEditions %nul%
exit /b

::========================================================================================================================================

:ced_postprep

timeout /t 5 %nul1%
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\cbs_%_time%.log" %nul%
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "%SystemRoot%\logs\DISM\dism_%_time%.log" %nul%

if not exist "!desktop!\ChangeEdition_Logs\" md "!desktop!\ChangeEdition_Logs\" %nul%
call :compresslog cbs\cbs_%_time%.log ChangeEdition_Logs\CBS %nul%
call :compresslog DISM\dism_%_time%.log ChangeEdition_Logs\DISM %nul%

echo:
if %winbuild% GEQ 9200 %psc% "if ((Get-WindowsOptionalFeature -Online -FeatureName NetFx3).State -eq 'Enabled') {Write-Host 'Checking .NET Framework 3.5 Status - Enabled'}"
echo Log files are copied to the ChangeEdition_Logs folder on your desktop.
echo:
call :dk_color %Blue% "In case there are errors, you should restart the system before trying again."
echo:
set fixes=%fixes% %mas%change_edition_issues
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%change_edition_issues"
exit /b

::========================================================================================================================================

::  Get Edition list

:ced_edilist

if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL) get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND LicenseDependsOn is NULL').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
%chkedi% do call set "_wtarget= !_wtarget! %%a "
exit /b

::========================================================================================================================================

::  Check pending reboot flags

:ced_rebootflag

set rebootreq=
reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" %nul% && set rebootreq=1
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" %nul% && set rebootreq=1

if defined rebootreq (
%eline%
echo Pending reboot flags found.
echo:
echo Make sure Windows is fully updated, restart the system and try again.
)
exit /b

::========================================================================================================================================

:ced_windowskey

for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b
for %%# in (%keyflow%) do (
call :k_pkey %targetSKU% '%%#'
if defined pkey call :k_pkeychannel !pkey!
if /i "!pkeychannel!"=="%%#" (
set key=!pkey!
exit /b
)
)
exit /b

::========================================================================================================================================

:ced_targetSKU

set k=%1
set targetSKU=
for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b

call :dk_reflection

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('GetEditionIdFromName', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([String], [int].MakeByRefType()), 1, 3);
set d1=%d1% $out = 0; [void]$TypeBuilder.CreateType()::GetEditionIdFromName('%k%', [ref]$out); $out

for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set targetSKU=%%a)
if "%targetSKU%"=="0" set targetSKU=
exit /b

::========================================================================================================================================

::  https://github.com/asdcorp/Set-WindowsCbsEdition

:cbsxml:[
param (
    [Parameter()]
    [String]$SetEdition,

    [Parameter()]
    [Switch]$GetTargetEditions,

    [Parameter()]
    [Switch]$StageCurrent
)

function Get-AssemblyIdentity {
    param (
        [String]$PackageName
    )

    $PackageName = [String]$PackageName
    $packageData = ($PackageName -split '~')

    if($packageData[3] -eq '') {
        $packageData[3] = 'neutral'
    }

    return "<assemblyIdentity name=`"$($packageData[0])`" version=`"$($packageData[4])`" processorArchitecture=`"$($packageData[2])`" publicKeyToken=`"$($packageData[1])`" language=`"$($packageData[3])`" />"
}

function Get-SxsName {
    param (
        [String]$PackageName
    )

    $name = ($PackageName -replace '[^A-z0-9\-\._]', '')

    if($name.Length -gt 40) {
        $name = ($name[0..18] -join '') + '\.\.' + ($name[-19..-1] -join '')
    }

    return $name.ToLower()
}

function Find-EditionXmlInSxs {
    param (
        [String]$Edition
    )

    $candidates = @($Edition, 'Client', 'Server')
    $winSxs = $Env:SystemRoot + '\WinSxS'
    $allInSxs = Get-ChildItem -Path $winSxs | select Name

    foreach($candidate in $candidates) {
        $name = Get-SxsName -PackageName "Microsoft-Windows-Editions-$candidate"
        $packages = $allInSxs | where name -Match ('^.*_'+$name+'_31bf3856ad364e35')

        if($packages.Length -eq 0) {
            continue
        }

        $package = $packages[-1].Name
        $testPath = $winSxs + "\$package\" + $Edition + 'Edition.xml'

        if(Test-Path -Path $testPath -PathType Leaf) {
            return $testPath
        }
    }

    return $null
}

function Find-EditionXml {
    param (
        [String]$Edition
    )

    $servicingEditions = $Env:SystemRoot + '\servicing\Editions'
    $editionXml = $Edition + 'Edition.xml'

    $editionXmlInServicing = $servicingEditions + '\' + $editionXml

    if(Test-Path -Path $editionXmlInServicing -PathType Leaf) {
        return $editionXmlInServicing
    }

    return Find-EditionXmlInSxs -Edition $Edition
}

function Write-UpgradeCandidates {
    param (
        [HashTable]$InstallCandidates
    )

    $editionCount = 0
    Write-Host 'Editions that can be upgraded to:'
    foreach($candidate in $InstallCandidates.Keys) {
        Write-Host "Target Edition : $candidate"
        $editionCount++
    }

    if($editionCount -eq 0) {
        Write-Host '(no editions are available)'
    }
}

function Write-UpgradeXml {
    param (
        [Array]$RemovalCandidates,
        [Array]$InstallCandidates,
        [Boolean]$Stage
    )

    $removeAction = 'remove'
    if($Stage) {
        $removeAction = 'stage'
    }

    Write-Output '<?xml version="1.0"?>'
    Write-Output '<unattend xmlns="urn:schemas-microsoft-com:unattend">'
    Write-Output '<servicing>'

    foreach($package in $InstallCandidates) {
        Write-Output '<package action="install">'
        Write-Output (Get-AssemblyIdentity -PackageName $package)
        Write-Output '</package>'
    }

    foreach($package in $RemovalCandidates) {
        Write-Output "<package action=`"$removeAction`">"
        Write-Output (Get-AssemblyIdentity -PackageName $package)
        Write-Output '</package>'
    }

    Write-Output '</servicing>'
    Write-Output '</unattend>'
}

function Write-Usage {
    Get-Help $script:MyInvocation.MyCommand.Path -detailed
}

$version = '1.0'
$getTargetsParam = $GetTargetEditions.IsPresent
$stageCurrentParam = $StageCurrent.IsPresent

if($SetEdition -eq '' -and ($false -eq $getTargetsParam)) {
    Write-Usage
    Exit 1
}

$removalCandidates = @();
$installCandidates = @{};

$packages = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages' | select Name | where { $_.name -match '^.*\\Microsoft-Windows-.*Edition~' }
foreach($package in $packages) {
    $state = (Get-ItemProperty -Path "Registry::$($package.Name)").CurrentState
    $packageName = ($package.Name -split '\\')[-1]
    $packageEdition = (($packageName -split 'Edition~')[0] -split 'Microsoft-Windows-')[-1]

    if($state -eq 0x40) {
        if($null -eq $installCandidates[$packageEdition]) {
            $installCandidates[$packageEdition] = @()
        }

        if($false -eq ($installCandidates[$packageEdition] -contains $packageName)) {
            $installCandidates[$packageEdition] = $installCandidates[$packageEdition] + @($packageName)
        }
    }

    if((($state -eq 0x50) -or ($state -eq 0x70)) -and ($false -eq ($removalCandidates -contains $packageName))) {
        $removalCandidates = $removalCandidates + @($packageName)
    }
}

if($getTargetsParam) {
    Write-UpgradeCandidates -InstallCandidates $installCandidates
    Exit
}

if($false -eq ($installCandidates.Keys -contains $SetEdition)) {
    Write-Error "The system cannot be upgraded to `"$SetEdition`""
    Exit 1
}

$xmlPath = $Env:SystemRoot + '\Temp' + '\CbsUpgrade.xml'

Write-UpgradeXml -RemovalCandidates $removalCandidates `
    -InstallCandidates $installCandidates[$SetEdition] `
    -Stage $stageCurrentParam >$xmlPath

$editionXml = Find-EditionXml -Edition $SetEdition
if($null -eq $editionXml) {
    Write-Warning 'Unable to find edition specific settings XML. Proceeding without it...'
}

Write-Host 'Starting the upgrade process. This may take a while...'

DISM.EXE /English /NoRestart /Online /Apply-Unattend:$xmlPath
$dismError = $LASTEXITCODE

Remove-Item -Path $xmlPath -Force

if(($dismError -ne 0) -and ($dismError -ne 3010)) {
    Write-Error 'Failed to upgrade to the target edition'
    Exit $dismError
}

if($null -ne $editionXml) {
    $destination = $Env:SystemRoot + '\' + $SetEdition + '.xml'
    Copy-Item -Path $editionXml -Destination $destination

    DISM.EXE /English /NoRestart /Online /Apply-Unattend:$editionXml
    $dismError = $LASTEXITCODE

    if(($dismError -ne 0) -and ($dismError -ne 3010)) {
        Write-Error 'Failed to apply edition specific settings'
        Exit $dismError
    }
}

Restart-Computer
:cbsxml:]

::========================================================================================================================================

::  Change edition using DISM API
::  Thanks to Alex (aka may, ave9858)

:dismapi:[
param (
    [Parameter()]
    [String]$TargetEdition,

    [Parameter()]
    [String]$Key
)

$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TB = $ModuleBuilder.DefineType(0)

[void]$TB.DefinePInvokeMethod('DismInitialize', 'DismApi.dll', 22, 1, [int], @([int], [IntPtr], [IntPtr]), 1, 3)
[void]$TB.DefinePInvokeMethod('DismOpenSession', 'DismApi.dll', 22, 1, [int], @([String], [IntPtr], [IntPtr], [UInt32].MakeByRefType()), 1, 3)
[void]$TB.DefinePInvokeMethod('_DismSetEdition', 'DismApi.dll', 22, 1, [int], @([UInt32], [String], [String], [IntPtr], [IntPtr], [IntPtr]), 1, 3)
$Dism = $TB.CreateType()

[void]$Dism::DismInitialize(2, 0, 0)
$Session = 0
[void]$Dism::DismOpenSession('DISM_{53BFAE52-B167-4E2F-A258-0A37B57FF845}', 0, 0, [ref]$Session)
if (!$Dism::_DismSetEdition($Session, "$TargetEdition", "$Key", 0, 0, 0)) {
    Restart-Computer
}
:dismapi:]

::========================================================================================================================================

::  1st column = Generic Retail/OEM/MAK/GVLK Key
::  2nd column = Key Type
::  3rd column = WMI Edition ID
::  4th column = Version name incase same Edition ID is used in different OS versions with different key
::  Separator  = _

::  For Windows 10/11 editions, HWID key is listed where ever possible, in Server versions, KMS key is listed where ever possible.
::  For Windows, generic keys are mentioned till 22000 and for Server, generic keys are mentioned till 17763, later ones are extracted from the pkeyhelper.dll

:changeeditiondata

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
if %winbuild% GTR 17763 exit /b
) else (
if %winbuild% GEQ 22000 exit /b
)
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" (set Cor=Cor) else (set Cor=)

set h=
for %%# in (
XGVPP-NMH47-7TTHJ-W3FW7-8HV%h%2C__OEM:NONSLP_Enterprise
D6RD9-D4N8T-RT9QX-YW6YT-FCW%h%WJ______Retail_Starter
3V6Q6-NQXCX-V8YXR-9QCYV-QPF%h%CT__Volume:MAK_EnterpriseN
3NFXW-2T27M-2BDW6-4GHRV-68X%h%RX______Retail_StarterN
VK7JG-NPHTM-C97JM-9MPGT-3V6%h%6T______Retail_Professional
2B87N-8KFHP-DKV6R-Y2C8J-PKC%h%KT______Retail_ProfessionalN
4CPRK-NM3K3-X6XXQ-RXX86-WXC%h%HW______Retail_CoreN
N2434-X9D7W-8PF6X-8DV9T-8TY%h%MD______Retail_CoreCountrySpecific
BT79Q-G7N6G-PGBYW-4YWX6-6F4%h%BT______Retail_CoreSingleLanguage
YTMG3-N6DKC-DKB77-7M9GH-8HV%h%X7______Retail_Core
XKCNC-J26Q9-KFHD2-FKTHY-KD7%h%2Y__OEM:NONSLP_PPIPro
YNMGQ-8RYV3-4PGQ3-C8XTP-7CF%h%BY______Retail_Education
84NGF-MHBT6-FXBX8-QWJK7-DRR%h%8H______Retail_EducationN
KCNVH-YKWX8-GJJB9-H9FDT-6F7%h%W2__Volume:MAK_EnterpriseS_VB
43TBQ-NH92J-XKTM7-KT3KK-P39%h%PB__OEM:NONSLP_EnterpriseS_RS5
NK96Y-D9CD8-W44CQ-R8YTK-DYJ%h%WX__OEM:NONSLP_EnterpriseS_RS1
FWN7H-PF93Q-4GGP8-M8RF3-MDW%h%WW__OEM:NONSLP_EnterpriseS_TH
RQFNW-9TPM3-JQ73T-QV4VQ-DV9%h%PT__Volume:MAK_EnterpriseSN_VB
M33WV-NHY3C-R7FPM-BQGPT-239%h%PG__Volume:MAK_EnterpriseSN_RS5
2DBW3-N2PJG-MVHW3-G7TDK-9HK%h%R4__Volume:MAK_EnterpriseSN_RS1
NTX6B-BRYC2-K6786-F6MVQ-M7V%h%2X__Volume:MAK_EnterpriseSN_TH
G3KNM-CHG6T-R36X3-9QDG6-8M8%h%K9______Retail_ProfessionalSingleLanguage
HNGCC-Y38KG-QVK8D-WMWRK-X86%h%VK______Retail_ProfessionalCountrySpecific
DXG7C-N36C4-C4HTG-X4T3X-2YV%h%77______Retail_ProfessionalWorkstation
WYPNQ-8C467-V2W6J-TX4WX-WT2%h%RQ______Retail_ProfessionalWorkstationN
8PTT6-RNW4C-6V7J2-C2D3X-MHB%h%PB______Retail_ProfessionalEducation
GJTYN-HDMQY-FRR76-HVGC7-QPF%h%8P______Retail_ProfessionalEducationN
C4NTJ-CX6Q2-VXDMR-XVKGM-F9D%h%JC__Volume:MAK_EnterpriseG
46PN6-R9BK9-CVHKB-HWQ9V-MBJ%h%Y8__Volume:MAK_EnterpriseGN
NJCF7-PW8QT-3324D-688JX-2YV%h%66______Retail_ServerRdsh
XQQYW-NFFMW-XJPBH-K8732-CKF%h%FD______OEM:DM_IoTEnterprise
QPM6N-7J2WJ-P88HH-P3YRH-YY7%h%4H__OEM:NONSLP_IoTEnterpriseS
K9VKN-3BGWV-Y624W-MCRMQ-BHD%h%CD______Retail_CloudEditionN
KY7PN-VR6RX-83W6Y-6DDYQ-T6R%h%4W______Retail_CloudEdition
V3WVW-N2PV2-CGWC3-34QGF-VMJ%h%2C______Retail_Cloud
NH9J3-68WK7-6FB93-4K3DF-DJ4%h%F6______Retail_CloudN
2HN6V-HGTM8-6C97C-RK67V-JQP%h%FD______Retail_CloudE
WC2BQ-8NRM3-FDDYY-2BFGV-KHK%h%QY_Volume:GVLK_ServerStandard%Cor%_RS1
CB7KF-BWN84-R7R2Y-793K2-8XD%h%DG_Volume:GVLK_ServerDatacenter%Cor%_RS1
JCKRF-N37P4-C2D82-9YXRT-4M6%h%3B_Volume:GVLK_ServerSolution_RS1
QN4C6-GBJD2-FB422-GHWJK-GJG%h%2R_Volume:GVLK_ServerCloudStorage_RS1
VP34G-4NPPG-79JTQ-864T4-R3M%h%QX_Volume:GVLK_ServerAzureCor_RS1
9JQNQ-V8HQ6-PKB8H-GGHRY-R62%h%H6______Retail_ServerAzureNano_RS1
VN8D3-PR82H-DB6BJ-J9P4M-92F%h%6J______Retail_ServerStorageStandard_RS1
48TQX-NVK3R-D8QR3-GTHHM-8FH%h%XC______Retail_ServerStorageWorkgroup_RS1
2HXDN-KRXHB-GPYC7-YCKFJ-7FV%h%DG_Volume:GVLK_ServerDatacenterACor_RS3
PTXN8-JFHJM-4WC78-MPCBR-9W4%h%KR_Volume:GVLK_ServerStandardACor_RS3
) do (
for /f "tokens=1-4 delims=_" %%A in ("%%#") do if /i %targetedition%==%%C (

if not defined key (
set 4th=%%D
if not defined 4th (
set "key=%%A" & set "_chan=%%B"
) else (
echo "%branch%" | find /i "%%D" %nul1% && (set "key=%%A" & set "_chan=%%B")
)
)
)
)
exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:change_offedition

set "line=echo ___________________________________________________________________________________________"

cls
if not defined terminal mode 98, 30
title  Change Office Edition %masver%

echo:
echo Initializing...
echo:

if not exist %SysPath%\sppsvc.exe (
%eline%
echo [%SysPath%\sppsvc.exe] file is missing. Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_reflection
call :dk_ckeckwmic
call :dk_sppissue

for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if not "%%j"=="" (
set fullbuild=%%i.%%j
) else (
for /f "tokens=3" %%G in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR" %nul6%') do if not errorlevel 1 set /a "UBR=%%G"
for /f "skip=2 tokens=3,4 delims=. " %%G in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildLabEx') do (
if defined UBR (set "fullbuild=%%G.!UBR!") else (set "fullbuild=%%G.%%H")
)
)

::========================================================================================================================================

::  Check Windows Edition
::  This is just to ensure that SPP/WMI are functional

cls
set osedition=0
if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL) get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
%chkedi% do if not errorlevel 1 (call set "osedition=%%a")

if %osedition%==0 (
%eline%
echo Failed to detect OS Edition. Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

::  Check installed Office 16.0 C2R

set o16c2r=
set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (set o16c2r=1&set o16c2r_reg=%_86%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (set o16c2r=1&set o16c2r_reg=%_68%\ClickToRun)

if not defined o16c2r_reg (
%eline%
echo Office C2R 2016 or later is not installed, which is required for this script.
echo Download and install Office from below URL and try again.
echo:
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto dk_done
)

call :ch_getinfo

::========================================================================================================================================

::  Check minimum required details

if %verchk% LSS 9029 (
%eline%
echo Installed Office version is %_version%.
echo Minimum required version is 16.0.9029.2167
echo Aborting...
echo:
call :dk_color %Blue% "Download and install latest Office from below URL and try again."
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto dk_done
)

for %%A in (
_oArch
_updch
_lang
_clversion
_version
_oIds
_c2rXml
_c2rExe
_c2rCexe
_masterxml
) do (
if not defined %%A (
%eline%
echo Failed to find %%A. Aborting...
echo:
call :dk_color %Blue% "Download and install Office from below URL and try again."
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto dk_done
)
)

if %winbuild% LSS 10240 if defined ltscfound (
%eline%
echo Installed Office appears to be from the Volume channel %ltsc19%%ltsc21%%ltsc24%,
echo which is not officially supported on your Windows build version %winbuild%.
echo Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

set unsupbuild=
if %winbuild% LSS 10240 if %winbuild% GEQ 9200 if %verchk% GTR 16026 set unsupbuild=1
if %winbuild% LSS 9200 if %verchk% GTR 12527 set unsupbuild=1

if defined unsupbuild (
%eline%
echo Unsupported Office %verchk% is installed on your Windows build version %winbuild%.
echo Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

:oemenu

cls
set fixes=
if not defined terminal mode 76, 25
title  Change Office Edition %masver%
echo:
echo:
echo:
echo:
echo         ____________________________________________________________
echo:
echo                 [1] Change all editions
echo                 [2] Add edition
echo                 [3] Remove edition
echo:
echo                 [4] Add/Remove apps
echo                 ____________________________________________
echo:
echo                 [5] Change Office Update Channel
echo                 [0] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "           " %_Green% "Choose a menu option using your keyboard [1,2,3,4,5,0]"
choice /C:123450 /N
set _el=!errorlevel!
if !_el!==6  exit /b
if !_el!==5  goto :oe_changeupdchnl
if !_el!==4  goto :oe_editedition
if !_el!==3  goto :oe_removeedition
if !_el!==2  set change=0& goto :oe_edition
if !_el!==1  set change=1& goto :oe_edition
goto :oemenu

::========================================================================================================================================

:oe_edition

cls
call :oe_chkinternet
if not defined _int (
goto :oe_goback
)

cls
if not defined terminal mode 76, 25
if %change%==1 (
title  Change all editions %masver%
) else (
title  Add edition %masver%
)

echo:
echo:
echo:
echo:
echo                 O365/Mondo editions have the latest features.     
echo         ____________________________________________________________
echo:
echo                 [1] Office Suites     - Retail
echo                 [2] Office Suites     - Volume
echo                 [3] Office SingleApps - Retail
echo                 [4] Office SingleApps - Volume
echo                 ____________________________________________
echo:
echo                 [0] Go Back
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "            " %_Green% "Choose a menu option using your keyboard [1,2,3,4,0]"
choice /C:12340 /N
set _el=!errorlevel!
if !_el!==5  goto :oemenu
if !_el!==4  set list=SingleApps_Volume&goto :oe_editionchangepre
if !_el!==3  set list=SingleApps_Retail&goto :oe_editionchangepre
if !_el!==2  set list=Suites_Volume&goto :oe_editionchangepre
if !_el!==1  set list=Suites_Retail&goto :oe_editionchangepre
goto :oe_edition

::========================================================================================================================================

:oe_editionchangepre

cls
call :ch_getinfo
call :oe_tempcleanup
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':getlist\:.*';iex ($f[1])"

:oe_editionchange

cls
if not defined terminal (
mode 98, 45
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=44;$B.Height=100;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)

if not exist %SystemRoot%\Temp\%list%.txt (
%eline%
echo Failed to generate available editions list.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto :oe_goback
)

set inpt=
set counter=0
set verified=0
set _notfound=
set targetedition=

%line%
echo:
call :dk_color %Gray% "Installed Office editions: %_oIds%"
call :dk_color %Gray% "You can select one of the following Office Editions."
if %winbuild% LSS 10240 (
echo Unsupported products such as 2019/2021/2024 are excluded from this list.
) else (
for %%# in (2019 2021 2024) do (
find /i "%%#" "%SystemRoot%\Temp\%list%.txt" %nul1% || (
if defined _notfound (set _notfound=%%#, !_notfound!) else (set _notfound=%%#)
)
)
if defined _notfound call :dk_color %Gray% "Office !_notfound! is not in this list because old version [%_version%] of Office is installed."
)
%line%
echo:

for /f "usebackq delims=" %%A in (%SystemRoot%\Temp\%list%.txt) do (
set /a counter+=1
if !counter! LSS 10 (
echo [!counter!]  %%A
) else (
echo [!counter!] %%A
)
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  Go Back
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto :oe_editionchange
if "%inpt%"=="0" (call :oe_tempcleanup & goto :oe_edition)
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto :oe_editionchange

::========================================================================================================================================

::  Set app exclusions

:oe_excludeappspre

cls
set suites=
echo %list% | find /i "Suites" %nul1% && (
set suites=1
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':getappnames\:.*';iex ($f[1])"
if not exist %SystemRoot%\Temp\getAppIds.txt (
%eline%
echo Failed to generate available apps list.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto :oe_goback
)
)

for %%# in (
Access
Excel
Lync
OneNote
Outlook
PowerPoint
Project
Publisher
Visio
Word
) do (
if defined suites (
find /i "%%#" "%SystemRoot%\Temp\getAppIds.txt" %nul1% && (set %%#_st=On) || (set %%#_st=)
) else (
set %%#_st=
)
)

if defined Lync_st set Lync_st=Off
set OneDrive_st=Off
if defined suites (set Teams_st=Off) else (set Teams_st=)

:oe_excludeapps

cls
if not defined terminal mode 98, 32

%line%
echo:
call :dk_color %Gray% "Target edition: %targetedition%"
call :dk_color %Gray% "You can exclude the below apps from installation."
%line%
if defined suites echo:
if defined Access_st     echo [A] Access           : %Access_st%
if defined Excel_st      echo [E] Excel            : %Excel_st%
if defined OneNote_st    echo [N] OneNote          : %OneNote_st%
if defined Outlook_st    echo [O] Outlook          : %Outlook_st%
if defined PowerPoint_st echo [P] PowerPoint       : %PowerPoint_st%
if defined Project_st    echo [J] Project          : %Project_st%
if defined Publisher_st  echo [R] Publisher        : %Publisher_st%
if defined Visio_st      echo [V] Visio            : %Visio_st%
if defined Word_st       echo [W] Word             : %Word_st%
echo:
if defined Lync_st       echo [L] SkypeForBusiness : %Lync_st%
if defined OneDrive_st   echo [D] OneDrive         : %OneDrive_st%
if defined Teams_st      echo [T] Teams            : %Teams_st%
%line%
echo:
echo [1] Continue
echo [0] Go Back
%line%
echo:
call :dk_color %_Green% "Choose a menu option using your keyboard:"
choice /C:AENOPJRVWLDT10 /N
set _el=!errorlevel!
if !_el!==14 goto :oemenu
if !_el!==13 call :excludelist & goto :oe_editionchangefinal
if !_el!==12 if defined Teams_st      (if "%Teams_st%"=="Off"      (set Teams_st=ON)      else (set Teams_st=Off))
if !_el!==11 if defined OneDrive_st   (if "%OneDrive_st%"=="Off"   (set OneDrive_st=ON)   else (set OneDrive_st=Off))
if !_el!==10 if defined Lync_st       (if "%Lync_st%"=="Off"       (set Lync_st=ON)       else (set Lync_st=Off))
if !_el!==9  if defined Word_st       (if "%Word_st%"=="Off"       (set Word_st=ON)       else (set Word_st=Off))
if !_el!==8  if defined Visio_st      (if "%Visio_st%"=="Off"      (set Visio_st=ON)      else (set Visio_st=Off))
if !_el!==7  if defined Publisher_st  (if "%Publisher_st%"=="Off"  (set Publisher_st=ON)  else (set Publisher_st=Off))
if !_el!==6  if defined Project_st    (if "%Project_st%"=="Off"    (set Project_st=ON)    else (set Project_st=Off))
if !_el!==5  if defined PowerPoint_st (if "%PowerPoint_st%"=="Off" (set PowerPoint_st=ON) else (set PowerPoint_st=Off))
if !_el!==4  if defined Outlook_st    (if "%Outlook_st%"=="Off"    (set Outlook_st=ON)    else (set Outlook_st=Off))
if !_el!==3  if defined OneNote_st    (if "%OneNote_st%"=="Off"    (set OneNote_st=ON)    else (set OneNote_st=Off))
if !_el!==2  if defined Excel_st      (if "%Excel_st%"=="Off"      (set Excel_st=ON)      else (set Excel_st=Off))
if !_el!==1  if defined Access_st     (if "%Access_st%"=="Off"     (set Access_st=ON)     else (set Access_st=Off))
goto :oe_excludeapps

:excludelist

set excludelist=
for %%# in (
access
excel
onenote
outlook
powerpoint
project
publisher
visio
word
lync
onedrive
teams
) do (
if /i "!%%#_st!"=="Off" if defined excludelist (set excludelist=!excludelist!,%%#) else (set excludelist=,%%#)
)
exit /b

::========================================================================================================================================

::  Final command to change/add edition

:oe_editionchangefinal

cls
if not defined terminal mode 105, 32

::  Check for Project and Visio with unsupported language

set projvis=
set langmatched=
echo: %Project_st% %Visio_st% | find /i "ON" %nul% && set projvis=1
echo: %targetedition% | findstr /i "Project Visio" %nul% && set projvis=1

if defined projvis (
for %%# in (
ar-sa
cs-cz
da-dk
de-de
el-gr
en-us
es-es
fi-fi
fr-fr
he-il
hu-hu
it-it
ja-jp
ko-kr
nb-no
nl-nl
pl-pl
pt-br
pt-pt
ro-ro
ru-ru
sk-sk
sl-si
sv-se
tr-tr
uk-ua
zh-cn
zh-tw
) do (
if /i "%_lang%"=="%%#" set langmatched=1
)
if not defined langmatched (
%eline%
echo %_lang% language is not available for Project/Visio apps.
echo:
call :dk_color %Blue% "Install Office in the supported language for Project/Visio from the below URL."
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto :oe_goback
)
)

::  Thanks to @abbodi1406 for first discovering OfficeClickToRun.exe uses
::  Thanks to @may for the suggestion to use it to change edition with CDN as a source
::  OfficeClickToRun.exe with productstoadd method is used here to add editions
::  It uses delta updates, meaning that since it's using same installed build, it will consume very less Internet

set "c2rcommand="%_c2rExe%" platform=%_oArch% culture=%_lang% productstoadd=%targetedition%.16_%_lang%_x-none cdnbaseurl.16=http://officecdn.microsoft.com/pr/%_updch% baseurl.16=http://officecdn.microsoft.com/pr/%_updch% version.16=%_version% mediatype.16=CDN sourcetype.16=CDN deliverymechanism=%_updch% %targetedition%.excludedapps.16=groove%excludelist% flt.useteamsaddon=disabled flt.usebingaddononinstall=disabled flt.usebingaddononupdate=disabled"

if %change%==1 (
set "c2rcommand=!c2rcommand! productstoremove=AllProducts"
)

echo:
echo Running the below command, please wait...
echo:
echo %c2rcommand%
%c2rcommand%
set errorcode=%errorlevel%
timeout /t 10 %nul%

echo:
if %errorcode% EQU 0 (
call :dk_color %Gray% "Now run the Office activation option from the main menu."
) else (
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

call :oe_tempcleanup
goto :oe_goback

::========================================================================================================================================

::  Edit Office edition

:oe_editedition

cls
title  Add/Remove Apps %masver%

call :oe_chkinternet
if not defined _int (
goto :oe_goback
)

set change=0
call :ch_getinfo
cls

if not defined terminal (
mode 98, 35
)

set inpt=
set counter=0
set verified=0
set targetedition=

%line%
echo:
call :dk_color %Gray% "You can edit [add/remove apps] one of the following Office editions."
%line%
echo:

for %%A in (%_oIds%) do (
set /a counter+=1
echo [!counter!] %%A
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  Go Back
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto :oe_editedition
if "%inpt%"=="0" goto :oemenu
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto :oe_editedition

::===============

cls
if not defined terminal mode 98, 32

echo %targetedition% | findstr /i "Access Excel OneNote Outlook PowerPoint Project Publisher Skype Visio Word" %nul% && (set list=SingleApps) || (set list=Suites)
goto :oe_excludeappspre

::========================================================================================================================================

::  Remove Office editions

:oe_removeedition

title  Remove Office editions %masver%

call :ch_getinfo

cls
if not defined terminal (
mode 98, 35
)

set counter=0
for %%A in (%_oIds%) do (set /a counter+=1)

if !counter! LEQ 1 (
echo:
echo Only "%_oIds%" product is installed.
echo This option is available only when multiple products are installed.
goto :oe_goback
)

::===============

set inpt=
set counter=0
set verified=0
set targetedition=

%line%
echo:
call :dk_color %Gray% "You can uninstall one of the following Office editions."
%line%
echo:

for %%A in (%_oIds%) do (
set /a counter+=1
echo [!counter!] %%A
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  Go Back
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto :oe_removeedition
if "%inpt%"=="0" goto :oemenu
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto :oe_removeedition

::===============

cls
if not defined terminal mode 105, 32

set _lang=
echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=10) || (set _tok=9)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs\%_actconfig%\%targetedition%.16" /f "-" /k ^| findstr /i ".*16\\.*-.*"') do (
if defined _lang (set "_lang=!_lang!_%%a") else (set "_lang=_%%a")
)

set "c2rcommand="%_c2rExe%" platform=%_oArch% productstoremove=%targetedition%.16%_lang%"

echo:
echo Running the below command, please wait...
echo:
echo %c2rcommand%
%c2rcommand%

if %errorlevel% NEQ 0 (
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

goto :oe_goback

::========================================================================================================================================

::  Change Office update channel

:oe_changeupdchnl

title  Change Office update channel %masver%
call :ch_getinfo

cls
if not defined terminal (
mode 98, 33
)

call :oe_chkinternet
if not defined _int (
goto :oe_goback
)

if %winbuild% LSS 10240 (
echo %_oIds% | findstr "2019 2021 2024" %nul% && (
%eline%
echo Installed Office editions: %_oIds%
echo Unsupported Office edition is installed on your Windows build version %winbuild%.
goto :oe_goback
)
)

::===============

set inpt=
set counter=0
set verified=0
set targetFFN=
set targetchannel=

%line%
echo:
call :dk_color %Gray% "Installed update channel: %_AudienceData%, %_version%, Client: %_clversion%"
call :dk_color %Gray% "Unsupported update channels are excluded from this list."
%line%
echo:

for %%# in (
"5440FD1F-7ECB-4221-8110-145EFAA6372F_Insider Fast [Beta]  -    Insiders::DevMain"
"64256AFE-F5D9-4F86-8936-8840A6A4F5BE_Monthly Preview      -    Insiders::CC"
"492350F6-3A01-4F97-B9C0-C7C6DDF67D60_Monthly [Current]    -  Production::CC"
"55336B82-A18D-4DD6-B5F6-9E5095C314A6_Monthly Enterprise   -  Production::MEC"
"B8F9B850-328D-4355-9145-C59439A0C4CF_Semi Annual Preview  -    Insiders::FRDC"
"7FFBC6BF-BC32-4F92-8982-F9DD17FD3114_Semi Annual          -  Production::DC"
"EA4A4090-DE26-49D7-93C1-91BFF9E53FC3_DevMain Channel      -     Dogfood::DevMain"
"B61285DD-D9F7-41F2-9757-8F61CBA4E9C8_Microsoft Elite      -   Microsoft::DevMain"
"F2E724C1-748F-4B47-8FB8-8E0D210E9208_Perpetual2019 VL     -  Production::LTSC"
"1D2D2EA6-1680-4C56-AC58-A441C8C24FF9_Microsoft2019 VL     -   Microsoft::LTSC"
"5030841D-C919-4594-8D2D-84AE4F96E58E_Perpetual2021 VL     -  Production::LTSC2021"
"86752282-5841-4120-AC80-DB03AE6B5FDB_Microsoft2021 VL     -   Microsoft::LTSC2021"
"7983BAC0-E531-40CF-BE00-FD24FE66619C_Perpetual2024 VL     -  Production::LTSC2024"
"C02D8FE6-5242-4DA8-972F-82EE55E00671_Microsoft2024 VL     -   Microsoft::LTSC2024"
) do (
for /f "tokens=1-2 delims=_" %%A in ("%%~#") do (
set supported=
if %winbuild% LSS 10240 (echo %%B | findstr /i "LTSC DevMain" %nul% || set supported=1) else (set supported=1)
if %winbuild% GEQ 10240 (
if defined ltsc19 echo %%B | find /i "2019 VL" %nul% || set supported=
if defined ltsc21 echo %%B | find /i "2021 VL" %nul% || set supported=
if defined ltsc24 echo %%B | find /i "2024 VL" %nul% || set supported=
if not defined ltscfound echo %%B | find /i "LTSC" %nul% && set supported=
)
if defined supported (
set /a counter+=1
if !counter! LSS 10 (
echo [!counter!]  %%B
) else (
echo [!counter!] %%B
)
set targetFFN!counter!=%%A
set targetchannel!counter!=%%B
)
)
)

%line%
echo:
echo [R]  Learn about update channels
echo [0]  Go back
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto :oe_changeupdchnl
if "%inpt%"=="0" goto :oemenu
if /i "%inpt%"=="R" start https://learn.microsoft.com/microsoft-365-apps/updates/overview-update-channels & goto :oe_changeupdchnl
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetFFN=!targetFFN%inpt%!
set targetchannel=!targetchannel%inpt%!
if %verified%==0 goto :oe_changeupdchnl

::=======================

cls
if not defined terminal mode 105, 32

::  Get build number for the target FFN, using build number with OfficeC2RClient.exe command to trigger updates provides accurate results

set build=
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':getbuild\:.*';iex ($f[1])" %nul6%') do (set build=%%a)
echo "%build%" | find /i "16." %nul% || set build=

::  Cleanup Office update related registries, thanks to @abbodi1406
::  https://techcommunity.microsoft.com/t5/office-365-blog/how-to-manage-office-365-proplus-channels-for-it-pros/ba-p/795813
::  https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels#considerations-when-changing-channels

echo:
for /f "tokens=1 delims=-" %%A in ("%targetchannel%") do (echo Target update channel: %%A)
echo:
echo Cleaning Office update registry keys...
echo Adding new update channel to registry keys...

%nul% reg add %o16c2r_reg%\Configuration /v CDNBaseUrl /t REG_SZ /d "https://officecdn.microsoft.com/pr/%targetFFN%" /f
%nul% reg add %o16c2r_reg%\Configuration /v UpdateChannel /t REG_SZ /d "https://officecdn.microsoft.com/pr/%targetFFN%" /f
%nul% reg add %o16c2r_reg%\Configuration /v UpdateChannelChanged /t REG_SZ /d "True" /f
%nul% reg delete %o16c2r_reg%\Configuration /v UnmanagedUpdateURL /f
%nul% reg delete %o16c2r_reg%\Configuration /v UpdateUrl /f
%nul% reg delete %o16c2r_reg%\Configuration /v UpdatePath /f
%nul% reg delete %o16c2r_reg%\Configuration /v UpdateToVersion /f
%nul% reg delete %o16c2r_reg%\Updates /v UpdateToVersion /f
%nul% reg delete HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate /f
%nul% reg delete HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate /f /reg:32
%nul% reg delete HKCU\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate /f
%nul% reg delete HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate /f
%nul% reg delete HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate /f /reg:32
%nul% reg delete HKCU\Software\Policies\Microsoft\cloud\office\16.0\Common\officeupdate /f

if not defined build (
if %winbuild% GEQ 9200 call :dk_color %Gray% "Failed to detect build number for the target FFN."
set "updcommand="%_c2rCexe%" /update user"
) else (
set "updcommand="%_c2rCexe%" /update user updatetoversion=%build%"
)
echo Running the below command to trigger updates...
echo:
echo %updcommand%
%updcommand%
echo:
echo Help - %mas%troubleshoot
goto :oe_goback

::========================================================================================================================================

:oe_goback

call :oe_tempcleanup

echo:
if defined fixes (
call :dk_color %White% "Follow ALL the ABOVE blue lines.   "
call :dk_color2 %Blue% "Press [1] to Open Support Webpage " %Gray% " Press [0] to Ignore"
choice /C:10 /N
if !errorlevel!==1 (for %%# in (%fixes%) do (start %%#))
)

if defined terminal (
call :dk_color %_Yellow% "Press [0] key to go back..."
choice /c 0 /n
) else (
call :dk_color %_Yellow% "Press any key to go back..."
pause %nul1%
)
goto :oemenu

::========================================================================================================================================

:oe_tempcleanup

del /f /q %SystemRoot%\Temp\SingleApps_Volume.txt %nul%
del /f /q %SystemRoot%\Temp\SingleApps_Retail.txt %nul%
del /f /q %SystemRoot%\Temp\Suites_Volume.txt %nul%
del /f /q %SystemRoot%\Temp\Suites_Retail.txt %nul%
del /f /q %SystemRoot%\Temp\getAppIds.txt %nul%
exit /b

::========================================================================================================================================

::  Fetch required info

:ch_getinfo

set _oRoot=
set _oArch=
set _updch=
set _oIds=
set _lang=
set _cfolder=
set _version=
set _clversion=
set _AudienceData=
set _actconfig=
set _c2rXml=
set _c2rExe=
set _c2rCexe=
set _masterxml=
set ltsc19=
set ltsc21=
set ltsc24=
set ltscfound=

for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ClientFolder" %nul6%') do (set "_cfolder=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v AudienceId" %nul6%') do (set "_updch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ClientCulture" %nul6%') do (set "_lang=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ClientVersionToReport" %nul6%') do (set "_clversion=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v AudienceData" %nul6%') do (set "_AudienceData=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\ProductReleaseIDs /v ActiveConfiguration" %nul6%') do (set "_actconfig=%%b")

echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=9) || (set _tok=8)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs\%_actconfig%" /f ".16" /k %nul6% ^| findstr /i "Retail Volume"') do (
if defined _oIds (set "_oIds=!_oIds! %%a") else (set "_oIds=%%a")
)
set _oIds=%_oIds:.16=%

set verchk=0
for /f "tokens=3 delims=." %%a in ("%_version%") do set "verchk=%%a"

if exist "%_oRoot%\Licenses16\c2rpridslicensefiles_auto.xml" set "_c2rXml=%_oRoot%\Licenses16\c2rpridslicensefiles_auto.xml"

if exist "%ProgramData%\Microsoft\ClickToRun\ProductReleases\%_actconfig%\x-none.16\MasterDescriptor.x-none.xml" (
set "_masterxml=%ProgramData%\Microsoft\ClickToRun\ProductReleases\%_actconfig%\x-none.16\MasterDescriptor.x-none.xml"
)

if exist "%_cfolder%\OfficeClickToRun.exe" (
set "_c2rExe=%_cfolder%\OfficeClickToRun.exe"
)

if exist "%_cfolder%\OfficeC2RClient.exe" (
set "_c2rCexe=%_cfolder%\OfficeC2RClient.exe"
)

set "audidata4=%_AudienceData:~-4%"

if /i "%audidata4%"=="LTSC" set ltsc19=LTSC
echo %_clversion% %_version% | findstr "16.0.103 16.0.104 16.0.105" %nul% && set ltsc19=LTSC

if /i "%audidata4%"=="2021" set ltsc21=LTSC2021
echo %_clversion% %_version% | findstr "16.0.14332" %nul% && set ltsc21=LTSC2021

if /i "%audidata4%"=="2024" set ltsc24=LTSC2024
::  LTSC 2024 build is not fixed yet

if not "%ltsc19%%ltsc21%%ltsc24%"=="" set ltscfound=1

exit /b

::========================================================================================================================================

::  Check Internet connection

:oe_chkinternet

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not "%%#"=="" set _int=1)
)

if not defined _int (
%psc% "If([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet){Exit 0}Else{Exit 1}"
if !errorlevel!==0 (set _int=1)
)

if not defined _int (
%eline%
call :dk_color %Red% "Internet is not connected."
call :dk_color %Blue% "Internet is required for this operation."
)
exit /b

::========================================================================================================================================

::  Get available build number for a FFN

:getbuild:
$Tls12 = [Enum]::ToObject([System.Net.SecurityProtocolType], 3072)
[System.Net.ServicePointManager]::SecurityProtocol = $Tls12

$FFN = $env:targetFFN
$windowsBuild = [System.Environment]::OSVersion.Version.Build

$baseUrl = "https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData?audienceFFN=$FFN"
$url = if ($windowsBuild -lt 9200) { "$baseUrl&osver=Client|6.1" } elseif ($windowsBuild -lt 10240) { "$baseUrl&osver=Client|6.3" } else { $baseUrl }

$response = if ($windowsBuild -ge 9200) { irm -Uri $url -Method Get } else { (New-Object System.Net.WebClient).DownloadString($url) }

if ($windowsBuild -lt 9200) {
    if ($response -match '"AvailableBuild"\s*:\s*"([^"]+)"') { Write-Host $matches[1] }
} else {
    Write-Host $response.AvailableBuild
}
:getbuild:

::========================================================================================================================================

::  Get available edition list from c2rpridslicensefiles_auto.xml
::  and filter the list using MasterDescriptor.x-none.xml
::  and exclude unsupported products on Windows 7/8/8.1

:getlist:
$xmlPath1 = $env:_c2rXml
$xmlPath2 = $env:_masterxml
$outputDir = $env:SystemRoot + "\Temp\"
$buildNumber = [System.Environment]::OSVersion.Version.Build
$excludedKeywords = @("2019", "2021", "2024")
$productReleaseIds = @()

if (Test-Path $xmlPath1) {
    $xml1 = New-Object -TypeName System.Xml.XmlDocument
    $xml1.Load($xmlPath1)
    foreach ($node in $xml1.SelectNodes("//ProductReleaseId")) {
        $id = $node.GetAttribute("id")
        $exclude = $false
        if ($buildNumber -lt 10240) {
            foreach ($keyword in $excludedKeywords) {
                if ($id -match $keyword) { $exclude = $true; break }
            }
        }
        if ($id -ne "CommonLicenseFiles" -and -not $exclude) { $productReleaseIds += $id }
    }
}

$categories = @{
    "Suites_Retail" = @(); "Suites_Volume" = @()
    "SingleApps_Retail" = @(); "SingleApps_Volume" = @()
}

foreach ($id in $productReleaseIds) {
    $category = if ($id -match "Retail") { "Retail" } else { "Volume" }
    $categories["SingleApps_$category"] += $id
}

if (Test-Path $xmlPath2) {
    $xml2 = New-Object -TypeName System.Xml.XmlDocument
    $xml2.Load($xmlPath2)
    foreach ($sku in $xml2.SelectNodes("//SKU")) {
        $skuId = $sku.GetAttribute("ID")
        if ($productReleaseIds -contains $skuId) {
            $appIds = $sku.SelectNodes("Apps/App") | ForEach-Object { $_.GetAttribute("id") }
            if ($appIds -contains "Excel" -and $appIds -contains "Word") {
                $category = if ($skuId -match "Retail") { "Retail" } else { "Volume" }
                $categories["Suites_$category"] += $skuId
                $categories["SingleApps_$category"] = $categories["SingleApps_$category"] | Where-Object { $_ -ne $skuId }
            }
        }
    }
}

foreach ($section in $categories.Keys) {
    $filePath = Join-Path -Path $outputDir -ChildPath "$section.txt"
    $ids = $categories[$section]
    if ($ids.Count -gt 0) { $ids | Out-File -FilePath $filePath -Encoding ASCII }
}
:getlist:

::========================================================================================================================================

::  Get App list for a specific product ID using MasterDescriptor.x-none.xml

:getappnames:
$xmlPath = $env:_masterxml
$targetSkuId = $env:targetedition
$outputDir = $env:SystemRoot + "\Temp\"
$outputFile = Join-Path -Path $outputDir -ChildPath "getAppIds.txt"
$excludeIds = @("shared", "PowerPivot", "PowerView", "MondoOnly", "OSM", "OSMUX", "Groove", "DCF")

$xml = New-Object -TypeName System.Xml.XmlDocument
$xml.Load($xmlPath)

$appIdsList = @()
$skuNodes = $xml.SelectNodes("//SKU[@ID='$targetSkuId']")

foreach ($skuNode in $skuNodes) {
    foreach ($app in $skuNode.SelectNodes("Apps/App")) {
        $appId = $app.GetAttribute("id")
        if ($excludeIds -notcontains $appId) {
            $appIdsList += $appId
        }
    }
}

if ($appIdsList.Count -gt 0) {
    $appIdsList | Out-File -FilePath $outputFile -Encoding ASCII
}
:getappnames:

::========================================================================================================================================
::
:: Leave empty line below
