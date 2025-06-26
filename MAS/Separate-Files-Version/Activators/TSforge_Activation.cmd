@set masver=3.4
@echo off



::============================================================================
::
::   Homepage: mass grave[.]dev
::      Email: mas.help@outlook.com
::
::============================================================================



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

::  Choose activation method:
::  In builds 19041 and later, the script will auto select StaticCID (requires internet). If no internet is detected, it will then auto select the KMS4k method. For builds lower than 19041, the script will auto select ZeroCID.
::  To change the activation method, run the script with the parameters "/Z-SCID", "/Z-ZCID", or "/Z-KMS4k", or modify the option from Auto to SCID, ZCID, or KMS4k in the line below.
set _actmethod=Auto

::  Debug Mode:
::  To run the script in debug mode, change 0 to any parameter above that you want to run, in below line
set "_debug=0"

::  Script will run in unattended mode if parameters are used OR value is changed in above lines.
::  If multiple options are selected then script will only pick one from the advanced option.



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

::  Debug code

if "%_debug%" EQU "0" (
set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"
goto :_debug
)

set "nul1="
set "nul2="
set "nul6="
set "nul="

@echo on
@prompt $G
@call :_debug "%_debug%" >"%~dp0_tmp.log" 2>&1
@cmd /u /c type "%~dp0_tmp.log">"%~dp0_Debug.log"
@del "%~dp0_tmp.log"
@echo off
@exit /b

:_debug

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
if /i "%%A"=="/Z-SCID"                 (set _actmethod=SCID)
if /i "%%A"=="/Z-ZCID"                 (set _actmethod=ZCID)
if /i "%%A"=="/Z-KMS4k"                (set _actmethod=KMS4k)
)

if not defined tsids set _actman=0
for %%A in (%_actwin% %_actesu% %_actoff% %_actprojvis% %_actwinesuoff% %_actwinhost% %_actoffhost% %_actappx% %_actman% %_resall%) do (if "%%A"=="1" set _unattended=1)
if /i not %_actmethod%==Auto set _unattended=1

::========================================================================================================================================

call :dk_setvar

if %winbuild% EQU 1 (
%eline%
echo Failed to detect Windows build number.
echo:
setlocal EnableDelayedExpansion
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
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
goto dk_done
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
goto dk_done
)

if %winbuild% LSS 7600 (
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Install %nul2% | find /i "0x1" %nul1% || (
%eline%
echo .NET 3.5 Framework is not installed in your system.
echo Install it using the following URL.
echo:
echo https://www.microsoft.com/en-us/download/details.aspx?id=25150
if %_unattended%==0 start https://www.microsoft.com/en-us/download/details.aspx?id=25150
goto dk_done
)
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

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo Right click on this script and select 'Run as administrator'.
goto dk_done
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
goto dk_done
)

REM check Powershell core version

cmd /c "%psc% "$PSVersionTable.PSEdition"" | find /i "Core" %nul1% && (
echo Windows Powershell is needed for MAS but it seems to be replaced with Powershell core. Aborting...
goto dk_done
)

REM check for Mal-ware that may cause issues with Powershell

for /r "%ProgramFiles%\" %%f in (secureboot.exe) do if exist "%%f" (
echo "%%f"
echo Mal%blank%ware found, PowerShell is not working properly.
echo:
set fixes=%fixes% %mas%remove_mal%w%ware
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%remove_mal%w%ware"
goto dk_done
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
echo               [2] Activate - ESU
echo               [3] Activate - Office [All]
echo               [4] Activate - Office [Project/Visio]
echo               [5] Activate - All
echo               _______________________________________________  
echo: 
echo                   Advanced Options:
echo:
echo               [A] Activate - Windows %KS% Host
echo               [B] Activate - Office %KS% Host
echo               [C] Activate - Windows 8/8.1 APPX Sideloading
echo               [D] Activate - Manually Select Products
if defined _vis (
echo               [E] Reset    - Rearm/Timers
) else (
echo               [E] Reset    - Rearm/Timers/Tamper/Lock
)
echo               [F] Change   - Activation Method [%_actmethod%]
echo               _______________________________________________       
echo:
echo               [6] Remove TSforge Activation
echo               [7] Download Office
echo               [0] %_exitmsg%
echo        ______________________________________________________________
echo:
call :dk_color2 %_White% "            " %_Green% "Choose a menu option using your keyboard..."
choice /C:12345ABCDEF670 /N
set _el=!errorlevel!

if !_el!==14 exit /b
if !_el!==13 start %mas%genuine-installation-media & goto :ts_menu
if !_el!==12 call :ts_remove & cls & goto :ts_menu
if !_el!==11 goto :ts_changemethod
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

goto :ts_start

::========================================================================================================================================

:ts_changemethod

cls
if not defined terminal mode 76, 36

echo:
echo:
echo:
echo        ______________________________________________________________
echo: 
call :dk_color2 %_White% "             [1] " %_Green% "Auto"
echo                  Builds ^>= 19041 - StaticCID (KMS4k if offline)
echo                  Builds ^<  19041 - ZeroCID
echo              __________________________________________________
echo: 
echo              [2] StaticCID
echo                  Needs Internet
echo                  Not for Windows 7 or older
echo              __________________________________________________
echo:
echo              [3] ZeroCID
echo                  Works reliably on builds below 19041
echo                  May break on builds between 19041-26100
echo                  Does not work on builds above 26100.4188
echo              __________________________________________________
echo:
echo              [4] KMS4k
echo                  Volume licenses only
echo                  Activates for 4000+ years
echo              __________________________________________________
echo:
echo              [5] Learn More
echo              [0] %_exitmsg%
echo        ______________________________________________________________
echo:
call :dk_color2 %_White% "            " %_Green% "Choose a menu option using your keyboard..."
choice /C:123450 /N
set _el=!errorlevel!

if !_el!==6 goto :ts_menu
if !_el!==5  cls & start %mas%tsforge &goto :ts_menu
if !_el!==4  cls & set "_actmethod=KMS4k" & goto :ts_menu
if !_el!==3  cls & set "_actmethod=ZCID"  & goto :ts_menu
if !_el!==2  cls & set "_actmethod=SCID"  & goto :ts_menu
if !_el!==1  cls & set "_actmethod=Auto"  & goto :ts_menu
goto :ts_changemethod

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

if not exist %SysPath%\%_slexe% (
%eline%
echo [%SysPath%\%_slexe%] file is missing, aborting...
echo:
if not defined results (
call :dk_color %Blue% "Go back to Main Menu, select Troubleshoot and run DISM Restore and SFC Scan options."
call :dk_color %Blue% "After that, restart system and try activation again."
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
)
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
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)
)

if %winbuild% LSS 9200 if exist "%SysPath%\wlms\wlms.exe" (
sc query wlms | find /i "RUNNING" %nul% && (
sc stop %_slser% %nul%
if !errorlevel! EQU 1051 (
%eline%
echo Evaluation WLMS service is running, %_slser% service can not be stopped. Aborting...
echo Install Non-Eval version for Windows build %winbuild%.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
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

if /i %_actmethod%==SCID set tsmethod=StaticCID
if /i %_actmethod%==ZCID set tsmethod=ZeroCID
if /i %_actmethod%==KMS4k set tsmethod=KMS4k

if /i %_actmethod%==Auto (
if %winbuild% GEQ 19041 (
set tsmethod=StaticCID
) else (
set tsmethod=ZeroCID
)
)

if %winbuild% LSS 9200 if /i %tsmethod%==StaticCID (
%eline%
echo StaticCID method is supported only on Windows 8 and later.
goto dk_done
)

::========================================================================================================================================

::  Check Internet connection

if /i %tsmethod%==StaticCID (
set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not "%%#"=="" set _int=1)
)

if not defined _int (
%psc% "If([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet){Exit 0}Else{Exit 1}"
if !errorlevel!==0 (set _int=1&set ping_f= But Ping Failed)
)

if defined _int (
echo Checking Internet Connection            [Connected!ping_f!]
) else (
if /i %_actmethod%==Auto if not %_actman%==1 set tsmethod=KMS4k
if /i !tsmethod!==KMS4k (
call :dk_color %Gray% "Checking Internet Connection            [Not Connected]"
call :dk_color %Blue% "Switching To KMS4k Activation Method because Internet is needed for StaticCID method."
) else (
set error=1
call :dk_color %Red% "Checking Internet Connection            [Not Connected]"
call :dk_color %Blue% "Internet is required for TSforge StaticCID option."
)
echo:
)
)

::========================================================================================================================================

echo Initiating Diagnostic Tests...

set "_serv=%_slser% Winmgmt"

::  Software Protection
::  Windows Management Instrumentation

call :dk_errorcheck

call :ts_getedition
if not defined tsedition (
call :dk_color %Red% "Checking Windows Edition ID             [Not found in installed licenses, aborting...]"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
goto :dk_done
)

if /i !tsmethod!==KMS4k (
call :_taskclear-cache
echo Clearing %KS% Cache                      [Successful]
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

set /a UBR=0
if %winbuild% EQU 26100 (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR %nul6%') do if not errorlevel 1 set /a UBR=%%b
if !UBR! LSS 4188 (set dontcheckact=1)
)

if not defined dontcheckact call :ts_checkwinperm
if defined _perm (
call :dk_color %Gray% "Checking OS Activation                  [Windows is already permanently activated]"
goto :ts_esu
)

if %winbuild% LSS 9200 if /i %tsmethod%==KMS4k goto :ts_oldks
if defined _vis goto :ts_winvista

set tempid=
if /i %tsmethod%==KMS4k (set keytype=ks) else (set keytype=zero)
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':wintsid\:.*';iex ($f[1])" %nul6%') do (
echo "%%a" | findstr /r ".*-.*-.*-.*-.*" %nul1% && (set tsids=!tsids! %%a& set tempid=%%a)
)

if defined tempid (
echo Checking Activation ID                  [%tempid%] [%tsedition%]
) else (
call :dk_color %Red% "Checking Activation ID                  [Not Found] [%tsedition%] [%osSKU%]"
set error=1
if /i %tsmethod%==KMS4k (
if /i %_actmethod%==Auto (
call :dk_color %Blue% "Connect to the Internet and try again. Script will use the StaticCID activation method."
) else (
call :dk_color %Blue% "Use non-KMS4K activation options from the previous menu."
)
)
goto :ts_esu
)

if defined winsub (
call :dk_color %Blue% "Windows Subscription [SKU ID-%slcSKU%] found. Script will activate base edition [SKU ID-%regSKU%]."
echo:
)

goto :ts_esu

::========================================================================================================================================

:ts_oldks

::  KMS keys for KMS4k method because TSforge cannot install KMS key on Windows Vista and 7

::  1st column = Activation ID
::  2nd column = Generic key
::  3rd column = Edition ID
::  Separator  = _

set f=
set key=
set tempid=
if not defined allapps call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f

for %%# in (
:: Windows 7
ae2ee509-1b34-41c0-acb7-6d4650168915_33PXH-7Y6KF-2VJC9-XBBR8-HV%f%THH_Enterprise
1cb6d605-11b3-4e14-bb30-da91c8e3983a_YDRBP-3D83W-TY26F-D46B2-XC%f%KRJ_EnterpriseN
b92e9980-b9d5-4821-9c94-140f632f6312_FJ82H-XT6CR-J8D7P-XQJJ2-GP%f%DD4_Professional
54a09a0d-d57b-4c10-8b69-a842d6590ad5_MRPKT-YTG23-K7D7T-X2JMM-QY%f%7MG_ProfessionalN
db537896-376f-48ae-a492-53d0547773d0_YBYF6-BHCR3-JPKRB-CDW7B-F9%f%BK4_Embedded_POSReady
aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395_73KQT-CD9G6-K7TQG-66MRP-CQ%f%22C_Embedded_ThinPC
5a041529-fef8-4d07-b06f-b59b573b32d2_W82YF-2Q76Y-63HXB-FGJG9-GF%f%7QX_ProfessionalE
46bbed08-9c7b-48fc-a614-95250573f4ea_C29WB-22CC8-VJ326-GHFJW-H9%f%DH4_EnterpriseE
:: Windows Server 2008 R2
68531fb9-5511-4989-97be-d11a0f55633f_YC6KT-GKW9T-YTKYR-T4X34-R7%f%VHC_ServerStandard
7482e61b-c589-4b7f-8ecc-46d455ac3b87_74YFP-3QFB3-KQT8W-PMXWJ-7M%f%648_ServerDatacenter
620e2b3d-09e7-42fd-802a-17a13652fe7a_489J6-VHDMP-X63PK-3K798-CP%f%X3Y_ServerEnterprise
7482e61b-c589-4b7f-8ecc-46d455ac3b87_74YFP-3QFB3-KQT8W-PMXWJ-7M%f%648_ServerDatacenterCore
68531fb9-5511-4989-97be-d11a0f55633f_YC6KT-GKW9T-YTKYR-T4X34-R7%f%VHC_ServerStandardCore
620e2b3d-09e7-42fd-802a-17a13652fe7a_489J6-VHDMP-X63PK-3K798-CP%f%X3Y_ServerEnterpriseCore
8a26851c-1c7e-48d3-a687-fbca9b9ac16b_GT63C-RJFQ3-4GMB6-BRFB9-CB%f%83V_ServerEnterpriseIA64
a78b8bd9-8017-4df5-b86a-09f756affa7c_6TPJF-RBVHG-WBW2R-86QPH-6R%f%TM4_ServerWeb
cda18cf3-c196-46ad-b289-60c072869994_TT8MH-CG224-D3D7Q-498W2-9Q%f%CTX_ServerHPC
a78b8bd9-8017-4df5-b86a-09f756affa7c_6TPJF-RBVHG-WBW2R-86QPH-6R%f%TM4_ServerWebCore
f772515c-0e87-48d5-a676-e6962c3e1195_736RG-XDKJK-V34PF-BHK87-J6%f%X3K_ServerEmbeddedSolution
:: Windows Vista
cfd8ff08-c0d7-452b-9f60-ef5c70c32094_VKK3X-68KWM-X2YGT-QR4M6-4B%f%WMV_Enterprise
4f3d1606-3fea-4c01-be3c-8d671c401e3b_YFKBB-PQJJV-G996G-VWGXY-2V%f%3X8_Business
2c682dc2-8b68-4f63-a165-ae291d4cf138_HMBQG-8H2RH-C77VX-27R82-VM%f%QBT_BusinessN
d4f54950-26f2-4fb4-ba21-ffab16afcade_VTC42-BM838-43QHV-84HX6-XJ%f%XKV_EnterpriseN
:: Windows Server 2008
ad2542d4-9154-4c6d-8a44-30f11ee96989_TM24T-X9RMF-VWXK6-X8JC9-BF%f%GM2_ServerStandard
68b6e220-cf09-466b-92d3-45cd964b9509_7M67G-PC374-GR742-YH8V4-TC%f%BY3_ServerDatacenter
c1af4d90-d1bc-44ca-85d4-003ba33db3b9_YQGMW-MPWTJ-34KDK-48M3W-X4%f%Q6V_ServerEnterprise
01ef176b-3e0d-422a-b4f8-4ea880035e8f_4DWFP-JF3DJ-B7DTH-78FJB-PD%f%RHK_ServerEnterpriseIA64
ddfa9f7c-f09e-40b9-8c1a-be877a9a7f4b_WYR28-R7TFJ-3X2YQ-YCY4H-M2%f%49D_ServerWeb
7afb1156-2c1d-40fc-b260-aab7442b62fe_RCTX3-KWVHP-BR6TB-RB6DM-6X%f%7HP_ServerComputeCluster
2401e3d0-c50a-4b58-87b2-7e794b7d2607_W7VD6-7JFBR-RX26B-YKQ3Y-6F%f%FFJ_ServerStandardV
fd09ef77-5647-4eff-809c-af2b64659a45_22XQ2-VRXRG-P8D42-K34TD-G3%f%QQC_ServerDatacenterV
8198490a-add0-47b2-b3ba-316b12d647b4_39BXF-X8Q23-P2WWT-38T2F-G3%f%FPG_ServerEnterpriseV
) do (
for /f "tokens=1-3 delims=_" %%A in ("%%#") do if %tsedition%==%%C if not defined key (
echo "%allapps%" | find /i "%%A" %nul1% && (
set key=%%B
set tempid=%%A
)
)
)

if not defined key (
call :dk_color %Red% "Checking Activation ID                  [%tsedition% SKU-%osSKU% %KS% key is not available]"
call :dk_color %Blue% "Use ZeroCID activation method from the previous menu."
goto :ts_esu
)

echo Checking Activation ID                  [%tempid%] [%tsedition%]

set oldks=1
set generickey=1
call :dk_inskey "[%key%]"
set tsids=%tsids% %tempid%
goto :ts_esu

::========================================================================================================================================

:ts_winvista

::  Process Windows Vista

::  1st column = Activation ID
::  2nd column = Generic key
::  3rd column = Key channel
::  4th column = Edition ID
::  Separator  = _

::  Keys aren't available for these editions, but since these editions aren't publicly available, it doesn't matter
::  a797d61e-1475-470b-86c8-f737a72c188d StarterN
::  5e9f548a-c8a9-44e6-a6c2-3f8d0a7a99dd ServerComputeClusterV

set f=
set key=
set tempid=
if not defined allapps call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f

for %%# in (
::  WindowsVista
9de9abe2-d01d-4538-af84-4498bdbc2ba3_4D2XH-PRBMM-8Q22B-K8BM3-MR%f%W4W_____Retail_Business
db442be4-81ed-4ab3-9d66-2417e8a5c81c_76884-QXFY2-6Q2WX-2QTQ8-QX%f%X44_____Retail_BusinessN
b51791c2-b562-4b73-97b0-735a0e4429a6_YQPQV-RW8R3-XMPFG-RXG9R-JG%f%TVF_____Retail_Enterprise
58c37517-42f8-4723-bb44-30b05791ff2a_Q7J9R-G63R4-BFMHF-FWM9R-RW%f%DMV_____Retail_EnterpriseN
95c6e80a-0ff8-4bd0-95f2-c4a39b79d09e_RCG7P-TX42D-HM8FM-TCFCW-3V%f%4VD_____Retail_HomeBasic
d0333dad-c14e-46f2-b62a-8b47a1b9768b_HY2VV-XC6FF-MD6WV-FPYBQ-GF%f%JBT_____Retail_HomeBasicN
9e042223-03bf-49ae-808f-ff37f128d40d_X9HTF-MKJQQ-XK376-TJ7T4-76%f%PKF_____Retail_HomePremium
92d8977c-d506-4e63-b500-6d39283b6cd5_KJ6TP-PF9W2-23T3Q-XTV7M-PX%f%DT2_____Retail_HomePremiumN
89e51a3c-76c0-4beb-a650-53d34c8f8186_X9PYV-YBQRV-9BXWV-TQDMK-QD%f%WK4_____Retail_Starter
30fab9cc-8614-4339-989f-7ce61fb7a5c4_VMCB9-FDRV6-6CDQM-RV23K-RP%f%8F7_____Retail_Ultimate
1eefed20-8ac0-478c-8774-70cd44782ea1_CVX38-P27B4-2X8BT-RXD4J-V7%f%CKX_____Retail_UltimateN
::  WindowsServer2008
c9ad502b-ef48-41d1-a2a0-38a38e82fed0_24FV9-H7JW8-C8Q6X-BQKMK-K9%f%77J_____Retail_ServerComputeCluster
866e924e-c2a3-4872-aca1-6b48c13962d5_6QBHY-DXTPJ-T9W3P-DTJXX-4V%f%QMB_____Retail_ServerDatacenter
d020c729-07f0-4f8f-87ce-bf803275c786_83TWG-TD3TC-HRDP2-K93FJ-Y3%f%4YC_OEM:NONSLP_ServerDatacenterV
32b40e5e-0c6d-4c6f-ab12-a031933fd2c6_MRB7H-QJRHG-FXTBR-B2Q2M-8W%f%MTJ_____Retail_ServerEnterprise
256cc990-1692-4ea8-965c-2d423d5dd24e_H4VB6-QPRWH-VDCYM-996P8-MH%f%KFY_OEM:NONSLP_ServerEnterpriseIA64
1ba5e036-e386-42c4-b7eb-16bdb4fa1945_H8H7M-HDPQT-PJHQF-M7B83-9C%f%VGV_____Retail_ServerEnterpriseV
8df04457-07c8-4301-bce9-d61eb76cb2d6_RGBMC-PQBVF-94Q9K-HD63B-VY%f%6MP_____Retail_ServerHomePremium
5bd23b19-aa71-4a5b-8b68-c8801c2baff6_6C8KR-MD3QK-9GWFW-44CY2-W9%f%CBM_____Retail_ServerHomeStandard
b86c7736-91ff-4de9-bfa9-b32b8a09acac_7XRBY-6MP2K-VQPT8-F37JV-YY%f%Q83_____Retail_ServerMediumBusinessManagement
d3f5642f-081d-40b2-a4b9-efd3054d4584_6PDTD-JK48J-662TF-8J2QV-R4%f%CRB_____Retail_ServerMediumBusinessMessaging
c6936a36-69f3-4994-9857-3069c7b9ec7a_D694V-CMWKH-PY92X-PFQKQ-JC%f%B69_____Retail_ServerMediumBusinessSecurity
cc4c2cf8-ef29-4d8e-b168-2b65a3db3309_MRDK3-YYQF3-88BQJ-D6FJG-69%f%YJY_____Retail_ServerSBSPremium
b3827b27-bd38-4284-98af-e4f4d1c051a0_2KB23-GJRBD-W3T9C-6CH2W-39%f%B7V_____Retail_ServerSBSPrime
5dad0eff-3f6f-4310-8844-422f9dc7c84b_H4XDD-B27GY-667P6-XWVV7-GY%f%G8J_____Retail_ServerSBSStandard
603504f9-109f-49f0-9271-8c66f7878f58_8YVM4-YQBDH-7WDQM-R27WR-WV%f%CWG_____Retail_ServerStandard
65ab7338-9ad0-43fe-af1b-190b577495e2_H9MW3-6V7GK-94P9G-7FTPJ-VK%f%CKF_____Retail_ServerStandardV
2be204da-24a0-4943-b66c-81e8464acd7e_2264C-TD9T8-P8HPW-CC9GH-MH%f%M2V_____Retail_ServerStorageEnterprise
60207eba-8b4a-486c-a013-023b4b742c2f_RCYMT-YX342-8T6YY-XYHYC-3D%f%D7X_____Retail_ServerStorageExpress
368856e9-43f7-4601-8358-e561f36c7dd8_FKFT2-WXYY9-WBPY7-6YMY4-X4%f%8JF_____Retail_ServerStorageStandard
4bf433fa-ab04-4c6c-b55b-00170e14b8cd_8X9J7-HCJ7J-3WDJT-QM7D8-46%f%4YH_____Retail_ServerStorageWorkgroup
a77a6806-f59e-4953-97d7-229317b8e6a6_BGT39-9FYH7-X2CYD-T628F-QP%f%QPW_____Retail_ServerWeb
f92f836d-4d3e-4e90-a08f-2d612d65e716_HPH76-FHFPP-DRW9D-7W2V4-HW%f%GKT_____Retail_ServerWinSB
3059a9fd-b068-4f0d-acaf-66324dca67ac_2V8G6-KRXYR-MMGXJ-6RWM3-GX%f%CCG_____Retail_ServerWinSBV
) do (
for /f "tokens=1-4 delims=_" %%A in ("%%#") do if %tsedition%==%%D if not defined key (
echo "%allapps%" | find /i "%%A" %nul1% && (
set key=%%B
set tempid=%%A
)
)
)

if not defined key (
set error=1
call :dk_color %Red% "Checking Activation ID                  [%tsedition% SKU-%osSKU% not found in the system]"
call :dk_color %Blue% "%_fixmsg%"
goto :ts_esu
)

echo Checking Activation ID                  [%tempid%] [%tsedition%]

set generickey=1
call :dk_inskey "[%key%]"
set tsids=%tsids% %tempid%
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

set generickey=1
call :dk_inskey "[%key%]"

::========================================================================================================================================

:ts_esu

if not %_actesu%==1 goto :ts_off
if /i %tsmethod%==KMS4k (
call :dk_color %Red% "Skipping Windows ESU                    [KMS4k method is not supported with Windows ESU]"
if /i %_actmethod%==Auto (
call :dk_color %Blue% "Connect to the Internet and try again. Script will use the StaticCID activation method."
) else (
call :dk_color %Blue% "Use non-KMS4K activation options from the previous menu."
)
goto :ts_off
)

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
REM WindowsServer2008/WindowsServer2008R2
8e7bfb1e-acc1-4f56-abae-b80fce56cd4b_Server-ESU-PA[1-6y]_-ServerDatacenter-ServerDatacenterCore-ServerDatacenterV-ServerDatacenterVCore-ServerStandard-ServerStandardCore-ServerStandardV-ServerStandardVCore-ServerEnterprise-ServerEnterpriseCore-ServerEnterpriseV-ServerEnterpriseVCore-
REM Windows8.1
4afc620f-12a4-48ad-8015-2aebfbd6e47c_Client-ESU-Year3[1-3y]_-Enterprise-EnterpriseN-Professional-ProfessionalN-
11be7019-a309-4763-9a09-091d1722ffe3_Client-FES-ESU-Year3[1-3y]_-EmbeddedIndustry-EmbeddedIndustryE-
REM WindowsServer2012/2012R2
55b1dd2d-2209-4ea0-a805-06298bad25b3_Server-ESU-Year3[1-3y]_-ServerDatacenter-ServerDatacenterCore-ServerDatacenterV-ServerDatacenterVCore-ServerStandard-ServerStandardCore-ServerStandardV-ServerStandardVCore-
REM Windows10
f520e45e-7413-4a34-a497-d2765967d094_Client-ESU-Year1_-Education-EducationN-Enterprise-EnterpriseN-Professional-ProfessionalEducation-ProfessionalEducationN-ProfessionalN-ProfessionalWorkstation-ProfessionalWorkstationN-ServerRdsh-
1043add5-23b1-4afb-9a0f-64343c8f3f8d_Client-ESU-Year2_-Education-EducationN-Enterprise-EnterpriseN-Professional-ProfessionalEducation-ProfessionalEducationN-ProfessionalN-ProfessionalWorkstation-ProfessionalWorkstationN-ServerRdsh-
83d49986-add3-41d7-ba33-87c7bfb5c0fb_Client-ESU-Year3_-Education-EducationN-Enterprise-EnterpriseN-Professional-ProfessionalEducation-ProfessionalEducationN-ProfessionalN-ProfessionalWorkstation-ProfessionalWorkstationN-ServerRdsh-
0b533b5e-08b6-44f9-b885-c2de291ba456_Client-ESU-Year6[4-6y]_-Education-EducationN-Enterprise-EnterpriseN-Professional-ProfessionalEducation-ProfessionalEducationN-ProfessionalN-ProfessionalWorkstation-ProfessionalWorkstationN-ServerRdsh-
b8527af1-5389-447c-9a88-2d1691ea33d3_Client-IoT-ESU-Year1_-IoTEnterprise-
7b76ee02-0a75-4f08-85d5-bd0feadad0c0_Client-IoT-ESU-Year2_-IoTEnterprise-
4dac5a0c-5709-4595-a32c-14a56a4a6b31_Client-IoT-ESU-Year3_-IoTEnterprise-
f69e2d51-3bbd-4ddf-8da7-a145e9dca597_Client-IoT-ESU-Year6[4-6y]_-IoTEnterprise-
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

if defined esuexistsup if defined _vis (
set key=9FPV7-MWGT8-7XPDF-JC23W-WT7TW
REM This is a non-generic blocked MAK key for Server-ESU-PA
call :dk_inskey "[!key!]"
goto :ts_off
)

if defined esuexistsup (
echo "%tsids%" | find /i "4220f546-f522-46df-8202-4d07afd26454" %nul1% && (
echo "%tsids%" | find /i "7e94be23-b161-4956-a682-146ab291774c" %nul1% || (
call :dk_color %Gray% "To get Client-ESU-Year6[4-6y] license, install updates from the below URL."
call :dk_color %Blue% "%mas%tsforge#windows-esu"
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
echo %esueditionlist% | find /i "Professional" %nul1% && (
call :dk_color %Blue% "Go back to Main Menu, select Change Windows Edition option, and change it to [Professional] or any non-Home edition."
) || (
call :dk_color %Blue% "Go back to Main Menu, select Change Windows Edition option and change to any of the below listed editions."
echo [%esueditionlist%]
)
goto :ts_off
)

set esuavail=
if defined _vis if defined isServer set esuavail=1
if %winbuild% LEQ 7602 if not defined _vis if not defined isThinpc set esuavail=1
if %winbuild% GTR 7602 if %winbuild% LSS 10240 if defined isServer set esuavail=1
if %winbuild% GEQ 10240 if %winbuild% LEQ 19045 if not defined isServer set esuavail=1
if %winbuild% EQU 9600 set esuavail=1

if defined esuavail (
call :dk_color %Red% "Checking Activation ID                  [ESU license is not found, make sure Windows is fully updated]"
set fixes=%fixes% %mas%tsforge#windows-esu
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%tsforge#windows-esu"
) else (
call :dk_color %Gray% "Checking Activation ID                  [ESU is not available for %winos%]"
)

::========================================================================================================================================

:ts_off

if not %_actoff%==1 goto :ts_act

if %winbuild% LSS 9200 (
echo:
call :dk_color %Gray% "Checking Supported Office               [TSforge for Office is supported on Windows 8 and later versions]"
call :dk_color %Blue% "On Windows Vista / 7, use Ohook activation option for Office instead."
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
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o14msi=Office 2010 MSI )
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o14msi=Office 2010 MSI )
%nul% reg query %_68%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 
%nul% reg query %_86%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 

if not "%o14msi%%o14c2r%"=="" (
echo:
call :dk_color %Red% "Checking Unsupported Office Install     [ %o14msi%%o14c2r%]"
if defined o14msi call :dk_color %Blue% "Use Ohook activation option for Office 2010."
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
echo You only have the Office Dashboard app installed. You need to install the full version of Office.
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

call :oh_expiredpreview 2013
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

call :oh_expiredpreview 2016 2019 2021 2024
if "%_actprojvis%"=="0" call :oh_fixprids
call :ts_process

::========================================================================================================================================

::  mass grave[.]dev/office-license-is-not-genuine
::  Add registry keys for volume products so that 'non-genuine' banner won't appear 

set "kmskey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663"
if /i %tsmethod%==KMS4k (
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

if /i %tsmethod%==KMS4k (
echo:
call :dk_color %Red% "Skipping Windows %KS% Host               [KMS4k method is not supported with Windows %KS% Host]"
if /i %_actmethod%==Auto (
call :dk_color %Blue% "Connect to the Internet and try again. Script will use the StaticCID activation method."
) else (
call :dk_color %Blue% "Use non-KMS4K activation options from the previous menu."
)
goto :ts_act
)

echo:
if %winbuild% GEQ 10586 (
call :dk_color %Gray% "With %KS% Host license, system may randomly change Windows Edition later. It is a Windows issue and can be safely ignored."
)
call :dk_color %Gray% "%KS% Host [Not to be confused with %KS% Client] license causes the %_slser% service to run continuously."
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

if defined _vis goto :ts_whost_vista

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

:ts_whost_vista

::  Process Windows K-M-S host for Vista

::  1st column = Activation ID
::  2nd column = CSVLK key
::  3rd column = Edition IDs
::  Separator  = _

set f=
set key=
set tempid=
if not defined allapps call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f

for %%# in (
::  WindowsVista
212a64dc-43b1-4d3d-a30c-2fc69d2095c6_TWVG3-9Q4P8-W9XJF-Y76FJ-DW%f%Q4R_-Business-BusinessN-Enterprise-EnterpriseN-
::  WindowsServer2008
c90d1b4e-8aa8-439e-8b9e-b6d6b6a6d975_BHC4Q-6D7B7-QMVH7-4MKQH-Y9%f%VK7_-ServerComputeCluster-ServerDatacenter-ServerDatacenterV-ServerEnterprise-ServerEnterpriseIA64-ServerEnterpriseV-ServerStandard-ServerStandardV-ServerWeb-
56df4151-1f9f-41bf-acaa-2941c071872b_PVGKG-2R7XQ-7WTFD-FXTJR-DQ%f%BQ3_-ServerComputeCluster-ServerEnterprise-ServerEnterpriseV-ServerStandard-ServerStandardV-ServerWeb-
c448fa06-49d1-44ec-82bb-0085545c3b51_KH4PC-KJFX6-XFVHQ-GDK2G-JC%f%JY9_-ServerComputeCluster-ServerWeb-
) do (
for /f "tokens=1-3 delims=_" %%A in ("%%#") do if not defined key (
echo "%allapps%" | find /i "%%A" %nul1% && (
echo "%%C" | find /i "-%tsedition%-" %nul1% && (
set key=%%B
set tempid=%%A
)
)
)
)

if defined key (
echo Checking Activation ID                  [%tempid%] [%tsedition%]
) else (
call :dk_color %Red% "Checking Activation ID                  [Not Found] [%tsedition%] [%osSKU%]"
call :dk_color %Blue% "%KS% Host license is not found on your system. It is available for the below editions."
call :dk_color %Blue% "Business, BusinessN, Enterprise, EnterpriseN, and Server editions, etc."
goto :ts_act
)

call :dk_inskey "[%key%]"
set tsids=%tsids% %tempid%
goto :ts_act

::========================================================================================================================================

:ts_ohost

::  Process Office K-M-S host

echo:
echo Processing Office %KS% Host...

if defined _vis (
echo:
call :dk_color %Blue% "Windows Vista and Server 2008 do not support the installation of Office KMS Host."
goto :ts_act
)

if /i %tsmethod%==KMS4k (
echo:
call :dk_color %Red% "Skipping Office %KS% Host                [KMS4k method is not supported with Office %KS% Host]"
if /i %_actmethod%==Auto (
call :dk_color %Blue% "Connect to the Internet and try again. Script will use the StaticCID activation method."
) else (
call :dk_color %Blue% "Use non-KMS4K activation options from the previous menu."
)
goto :ts_act
)

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
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%tsforge#office-kms-host"
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

if /i %tsmethod%==KMS4k (
echo:
call :dk_color %Red% "Skipping Windows 8/8.1 APPX             [KMS4k method is not supported with Windows 8/8.1 APPX]"
if /i %_actmethod%==Auto (
call :dk_color %Blue% "Connect to the Internet and try again. Script will use the StaticCID activation method."
) else (
call :dk_color %Blue% "Use non-KMS4K activation options from the previous menu."
)
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
if defined _vis (
echo Processing Reset of Rearm / Timers...
) else (
echo Processing Reset of Rearm / Timers / Tamper / Lock...
)
echo:

set resetstuff=1
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':tsforge\:.*';iex ($f[1])"

if %errorlevel%==3 (
call :dk_color %Red% "Reset Failed."
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
) else (
call :dk_color %Green% "Reset process has been successfully done."
)

goto :dk_done

::========================================================================================================================================

:ts_actman

echo:
echo Processing Manual Activation...
echo:

if /i %tsmethod%==KMS4k (
echo:
call :dk_color %Red% "Skipping Manual Activation              [KMS4k method is not supported with it]"
if /i %_actmethod%==Auto (
call :dk_color %Blue% "Connect to the Internet and try again. Script will use the StaticCID activation method."
) else (
call :dk_color %Blue% "Use non-KMS4K activation options from the previous menu."
)
goto :dk_done
)

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

if defined _vis (
echo:
call :dk_color %Blue% "On Windows Vista and Server 2008, you must manually install the key before activating it."
)
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
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
)
)

if defined tsids (
echo:
if not defined _vis if not defined oldks echo Installing Forged Product Key Data...
if /i %tsmethod%==KMS4k (
echo Writing TrustedStore data...
) else (
if /i %tsmethod%==StaticCID (echo Depositing Static Confirmation ID...) else (echo Depositing Zero Confirmation ID...)
)
echo:
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':tsforge\:.*';& ([ScriptBlock]::Create($f[1])) %tsids%"
if !errorlevel!==3 (
if %_actman%==0 (if not defined error call :dk_color %Blue% "%_fixmsg%")
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
) else (
echo "%tsids%" | find /i "7e94be23-b161-4956-a682-146ab291774c" %nul1% && (
call :dk_color %Gray% "Windows Update can receive 1-3 years of ESU. 4-6 years ESU is not officially supported, but you can manually install updates."
)
echo "%tsids%" | findstr /i "4afc620f-12a4-48ad-8015-2aebfbd6e47c 11be7019-a309-4763-9a09-091d1722ffe3" %nul1% && (
call :dk_color %Gray% "ESU is not officially supported on Windows 8.1, but you can manually install updates until Jan-2024."
)
echo "%tsids%" | findstr /i "0b533b5e-08b6-44f9-b885-c2de291ba456 f69e2d51-3bbd-4ddf-8da7-a145e9dca597" %nul1% && (
call :dk_color %Gray% "Windows Update can receive 1-3 years of ESU. 4-6 years ESU is not officially supported, but it might be useful."
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
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
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
echo TSforge activation doesn't modify any Windows components and doesn't install any new files.
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
set _oMSI=
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

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o16msi=1&set o16msi_reg=%_86%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o16msi=1&set o16msi_reg=%_68%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o15msi=1&set o15msi_reg=%_86%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o15msi=1&set o15msi_reg=%_68%\15.0)

exit /b

::========================================================================================================================================

:oh_expiredpreview

for %%# in (%*) do (
if exist "!_oLPath!\ProPlus%%#PreviewVL_*.xrm-ms" if not exist "!_oLPath!\ProPlus%%#VL_*.xrm-ms" (
set error=1
set showfix=1
call :dk_color %Red% "Checking Expired Preview Products       [Office %%# Preview Found]"
call :dk_color %Blue% "Please run the Office updates first, and then attempt to activate it again."
)
)

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

set foundprod=
call :tsksdata chkprod %%#
if defined _oMSI if not defined foundprod if /i %tsmethod%==KMS4k (
set skipprocess=1
call :dk_color %Red% "Checking Product In Script              [%%# MSI Retail is not supported with KMS4k]"
if /i %_actmethod%==Auto (
call :dk_color %Blue% "Connect to the Internet and try again. Script will use the StaticCID activation method."
) else (
call :dk_color %Blue% "Use non-KMS4K activation options from the previous menu."
)
)

if "%_actprojvis%"=="1" (
echo %%# | findstr /i "Project Visio" %nul% || (
set skipprocess=1
call :dk_color %Gray% "Skipping Because Project/Visio Mode     [%%#]"
)
)

if "%_actprojvis%"=="0" if /i %tsmethod%==KMS4k echo %_oIds% | findstr /i "O365" %nul% && (
echo %%# | findstr /i "Project Visio" %nul% && (
set skipprocess=1
echo Skipping Because Mondo Is Available     [%%#]
)
)


if not defined skipprocess (

if /i not %tsmethod%==KMS4k (
set no365=
if "%oVer%"=="15" (echo %%# | findstr /i "O365HomePremRetail" %nul% && set no365=1)
if "%oVer%"=="16" (echo %%# | findstr /i "O365" %nul% && set no365=1)

if defined no365 (
set _License=MondoRetail
set _altoffid=MondoRetail
call :ks_osppready
echo Converting Unsupported O365 Office      [%%# To MondoRetail]
if "%oVer%"=="15" (call :dk_color %Gray% "Mondo 2013 is equivalent to O365 [15.0 version] in terms of the latest features.")
if "%oVer%"=="16" (call :dk_color %Gray% "Mondo 2016 is equivalent to O365 in terms of the latest features.")
)

if not defined _oMSI (
echo %%# | findstr /i "ARM" %nul% && (
set _License=MondoRetail
set _altoffid=MondoRetail
call :ks_osppready
echo Converting Unsupported OEM-ARM Office   [%%# To MondoRetail]
)
)
)

if not defined _oMSI if /i %tsmethod%==KMS4k if not defined foundprod (
call :tsksdata getinfo %%#
if defined _altoffid (
echo Converting Retail To Volume             [%%# To !_altoffid!]
) else (
set _License=MondoVolume
set _altoffid=MondoVolume
echo Converting Retail To Volume             [%%# To !_altoffid!] [Using Mondo because %%# is not found in the script]
)
echo %%# | find /i "O365" %nul% && (
if "%oVer%"=="15" (call :dk_color %Gray% "Mondo 2013 is equivalent to O365 [15.0 version] in terms of the latest features.")
if "%oVer%"=="16" (call :dk_color %Gray% "Mondo 2016 is equivalent to O365 in terms of the latest features.")
)
call :ks_osppready
)

if /i %tsmethod%==KMS4k (
echo !_License! | find /i "Retail" %nul% && (set keytype=zero) || (set keytype=ks)
) else (
set keytype=zero
)

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
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
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

if /i not %tsmethod%==KMS4k if defined winserver if defined _config if exist "%_oLPath%\Word2019VL_KMS_Client_AE*.xrm-ms" (
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

set _oMSI=1
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

if exist "%_common%\Microsoft Shared\OFFICE%oVer%\Office Setup Controller\pkeyconfig-office.xrm-ms" (
set "pkeypath=%_common%\Microsoft Shared\OFFICE%oVer%\Office Setup Controller\pkeyconfig-office.xrm-ms"
) else if exist "%_common2%\Microsoft Shared\OFFICE%oVer%\Office Setup Controller\pkeyconfig-office.xrm-ms" (
set "pkeypath=%_common2%\Microsoft Shared\OFFICE%oVer%\Office Setup Controller\pkeyconfig-office.xrm-ms"
)

call :msiofficedata %2

echo:
echo Processing Office...                    [MSI ^| %_version% ^| %_oArch%]

if not defined _oIds (
set error=1
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
exit /b
)

call :ts_process
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

:oh_clearblock

::  Find remnants of Office vNext/shared/device license block and remove it because it stops other licenses from appearing
::  https://learn.microsoft.com/en-us/office/troubleshoot/activation/reset-office-365-proplus-activation-state

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
::  https://learn.microsoft.com/en-us/deployoffice/device-based-licensing

for /f %%# in ('reg query "%o16c2r_reg%\Configuration" /f *.DeviceBasedLicensing %nul6% ^| findstr REG_') do reg delete "%o16c2r_reg%\Configuration" /v %%# /f %nul%

::  Remove OEM registry key
::  https://support.microsoft.com/en-us/office/office-repeatedly-prompts-you-to-activate-on-a-new-pc-a9a6b05f-f6ce-4d1f-8d49-eb5007b64ba1

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

::  Clean existing K-M-S cache from the registry

:_taskclear-cache

set w=
for %%# in (SppE%w%xtComObj.exe sppsvc.exe SLsvc.exe) do (
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%w%ge File Execu%w%tion Options\%%#" /f %nul%
)

set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"

if %winbuild% LSS 7600 (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL" %nul% && (
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"
)
)
if not defined SPPk (
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
)

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

::  Get Windows permanent activation status (not counting csvlk)

:ts_checkwinperm

%psc% "Get-WmiObject -Query 'SELECT Name, Description FROM SoftwareLicensingProduct WHERE LicenseStatus=''1'' AND GracePeriodRemaining=''0'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn IS NULL' | Where-Object { $_.Description -notmatch 'KMS' } | Select-Object -Property Name" %nul2% | findstr /i "Windows" %nul1% && set _perm=1||set _perm=
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

if defined generickey (set "keyecho=Installing Generic Product Key         ") else (set "keyecho=Installing Product Key                 ")
if %keyerror% EQU 0 (
if %sps%==SoftwareLicensingService call :dk_refresh
echo %keyecho% %~1 [Successful]
) else (
call :dk_color %Red% "%keyecho% %~1 [Failed] %keyerror%"
if not defined error (
if defined altapplist call :dk_color %Red% "Activation ID not found for this key."
call :dk_color %Blue% "%_fixmsg%"
set showfix=1
)
set error=1
)

set generickey=
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
echo:!allapps!> "!_ttemp!\chklen"
for %%A in ("!_ttemp!\chklen") do (set len=%%~zA)
del "!_ttemp!\chklen" %nul%

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

if %winbuild% LSS 7600 exit /b

::  This key is left by the system in rearm process and sppsvc sometimes fails to delete it, it causes issues in working of the Scheduled Tasks of SPP

set "ruleskey=HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\PersistedSystemState"
reg delete "%ruleskey%" /v "State" /f %nul%
reg delete "%ruleskey%" /v "SuppressRulesEngine" /f %nul%

set r1=$TB = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0);
set r2=%r1% [void]$TB.DefinePInvokeMethod('SLpTriggerServiceWorker', 'sppc.dll', 22, 1, [Int32], @([UInt32], [IntPtr], [String], [UInt32]), 1, 3);
set d1=%r2% [void]$TB.CreateType()::SLpTriggerServiceWorker(0, 0, 'reeval', 0)
%psc% "Start-Job { Stop-Service sppsvc -force } | Wait-Job -Timeout 20 | Out-Null; %d1%"
exit /b

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
	Get-ChildItem $Loc -Recurse -Filter *.xrm-ms | ForEach-Object {InstallLicenseFile $_.FullName}
}
function ReinstallLicenses() {
	$Paths = @("$env:SysPath\oem", "$env:SysPath\licensing", "$env:SysPath\spp\tokens")
	foreach ($Path in $Paths) {
    if (Test-Path $Path) { InstallLicenseDir "$Path" }
	}
}
:xrm:

::  Check wmic.exe

:dk_ckeckwmic

if %winbuild% LSS 9200 (set _wmic=1&exit /b)
set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
cmd /c "wmic path Win32_ComputerSystem get CreationClassName /value" %nul2% | find /i "computersystem" %nul1% && set _wmic=1
)
exit /b

::  Show info for potential script stuck scenario

:dk_sppissue

sc start %_slser% %nul%
set spperror=%errorlevel%

if %spperror% NEQ 1056 if %spperror% NEQ 0 (
%eline%
echo sc start %_slser% [Error Code: %spperror%]
)

echo:
%psc% "$job = Start-Job { (Get-WmiObject -Query 'SELECT * FROM %sps%').Version }; if (-not (Wait-Job $job -Timeout 30)) {write-host '%_slser% is not working correctly. Check this webpage for help - %mas%troubleshoot'}"
exit /b

::  Get Product name (WMI/REG methods are not reliable in all conditions, hence winbrand.dll method is used)

:dk_product

set d1=%ref% $meth = $TypeBuilder.DefinePInvokeMethod('BrandingFormatString', 'winbrand.dll', 'Public, Static', 1, [String], @([String]), 1, 3);
set d1=%d1% $meth.SetImplementationFlags(128); $TypeBuilder.CreateType()::BrandingFormatString('%%WINDOWS_LONG%%') -replace [string][char]0xa9, '' -replace [string][char]0xae, '' -replace [string][char]0x2122, ''

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

sc start %_slser% %nul%
echo "%errorlevel%" | findstr "577 225" %nul% && (
set "results=%results%[Likely File Infector]"
) || (
if not exist %SysPath%\%_slexe% if not exist %SysPath%\alg.exe (set "results=%results%[Likely File Infector]")
)

if not "%results%%pupfound%"=="" (
if defined pupfound call :dk_color %Gray% "Checking PUP Activators                 [Found%pupfound%]"
if defined results call :dk_color %Red% "Checking Probable Mal%w%ware Infection..."
if defined results (call :dk_color %Red% "%results%"&set showfix=1)
set fixes=%fixes% %mas%remove_mal%w%ware
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%remove_mal%w%ware"
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
if /i %%#==SLsvc            sc config %%# start= auto %nul%
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
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%fix_service"
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


::  https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/windows-setup-states

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
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%evaluation_editions"
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

if not exist "%SysPath%\spp\tokens\skus\%osedition%\%osedition%*.xrm-ms" if not exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%osedition%\*-%osedition%-*.xrm-ms" if not exist "%SysPath%\licensing\skus\Security-Licensing-SLC-Component-SKU-%osedition%\*-%osedition%-*.xrm-ms" (
set skunotfound=1
call :dk_color %Red% "Checking License Files                  [Not Found] [%osedition%]"
)

if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*-%osedition%-*.mum" (
if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-%osedition%Edition*.mum" (
call :dk_color %Red% "Checking Package Files                  [Not Found] [%osedition%]"
)
)
)
)


if %_wmic% EQU 1 wmic path %sps% get Version %nul%
if %_wmic% EQU 0 %psc% "try { $null=([WMISEARCHER]'SELECT * FROM %sps%').Get().Version; exit 0 } catch { exit $_.Exception.InnerException.HResult }" %nul%
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


for %%# in (SppEx%w%tComObj.exe SLsvc.exe sppsvc.exe sppsvc.exe\PerfOptions) do (
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


if %winbuild% GEQ 7600 for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" %nul6%') do if /i %%b NEQ 0x0 (
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" /t REG_DWORD /d "0" /f %nul%
call :dk_color %Red% "Checking SkipRearm                      [Default 0 Value Not Found. Changing To 0]"
%psc% "Start-Job { Stop-Service sppsvc -force } | Wait-Job -Timeout 20 | Out-Null"
)


if %winbuild% GEQ 7600 reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Plugins\Objects\msft:rm/algorithm/hwid/4.0" /f ba02fed39662 /d %nul% || (
call :dk_color %Red% "Checking SPP Registry Key               [Incorrect ModuleId Found]"
set fixes=%fixes% %mas%issues_due_to_gaming_spoofers
call :dk_color2 %Blue% "Most likely caused by gaming spoofers. Check this webpage for help - " %_Yellow% " %mas%issues_due_to_gaming_spoofers"
set error=1
set showfix=1
)


set tokenstore=
if %winbuild% GEQ 7600 (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore %nul6%') do call set "tokenstore=%%b"
if %winbuild% LSS 9200 set "tokenstore=%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform"
if %winbuild% GEQ 9200 if /i not "!tokenstore!"=="%SysPath%\spp\store" if /i not "!tokenstore!"=="%SysPath%\spp\store\2.0" if /i not "!tokenstore!"=="%SysPath%\spp\store_test\2.0" (
set toerr=1
set error=1
set showfix=1
call :dk_color %Red% "Checking TokenStore Registry Key        [Correct Path Not Found] [!tokenstore!]"
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
)
)

::  This code creates token folder only if it's missing and sets default permission for it

if %winbuild% GEQ 7600 if not defined toerr if not exist "%tokenstore%\" (
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
%psc% "if (-not $env:_vis) {Start-Job { Stop-Service %_slser% -force } | Wait-Job -Timeout 20 | Out-Null}; $sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
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


if %winbuild% GEQ 7600 if exist "%tokenstore%\" if not exist "%tokenstore%\tokens.dat" (
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

REM  https://learn.microsoft.com/en-us/office/troubleshoot/activation/license-issue-when-start-office-application

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
    if ($subkeyName -match '8DEC0AF1-0341-4b93-85CD-72606C2DF94C.*') {
        $count++
    }
}
$osVersion = [System.Environment]::OSVersion.Version
$minBuildNumber = 14393
if ($osVersion.Build -ge $minBuildNumber) {
    $subkeyHashTable = @{}
    foreach ($subkeyName in $wpaKey.GetSubKeyNames()) {
        if ($subkeyName -match '8DEC0AF1-0341-4b93-85CD-72606C2DF94C.*') {
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
    if ($_ -match '8DEC0AF1-0341-4b93-85CD-72606C2DF94C.*') {
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

:dk_done

echo:
if %_unattended%==1 timeout /t 2 & exit /b

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

:tsforge:
$src = @'
#if !POWERSHELL2
namespace ActivationWs
{

/*

This code is adapted from ActivationWs.
Original Repository: https://github.com/dadorner-msft/activationws

MIT License

Copyright (c) Daniel Dorner

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

*/

using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

    public static class ActivationHelper {
        // Key for HMAC/SHA256 signature.
        private static readonly byte[] MacKey = new byte[64] {
            254,  49, 152, 117, 251,  72, 132, 134,
            156, 243, 241, 206, 153, 168, 144, 100,
            171,  87,  31, 202,  71,   4,  80,  88,
            48,   36, 226,  20,  98, 135, 121, 160,
            0,     0,   0,   0,   0,   0,   0,   0,
            0,     0,   0,   0,   0,   0,   0,   0,
            0,     0,   0,   0,   0,   0,   0,   0,
            0,     0,   0,   0,   0,   0,   0,   0
        };

        private const string Action = "http://www.microsoft.com/BatchActivationService/BatchActivate";

        private static readonly Uri Uri = new Uri("https://activation.sls.microsoft.com/BatchActivation/BatchActivation.asmx");

        private static readonly XNamespace SoapSchemaNs = "http://schemas.xmlsoap.org/soap/envelope/";
        private static readonly XNamespace XmlSchemaInstanceNs = "http://www.w3.org/2001/XMLSchema-instance";
        private static readonly XNamespace XmlSchemaNs = "http://www.w3.org/2001/XMLSchema";
        private static readonly XNamespace BatchActivationServiceNs = "http://www.microsoft.com/BatchActivationService";
        private static readonly XNamespace BatchActivationRequestNs = "http://www.microsoft.com/DRM/SL/BatchActivationRequest/1.0";
        private static readonly XNamespace BatchActivationResponseNs = "http://www.microsoft.com/DRM/SL/BatchActivationResponse/1.0";

        public static string CallWebService(int requestType, string installationId, string extendedProductId) {
            XDocument soapRequest = CreateSoapRequest(requestType, installationId, extendedProductId);
            HttpWebRequest webRequest = CreateWebRequest(soapRequest);
            XDocument soapResponse = new XDocument();

            try {
                IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);
                asyncResult.AsyncWaitHandle.WaitOne();

                // Read data from the response stream.
                using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult))
                using (StreamReader streamReader = new StreamReader(webResponse.GetResponseStream())) {
                    soapResponse = XDocument.Parse(streamReader.ReadToEnd());
                }

                return ParseSoapResponse(soapResponse);

            } catch {
                throw;
            }
        }

        private static XDocument CreateSoapRequest(int requestType, string installationId, string extendedProductId) {
            // Create an activation request.           
            XElement activationRequest = new XElement(BatchActivationRequestNs + "ActivationRequest",
                new XElement(BatchActivationRequestNs + "VersionNumber", "2.0"),
                new XElement(BatchActivationRequestNs + "RequestType", requestType),
                new XElement(BatchActivationRequestNs + "Requests",
                    new XElement(BatchActivationRequestNs + "Request",
                        new XElement(BatchActivationRequestNs + "PID", extendedProductId),
                        requestType == 1 ? new XElement(BatchActivationRequestNs + "IID", installationId) : null)
                )
            );

            // Get the unicode byte array of activationRequest and convert it to Base64.
            byte[] bytes = Encoding.Unicode.GetBytes(activationRequest.ToString());
            string requestXml = Convert.ToBase64String(bytes);

            XDocument soapRequest = new XDocument();

            using (HMACSHA256 hMACSHA = new HMACSHA256(MacKey)) {
                // Convert the HMAC hashed data to Base64.
                string digest = Convert.ToBase64String(hMACSHA.ComputeHash(bytes));

                soapRequest = new XDocument(
                new XDeclaration("1.0", "UTF-8", "no"),
                new XElement(SoapSchemaNs + "Envelope",
                    new XAttribute(XNamespace.Xmlns + "soap", SoapSchemaNs),
                    new XAttribute(XNamespace.Xmlns + "xsi", XmlSchemaInstanceNs),
                    new XAttribute(XNamespace.Xmlns + "xsd", XmlSchemaNs),
                    new XElement(SoapSchemaNs + "Body",
                        new XElement(BatchActivationServiceNs + "BatchActivate",
                            new XElement(BatchActivationServiceNs + "request",
                                new XElement(BatchActivationServiceNs + "Digest", digest),
                                new XElement(BatchActivationServiceNs + "RequestXml", requestXml)
                            )
                        )
                    )
                ));

            }

            return soapRequest;
        }

        private static HttpWebRequest CreateWebRequest(XDocument soapRequest) {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(Uri);
            webRequest.Accept = "text/xml";
            webRequest.ContentType = "text/xml; charset=\"utf-8\"";
            webRequest.Headers.Add("SOAPAction", Action);
            webRequest.Host = "activation.sls.microsoft.com";
            webRequest.Method = "POST";

            try {
                // Insert SOAP envelope
                using (Stream stream = webRequest.GetRequestStream()) {
                    soapRequest.Save(stream);
                }

                return webRequest;

            } catch {
                throw;
            }
        }

        private static string ParseSoapResponse(XDocument soapResponse) {
            if (soapResponse == null) {
                throw new ArgumentNullException("soapResponse", "The remote server returned an unexpected response.");
            }

            if (!soapResponse.Descendants(BatchActivationServiceNs + "ResponseXml").Any()) {
                throw new Exception("The remote server returned an unexpected response");
            }

            try {
                XDocument responseXml = XDocument.Parse(soapResponse.Descendants(BatchActivationServiceNs + "ResponseXml").First().Value);

                if (responseXml.Descendants(BatchActivationResponseNs + "ErrorCode").Any()) {
                    string errorCode = responseXml.Descendants(BatchActivationResponseNs + "ErrorCode").First().Value;

                    switch (errorCode) {
                        case "0x7F":
                            throw new Exception("The Multiple Activation Key has exceeded its limit");

                        case "0x67":
                            throw new Exception("The product key has been blocked");

                        case "0x68":
                            throw new Exception("Invalid product key");

                        case "0x86":
                            throw new Exception("Invalid key type");

                        case "0x90":
                            throw new Exception("Please check the Installation ID and try again");

                        default:
                            throw new Exception(string.Format("The remote server reported an error ({0})", errorCode));
                    }

                } else if (responseXml.Descendants(BatchActivationResponseNs + "ResponseType").Any()) {
                    string responseType = responseXml.Descendants(BatchActivationResponseNs + "ResponseType").First().Value;

                    switch (responseType) {
                        case "1":
                            string confirmationId = responseXml.Descendants(BatchActivationResponseNs + "CID").First().Value;
                            return confirmationId;

                        case "2":
                            string activationsRemaining = responseXml.Descendants(BatchActivationResponseNs + "ActivationRemaining").First().Value;
                            return activationsRemaining;

                        default:
                            throw new Exception("The remote server returned an unrecognized response");
                    }

                } else {
                    throw new Exception("The remote server returned an unrecognized response");
                }

            } catch {
                throw;
            }
        }
    }
}
#endif

// Common.cs
namespace LibTSforge
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;

    public enum PSVersion
    {
        Vista,
        Win7,
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

        // 2^31 - 8 minutes
        public static readonly ulong TimerMax = (ulong)TimeSpan.FromMinutes(2147483640).Ticks;

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
        [DllImport("kernel32.dll")]
        public static extern uint GetSystemDefaultLCID();

        [DllImport("kernel32.dll")]
        public static extern bool Wow64EnableWow64FsRedirection(bool Wow64FsEnableRedirection);

        public static string DecodeString(byte[] data)
        {
            return Encoding.Unicode.GetString(data).Trim('\0');
        }

        public static byte[] EncodeString(string str)
        {
            return Encoding.Unicode.GetBytes(str + '\0');
        }

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

            throw new NotSupportedException("Unable to auto-detect version info");
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
        public readonly Dictionary<Guid, ProductConfig> Products = new Dictionary<Guid, ProductConfig>();
        private readonly List<Guid> loadedPkeyConfigs = new List<Guid>();

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

                    KeyRange keyRange = new KeyRange
                    {
                        Start = int.Parse(rangeNode.SelectSingleNode("./p:Start", nsmgr).InnerText),
                        End = int.Parse(rangeNode.SelectSingleNode("./p:End", nsmgr).InnerText),
                        EulaType = rangeNode.SelectSingleNode("./p:EulaType", nsmgr).InnerText,
                        PartNumber = rangeNode.SelectSingleNode("./p:PartNumber", nsmgr).InnerText,
                        Valid = rangeNode.SelectSingleNode("./p:IsValid", nsmgr).InnerText.ToLower() == "true"
                    };

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
                        ProductConfig productConfig = new ProductConfig
                        {
                            GroupId = group,
                            Edition = configNode.SelectSingleNode("./p:EditionId", nsmgr).InnerText,
                            Description = configNode.SelectSingleNode("./p:ProductDescription", nsmgr).InnerText,
                            Channel = configNode.SelectSingleNode("./p:ProductKeyType", nsmgr).InnerText,
                            Randomized = configNode.SelectSingleNode("./p:ProductKeyType", nsmgr).InnerText.ToLower() == "true",
                            Algorithm = algorithms[group],
                            Ranges = keyRanges,
                            ActivationId = refActId
                        };

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
    }
}


// SPP/ProductKey.cs
namespace LibTSforge.SPP
{
    using System;
    using System.IO;
    using System.Linq;
    using Crypto;
    using PhysicalStore;

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
        public readonly string EulaType;
        public readonly string PartNumber;
        public readonly string Edition;
        public readonly string Channel;
        public readonly Guid ActivationId;

        private string mpc;
        private string pid2;

        public byte[] KeyBytes
        {
            get { return BitConverter.GetBytes(klow).Concat(BitConverter.GetBytes(khigh)).ToArray(); }
        }

        public ProductKey()
        {

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
            VariableBag pkb = new VariableBag(PSVersion.WinModern);
            pkb.Blocks.AddRange(new[]
            {
                new CRCBlockModern
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingProductKey",
                    ValueAsStr = ToString()
                },
                new CRCBlockModern
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingMiscData",
                    Value = new byte[] { }
                },
                new CRCBlockModern
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingAlgorithm",
                    ValueAsStr = GetAlgoUri()
                }
            });

            return new Guid(CryptoUtils.SHA256Hash(pkb.Serialize()).Take(16).ToArray());
        }

        public string GetMPC()
        {
            if (mpc != null)
            {
                return mpc;
            }

            int build = Environment.OSVersion.Version.Build;

            mpc = build >= 10240 ? "03612" :
                    build >= 9600 ? "06401" :
                    build >= 9200 ? "05426" :
                    "55041";

            // setup.cfg doesn't exist in Windows 8+
            string setupcfg = string.Format(@"{0}\oobe\{1}", Environment.SystemDirectory, "setup.cfg");

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
                ulong shortauth = ((ulong)Group << 41) | (Security << 31) | ((ulong)Serial << 1) | (Upgrade ? (ulong)1 : 0);
                return BitConverter.GetBytes(shortauth);
            }

            int serialHigh = Serial / 1000000;
            int serialLow = Serial % 1000000;

            BinaryWriter writer = new BinaryWriter(new MemoryStream());
            string algoId = Algorithm == PKeyAlgorithm.PKEY2005 ? "B8731595-A2F6-430B-A799-FBFFB81A8D73" : "660672EF-7809-4CFD-8D54-41B7FB738988";

            writer.Write(new Guid(algoId).ToByteArray());
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

        [DllImport("slc.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SLOpen(out IntPtr hSLC);

        [DllImport("slc.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SLClose(IntPtr hSLC);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetWindowsInformationDWORD(string ValueName, ref int Value);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLInstallProofOfPurchase(IntPtr hSLC, string pwszPKeyAlgorithm, string pwszPKeyString, uint cbPKeySpecificData, byte[] pbPKeySpecificData, ref Guid PKeyId);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLUninstallProofOfPurchase(IntPtr hSLC, ref Guid PKeyId);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetPKeyInformation(IntPtr hSLC, ref Guid pPKeyId, string pwszValueName, out SLDATATYPE peDataType, out uint pcbValue, out IntPtr ppbValue);

        [DllImport("slcext.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLActivateProduct(IntPtr hSLC, ref Guid pProductSkuId, byte[] cbAppSpecificData, byte[] pvAppSpecificData, byte[] pActivationInfo, string pwszProxyServer, ushort wProxyPort);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGenerateOfflineInstallationId(IntPtr hSLC, ref Guid pProductSkuId, ref string ppwszInstallationId);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLDepositOfflineConfirmationId(IntPtr hSLC, ref Guid pProductSkuId, string pwszInstallationId, string pwszConfirmationId);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetSLIDList(IntPtr hSLC, SLIDTYPE eQueryIdType, ref Guid pQueryId, SLIDTYPE eReturnIdType, out uint pnReturnIds, out IntPtr ppReturnIds);

        [DllImport("slc.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SLGetLicensingStatusInformation(IntPtr hSLC, ref Guid pAppID, IntPtr pProductSkuId, string pwszRightName, out uint pnStatusCount, out IntPtr ppLicensingStatus);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetInstalledProductKeyIds(IntPtr hSLC, ref Guid pProductSkuId, out uint pnProductKeyIds, out IntPtr ppProductKeyIds);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLConsumeWindowsRight(uint unknown);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetProductSkuInformation(IntPtr hSLC, ref Guid pProductSkuId, string pwszValueName, out SLDATATYPE peDataType, out uint pcbValue, out IntPtr ppbValue);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLGetLicense(IntPtr hSLC, ref Guid pLicenseFileId, out uint pcbLicenseFile, out IntPtr ppbLicenseFile);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLSetCurrentProductKey(IntPtr hSLC, ref Guid pProductSkuId, ref Guid pProductKeyId);

        [DllImport("slc.dll", CharSet = CharSet.Unicode)]
        private static extern uint SLFireEvent(IntPtr hSLC, string pwszEventId, ref Guid pApplicationId);

        private class SLContext : IDisposable
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
                uint count;
                IntPtr pProductKeyIds;

                uint status = SLGetSLIDList(sl.Handle, SLIDTYPE.SL_ID_PRODUCT_SKU, ref actId, SLIDTYPE.SL_ID_PKEY, out count, out pProductKeyIds);

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

        public static void RefreshTrustedTime(Guid actId)
        {
            using (SLContext sl = new SLContext())
            {
                SLDATATYPE type;
                uint count;
                IntPtr ppbValue;

                SLGetProductSkuInformation(sl.Handle, ref actId, "TrustedTime", out type, out count, out ppbValue);
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
                uint count;
                IntPtr pAppIds;

                uint status = SLGetSLIDList(sl.Handle, SLIDTYPE.SL_ID_PRODUCT_SKU, ref actId, SLIDTYPE.SL_ID_APPLICATION, out count, out pAppIds);

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
                uint count;
                IntPtr ppReturnLics;

                uint status = SLGetSLIDList(sl.Handle, SLIDTYPE.SL_ID_LICENSE, ref licId, SLIDTYPE.SL_ID_LICENSE_FILE, out count, out ppReturnLics);

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
                return status != 0xC004F012;
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

        public static void UninstallProductKey(Guid pkeyId)
        {
            using (SLContext sl = new SLContext())
            {
                SLUninstallProofOfPurchase(sl.Handle, ref pkeyId);
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


// SPP/SPPUtils.cs
namespace LibTSforge.SPP
{
    using Microsoft.Win32;
    using System;
    using System.IO;
    using System.Linq;
    using System.ServiceProcess;
    using Crypto;
    using PhysicalStore;
    using TokenStore;

    public static class SPPUtils
    {
        public static void KillSPP(PSVersion version)
        {
            ServiceController sc;

            string svcName = version == PSVersion.Vista ? "slsvc" : "sppsvc";

            try
            {
                sc = new ServiceController(svcName);

                if (sc.Status == ServiceControllerStatus.Stopped)
                    return;
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException(string.Format("Unable to access {0}: ", svcName) + ex.Message);
            }

            Logger.WriteLine(string.Format("Stopping {0}...", svcName));

            bool stopped = false;

            for (int i = 0; stopped == false && i < 1080; i++)
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
                catch (InvalidOperationException ex)
                {
                    Logger.WriteLine("Warning: Stopping sppsvc failed, retrying. Details: " + ex.Message);
                    System.Threading.Thread.Sleep(500);
                    continue;
                }

                stopped = true;
            }

            if (!stopped)
                throw new System.TimeoutException(string.Format("Failed to stop {0}", svcName));

            Logger.WriteLine(string.Format("{0} stopped successfully.", svcName));

            if (version == PSVersion.Vista && SPSys.IsSpSysRunning())
            {
                Logger.WriteLine("Unloading spsys...");

                int status = SPSys.ControlSpSys(false);

                if (status < 0)
                {
                    throw new IOException("Failed to unload spsys");
                }

                Logger.WriteLine("spsys unloaded successfully.");
            }
        }

        public static void RestartSPP(PSVersion version)
        {
            if (version == PSVersion.Vista)
            {
                ServiceController sc;

                try
                {
                    sc = new ServiceController("slsvc");

                    if (sc.Status == ServiceControllerStatus.Running)
                        return;
                }
                catch (InvalidOperationException ex)
                {
                    throw new InvalidOperationException("Unable to access slsvc: " + ex.Message);
                }

                Logger.WriteLine("Starting slsvc...");

                bool started = false;

                for (int i = 0; started == false && i < 360; i++)
                {
                    try
                    {
                        if (sc.Status != ServiceControllerStatus.StartPending)
                            sc.Start();

                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromMilliseconds(500));
                    }
                    catch (System.ServiceProcess.TimeoutException)
                    {
                        continue;
                    }
                    catch (InvalidOperationException ex)
                    {
                        Logger.WriteLine("Warning: Starting slsvc failed, retrying. Details: " + ex.Message);
                        System.Threading.Thread.Sleep(500);
                        continue;
                    }

                    started = true;
                }

                if (!started)
                    throw new System.TimeoutException("Failed to start slsvc");

                Logger.WriteLine("slsvc started successfully.");
            }

            SLApi.RefreshLicenseStatus();
        }

        public static bool DetectCurrentKey()
        {
            SLApi.RefreshLicenseStatus();

            using (RegistryKey wpaKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\WPA"))
            {
                foreach (string subKey in wpaKey.GetSubKeyNames())
                {
                    if (subKey.StartsWith("8DEC0AF1"))
                    {
                        return subKey.Contains("P");
                    }
                }
            }

            throw new FileNotFoundException("Failed to autodetect key type, specify physical store key with /prod or /test arguments.");
        }

        public static string GetPSPath(PSVersion version)
        {
            switch (version)
            {
                case PSVersion.Vista:
                case PSVersion.Win7:
                    return Directory.GetFiles(
                        Environment.GetFolderPath(Environment.SpecialFolder.System),
                        "7B296FB0-376B-497e-B012-9C450E1B7327-*.C7483456-A289-439d-8115-601632D005A0")
                    .FirstOrDefault() ?? "";
                default:
                    string psDir = Environment.ExpandEnvironmentVariables(
                        (string)Registry.GetValue(
                            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform",
                            "TokenStore",
                            ""
                        )
                    );
                    string psPath = Path.Combine(psDir, "data.dat");

                    if (string.IsNullOrEmpty(psDir) || !File.Exists(psPath))
                    {
                        string[] psDirs =
                        {
                            @"spp\store",
                            @"spp\store\2.0",
                            @"spp\store_test",
                            @"spp\store_test\2.0"
                        };

                        foreach (string dir in psDirs)
                        {
                            psPath = Path.Combine(
                                Path.Combine(
                                    Environment.GetFolderPath(Environment.SpecialFolder.System),
                                    dir
                                ),
                                "data.dat"
                            );

                            if (File.Exists(psPath)) return psPath;
                        }
                    } 
                    else
                    {
                        return psPath;
                    }

                    throw new FileNotFoundException("Failed to locate physical store.");
            }
        }

        public static string GetTokensPath(PSVersion version)
        {
            switch (version)
            {
                case PSVersion.Vista:
                    return Path.Combine(
                        Environment.ExpandEnvironmentVariables("%WINDIR%"),
                        @"ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareLicensing\tokens.dat"
                    );
                case PSVersion.Win7:
                    return Path.Combine(
                        Environment.ExpandEnvironmentVariables("%WINDIR%"),
                        @"ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat"
                    );
                default:
                    string tokDir = Environment.ExpandEnvironmentVariables(
                        (string)Registry.GetValue(
                            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform",
                            "TokenStore",
                            ""
                        )
                    );
                    string tokPath = Path.Combine(tokDir, "tokens.dat");

                    if (string.IsNullOrEmpty(tokDir) || !File.Exists(tokPath))
                    {
                        string[] tokDirs =
                        {
                            @"spp\store",
                            @"spp\store\2.0",
                            @"spp\store_test",
                            @"spp\store_test\2.0"
                        };

                        foreach (string dir in tokDirs)
                        {
                            tokPath = Path.Combine(
                                Path.Combine(
                                    Environment.GetFolderPath(Environment.SpecialFolder.System),
                                    dir
                                ),
                                "tokens.dat"
                            );

                            if (File.Exists(tokPath)) return tokPath;
                        }
                    }
                    else
                    {
                        return tokPath;
                    }

                    throw new FileNotFoundException("Failed to locate token store.");
            }
        }

        public static IPhysicalStore GetStore(PSVersion version, bool production)
        {
            string psPath = GetPSPath(version);

            switch (version)
            {
                case PSVersion.Vista:
                    return new PhysicalStoreVista(psPath, production);
                case PSVersion.Win7:
                    return new PhysicalStoreWin7(psPath, production);
                default:
                    return new PhysicalStoreModern(psPath, production, version);
            }
        }

        public static ITokenStore GetTokenStore(PSVersion version)
        {
            string tokPath = GetTokensPath(version);

            return new TokenStoreModern(tokPath);
        }

        public static void DumpStore(PSVersion version, bool production, string filePath, string encrFilePath)
        {
            bool manageSpp = false;

            if (encrFilePath == null)
            {
                encrFilePath = GetPSPath(version);
                manageSpp = true;
                KillSPP(version);
            }

            if (string.IsNullOrEmpty(encrFilePath) || !File.Exists(encrFilePath))
            {
                throw new FileNotFoundException("Store does not exist at expected path '" + encrFilePath + "'.");
            }

            try
            {
                using (FileStream fs = File.Open(encrFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    byte[] encrData = fs.ReadAllBytes();
                    File.WriteAllBytes(filePath, PhysStoreCrypto.DecryptPhysicalStore(encrData, production, version));
                }
                Logger.WriteLine("Store dumped successfully to '" + filePath + "'.");
            }
            finally
            {
                if (manageSpp)
                {
                    RestartSPP(version);
                }
            }
        }

        public static void LoadStore(PSVersion version, bool production, string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException("Store file '" + filePath + "' does not exist.");
            }

            KillSPP(version);

            using (IPhysicalStore store = GetStore(version, production))
            {
                store.WriteRaw(File.ReadAllBytes(filePath));
            }

            RestartSPP(version);

            Logger.WriteLine("Loaded store file successfully.");
        }
    }
}


// SPP/SPSys.cs
namespace LibTSforge.SPP
{
    using Microsoft.Win32.SafeHandles;
    using System;
    using System.Runtime.InteropServices;

    public class SPSys
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CreateFile(string lpFileName, uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes, uint dwCreationDisposition, uint dwFlagsAndAttributes, IntPtr hTemplateFile);
        private static SafeFileHandle CreateFileSafe(string device)
        {
            return new SafeFileHandle(CreateFile(device, 0xC0000000, 0, IntPtr.Zero, 3, 0, IntPtr.Zero), true);
        }

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool DeviceIoControl([In] SafeFileHandle hDevice, [In] uint dwIoControlCode, [In] IntPtr lpInBuffer, [In] int nInBufferSize, [Out] IntPtr lpOutBuffer, [In] int nOutBufferSize, out int lpBytesReturned, [In] IntPtr lpOverlapped);

        public static bool IsSpSysRunning()
        {
            SafeFileHandle file = CreateFileSafe(@"\\.\SpDevice");
            IntPtr buffer = Marshal.AllocHGlobal(1);
            int bytesReturned;
            DeviceIoControl(file, 0x80006008, IntPtr.Zero, 0, buffer, 1, out bytesReturned, IntPtr.Zero);
            bool running = Marshal.ReadByte(buffer) != 0;
            Marshal.FreeHGlobal(buffer);
            file.Close();
            return running;
        }

        public static int ControlSpSys(bool start)
        {
            SafeFileHandle file = CreateFileSafe(@"\\.\SpDevice");
            IntPtr buffer = Marshal.AllocHGlobal(4);
            int bytesReturned;
            DeviceIoControl(file, start ? 0x8000a000 : 0x8000a004, IntPtr.Zero, 0, buffer, 4, out bytesReturned, IntPtr.Zero);
            int result = Marshal.ReadInt32(buffer);
            Marshal.FreeHGlobal(buffer);
            file.Close();
            return result;
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
            return Enumerable.SequenceEqual(signature, HMACSign(key, data));
        }

        public static byte[] SaltSHASum(byte[] salt, byte[] data)
        {
            SHA1 sha1 = SHA1.Create();
            byte[] sha_data = salt.Concat(data).ToArray();
            return sha1.ComputeHash(sha_data);
        }

        public static bool SaltSHAVerify(byte[] salt, byte[] data, byte[] checksum)
        {
            return Enumerable.SequenceEqual(checksum, SaltSHASum(salt, data));
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
        public static byte[] DecryptPhysicalStore(byte[] data, bool production, PSVersion version)
        {
            byte[] rsaKey = production ? Keys.PRODUCTION : Keys.TEST;
            BinaryReader br = new BinaryReader(new MemoryStream(data));
            br.BaseStream.Seek(0x10, SeekOrigin.Begin);
            byte[] aesKeySig = br.ReadBytes(0x80);
            byte[] encAesKey = br.ReadBytes(0x80);

            if (!CryptoUtils.RSAVerifySignature(rsaKey, encAesKey, aesKeySig))
            {
                throw new Exception("Failed to decrypt physical store.");
            }

            byte[] aesKey = CryptoUtils.RSADecrypt(rsaKey, encAesKey);
            byte[] decData = CryptoUtils.AESDecrypt(br.ReadBytes((int)br.BaseStream.Length - 0x110), aesKey);
            byte[] hmacKey = decData.Take(0x10).ToArray(); // SHA-1 salt on Vista
            byte[] hmacSig = decData.Skip(0x10).Take(0x14).ToArray(); // SHA-1 hash on Vista
            byte[] psData = decData.Skip(0x28).ToArray();

            if (version != PSVersion.Vista)
            {
                if (!CryptoUtils.HMACVerify(hmacKey, psData, hmacSig))
                {
                    throw new InvalidDataException("Failed to verify HMAC. Physical store is corrupt.");
                }
            }
            else
            {
                if (!CryptoUtils.SaltSHAVerify(hmacKey, psData, hmacSig))
                {
                    throw new InvalidDataException("Failed to verify checksum. Physical store is corrupt.");
                }
            }

            return psData;
        }

        public static byte[] EncryptPhysicalStore(byte[] data, bool production, PSVersion version)
        {
            Dictionary<PSVersion, int> versionTable = new Dictionary<PSVersion, int>
            {
                {PSVersion.Vista, 2},
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
            byte[] hmacSig = version != PSVersion.Vista ? CryptoUtils.HMACSign(hmacKey, data) : CryptoUtils.SaltSHASum(hmacKey, data);

            byte[] decData = { };
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
    using PhysicalStore;
    using SPP;
    using TokenStore;

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
            if (version == PSVersion.Vista) throw new NotSupportedException("This feature is not supported on Windows Vista/Server 2008.");
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
                Logger.WriteLine(string.Format("Installing generated product key {0} status {1:X}", pkey, status));

                if ((int)status < 0)
                {
                    throw new ApplicationException("Failed to install generated product key.");
                }

                Logger.WriteLine("Successfully deposited generated product key.");
                return;
            }

            Logger.WriteLine("Key range is PKEY2005, creating fake key data...");

            if (pkey.Channel == "Volume:GVLK" && version == PSVersion.Win7) throw new NotSupportedException("Fake GVLK generation is not supported on Windows 7.");

            VariableBag pkb = new VariableBag(version);
            pkb.Blocks.AddRange(new[]
            {
                new CRCBlockModern
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingProductKey",
                    ValueAsStr = pkey.ToString()
                },
                new CRCBlockModern
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingMPC",
                    ValueAsStr = pkey.GetMPC()
                },
                new CRCBlockModern {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingPid2",
                    ValueAsStr = pkey.GetPid2()
                },
                new CRCBlockModern
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingPid3",
                    Value = pkey.GetPid3()
                },
                new CRCBlockModern
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = "SppPkeyBindingPid4",
                    Value = pkey.GetPid4()
                },
                new CRCBlockModern
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyChannelId",
                    ValueAsStr = pkey.Channel
                },
                new CRCBlockModern
                {
                    DataType = CRCBlockType.STRING,
                    KeyAsStr = "SppPkeyBindingEditionId",
                    ValueAsStr = pkey.Edition
                },
                new CRCBlockModern
                {
                    DataType = CRCBlockType.BINARY,
                    KeyAsStr = (version == PSVersion.Win7) ? "SppPkeyShortAuthenticator" : "SppPkeyPhoneActivationData",
                    Value = pkey.GetPhoneData(version)
                },
                new CRCBlockModern
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

            SPPUtils.KillSPP(version);

            using (IPhysicalStore ps = SPPUtils.GetStore(version, production))
            {
                using (ITokenStore tks = SPPUtils.GetTokenStore(version))
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

                    string skuMetaName = actId + metSuffix;
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

                    string cachePath = SPPUtils.GetTokensPath(version).Replace("tokens.dat", @"cache\cache.dat");
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
    using System.Collections.Generic;
    using System.Linq;
    using PhysicalStore;
    using SPP;

    public static class GracePeriodReset
    {
        public static void Reset(PSVersion version, bool production)
        {
            SPPUtils.KillSPP(version);
            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
            {
                string value = "msft:sl/timer";
                List<PSBlock> blocks = store.FindBlocks(value).ToList();

                foreach (PSBlock block in blocks)
                {
                    store.DeleteBlock(block.KeyAsStr, block.ValueAsStr);
                }
            }

            SPPUtils.RestartSPP(version);
            Logger.WriteLine("Successfully reset all grace and evaluation period timers.");
        }
    }
}


// Modifiers/KeyChangeLockDelete.cs
namespace LibTSforge.Modifiers
{
    using System.Collections.Generic;
    using System.Linq;
    using PhysicalStore;
    using SPP;
    using System;

    public static class KeyChangeLockDelete
    {
        public static void Delete(PSVersion version, bool production)
        {
            if (version == PSVersion.Vista) throw new NotSupportedException("This feature is not supported on Windows Vista/Server 2008.");

            SPPUtils.KillSPP(version);
            Logger.WriteLine("Writing TrustedStore data...");
            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
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
    using PhysicalStore;
    using SPP;

    public static class KMSHostCharge
    {
        public static void Charge(PSVersion version, bool production, Guid actId)
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

            byte[] cmidGuids = { };
            byte[] reqCounts = { };
            byte[] kmsChargeData = { };

            BinaryWriter writer = new BinaryWriter(new MemoryStream());

            if (version == PSVersion.Vista)
            {
                writer.Write(new byte[44]);
                writer.Seek(0, SeekOrigin.Begin);

                writer.Write(totalClients);
                writer.Write(43200);
                writer.Write(32);

                writer.Seek(20, SeekOrigin.Begin);
                writer.Write((byte)currClients);

                writer.Seek(32, SeekOrigin.Begin);
                writer.Write((byte)currClients);

                writer.Seek(0, SeekOrigin.End);

                for (int i = 0; i < currClients; i++)
                {
                    writer.Write(Guid.NewGuid().ToByteArray());
                    writer.Write(ldapTimestamp - (10 * (i + 1)));
                }

                kmsChargeData = writer.GetBytes();
            } 
            else
            {
                for (int i = 0; i < currClients; i++)
                {
                    writer.Write(ldapTimestamp - (10 * (i + 1)));
                    writer.Write(Guid.NewGuid().ToByteArray());
                }

                cmidGuids = writer.GetBytes();

                writer = new BinaryWriter(new MemoryStream());

                writer.Write(new byte[40]);

                writer.Seek(4, SeekOrigin.Begin);
                writer.Write((byte)currClients);

                writer.Seek(24, SeekOrigin.Begin);
                writer.Write((byte)currClients);

                reqCounts = writer.GetBytes();
            }

            SPPUtils.KillSPP(version);

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
            {
                if (version != PSVersion.Vista)
                {
                    VariableBag kmsCountData = new VariableBag(version);
                    kmsCountData.Blocks.AddRange(new[]
                    {
                        new CRCBlockModern
                        {
                            DataType = CRCBlockType.BINARY,
                            KeyAsStr = "SppBindingLicenseData",
                            Value = hwidBlock
                        },
                        new CRCBlockModern
                        {
                            DataType = CRCBlockType.UINT,
                            Key = new byte[] { },
                            ValueAsInt = (uint)totalClients
                        },
                        new CRCBlockModern
                        {
                            DataType = CRCBlockType.UINT,
                            Key = new byte[] { },
                            ValueAsInt = 1051200000
                        },
                        new CRCBlockModern
                        {
                            DataType = CRCBlockType.UINT,
                            Key = new byte[] { },
                            ValueAsInt = (uint)currClients
                        },
                        new CRCBlockModern
                        {
                            DataType = CRCBlockType.BINARY,
                            Key = new byte[] { },
                            Value = cmidGuids
                        },
                        new CRCBlockModern
                        {
                            DataType = CRCBlockType.BINARY,
                            Key = new byte[] { },
                            Value = reqCounts
                        }
                    });

                    kmsChargeData = kmsCountData.Serialize();
                }

                string countVal = version == PSVersion.Vista ? "C8F6FFF1-79CE-404C-B150-F97991273DF1" : string.Format("msft:spp/kms/host/2.0/store/counters/{0}", appId);

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

            SPPUtils.RestartSPP(version);
        }
    }
}


// Modifiers/RearmReset.cs
namespace LibTSforge.Modifiers
{
    using System.Collections.Generic;
    using System.Linq;
    using PhysicalStore;
    using SPP;

    public static class RearmReset
    {
        public static void Reset(PSVersion version, bool production)
        {
            SPPUtils.KillSPP(version);

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
            {
                List<PSBlock> blocks;

                if (version == PSVersion.Vista)
                {
                    blocks = store.FindBlocks("740D70D8-6448-4b2f-9063-4A7A463600C5").ToList();
                }
                else if (version == PSVersion.Win7)
                {
                    blocks = store.FindBlocks(0xA0000).ToList();
                }
                else
                {
                    blocks = store.FindBlocks("__##USERSEP-RESERVED##__$$REARM-COUNT$$").ToList();
                }

                foreach (PSBlock block in blocks)
                {
                    if (version == PSVersion.Vista)
                    {
                        store.DeleteBlock(block.KeyAsStr, block.ValueAsStr);
                    }
                    else if (version == PSVersion.Win7)
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


// Modifiers/SetIIDParams.cs
namespace LibTSforge.Modifiers
{
    using PhysicalStore;
    using SPP;
    using System.IO;
    using System;

    public static class SetIIDParams
    {
        public static void SetParams(PSVersion version, bool production, Guid actId, PKeyAlgorithm algorithm, int group, int serial, ulong security)
        {
            if (version == PSVersion.Vista) throw new NotSupportedException("This feature is not supported on Windows Vista/Server 2008.");

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

            Guid pkeyId = SLApi.GetInstalledPkeyID(actId);

            SPPUtils.KillSPP(version);

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
            {
                string key = string.Format("SPPSVC\\{0}\\{1}", appId, actId);
                PSBlock keyBlock = store.GetBlock(key, pkeyId.ToString());

                if (keyBlock == null)
                {
                    throw new InvalidDataException("Failed to get product key data for activation ID " + actId + ".");
                }

                VariableBag pkb = new VariableBag(keyBlock.Data, version);

                ProductKey pkey = new ProductKey
                {
                    Group = group,
                    Serial = serial,
                    Security = security,
                    Algorithm = algorithm,
                    Upgrade = false
                };

                string blockName = version == PSVersion.Win7 ? "SppPkeyShortAuthenticator" : "SppPkeyPhoneActivationData";
                pkb.SetBlock(blockName, pkey.GetPhoneData(version));
                store.SetBlock(key, pkeyId.ToString(), pkb.Serialize());
            }

            Logger.WriteLine("Successfully set IID parameters.");
        }
    }
}


// Modifiers/TamperedFlagsDelete.cs
namespace LibTSforge.Modifiers
{
    using System.Linq;
    using PhysicalStore;
    using SPP;

    public static class TamperedFlagsDelete
    {
        public static void DeleteTamperFlags(PSVersion version, bool production)
        {
            SPPUtils.KillSPP(version);

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
            {
                if (version == PSVersion.Vista)
                {
                    DeleteFlag(store, "6BE8425B-E3CF-4e86-A6AF-5863E3DCB606");
                }
                else if (version == PSVersion.Win7)
                {
                    SetFlag(store, 0xA0001);
                }
                else
                {
                    DeleteFlag(store, "__##USERSEP-RESERVED##__$$RECREATED-FLAG$$");
                    DeleteFlag(store, "__##USERSEP-RESERVED##__$$RECOVERED-FLAG$$");
                }

                Logger.WriteLine("Successfully cleared the tamper state.");
            }

            SPPUtils.RestartSPP(version);
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
    using PhysicalStore;
    using SPP;

    public static class UniqueIdDelete
    {
        public static void DeleteUniqueId(PSVersion version, bool production, Guid actId)
        {
            if (version == PSVersion.Vista) throw new NotSupportedException("This feature is not supported on Windows Vista/Server 2008.");

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

            Guid pkeyId = SLApi.GetInstalledPkeyID(actId);

            SPPUtils.KillSPP(version);

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
            {
                string key = string.Format("SPPSVC\\{0}\\{1}", appId, actId);
                PSBlock keyBlock = store.GetBlock(key, pkeyId.ToString());

                if (keyBlock == null)
                {
                    throw new Exception("No product key found.");
                }

                VariableBag pkb = new VariableBag(keyBlock.Data, version);

                pkb.DeleteBlock("SppPkeyUniqueIdToken");

                store.SetBlock(key, pkeyId.ToString(), pkb.Serialize());
            }

            Logger.WriteLine("Successfully removed Unique ID for product key ID " + pkeyId);
        }
    }
}


// Activators/KMS4K.cs
namespace LibTSforge.Activators
{
    using System;
    using System.IO;
    using PhysicalStore;
    using SPP;

    public class KMS4k
    {
        public static void Activate(PSVersion version, bool production, Guid actId)
        {
            Guid appId;
            if (actId == Guid.Empty)
            {
                appId = SLApi.WINDOWS_APP_ID;
                actId = SLApi.GetDefaultActivationID(appId, true);

                if (actId == Guid.Empty)
                {
                    throw new NotSupportedException("No applicable activation IDs found.");
                }
            }
            else
            {
                appId = SLApi.GetAppId(actId);
            }

            if (SLApi.GetPKeyChannel(SLApi.GetInstalledPkeyID(actId)) != "Volume:GVLK")
            {
                throw new NotSupportedException("Non-Volume:GVLK product key installed.");
            }

            SPPUtils.KillSPP(version);

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
            {
                string key = string.Format("SPPSVC\\{0}\\{1}", appId, actId);

                ulong unknown = 0;
                ulong time1;
                ulong time2 = (ulong)DateTime.UtcNow.ToFileTime();
                ulong expiry = Constants.TimerMax;

                if (version == PSVersion.Vista || version == PSVersion.Win7)
                {
                    unknown = 0x800000000;
                    time1 = 0;
                }
                else
                {
                    long creationTime = BitConverter.ToInt64(store.GetBlock("__##USERSEP##\\$$_RESERVED_$$\\NAMESPACE__", "__##USERSEP-RESERVED##__$$GLOBAL-CREATION-TIME$$").Data, 0);
                    long tickCount = BitConverter.ToInt64(store.GetBlock("__##USERSEP##\\$$_RESERVED_$$\\NAMESPACE__", "__##USERSEP-RESERVED##__$$GLOBAL-TICKCOUNT-UPTIME$$").Data, 0);
                    long deltaTime = BitConverter.ToInt64(store.GetBlock(key, "__##USERSEP-RESERVED##__$$UP-TIME-DELTA$$").Data, 0);

                    time1 = (ulong)(creationTime + tickCount + deltaTime);
                    time2 /= 10000;
                    expiry /= 10000;
                }

                if (version == PSVersion.Vista)
                {
                    VistaTimer vistaTimer = new VistaTimer
                    {
                        Time = time2,
                        Expiry = Constants.TimerMax
                    };

                    string vistaTimerName = string.Format("msft:sl/timer/VLExpiration/VOLUME/{0}/{1}", appId, actId);

                    store.DeleteBlock(key, vistaTimerName);
                    store.DeleteBlock(key, actId.ToString());

                    BinaryWriter writer = new BinaryWriter(new MemoryStream());
                    writer.Write(Constants.KMSv4Response.Length);
                    writer.Write(Constants.KMSv4Response);
                    writer.Write(Constants.UniversalHWIDBlock);
                    byte[] kmsData = writer.GetBytes();

                    store.AddBlocks(new[]
                    {
                        new PSBlock
                        {
                            Type = BlockType.TIMER,
                            Flags = 0,
                            KeyAsStr = key,
                            ValueAsStr = vistaTimerName,
                            Data = vistaTimer.CastToArray()
                        },
                        new PSBlock
                        {
                            Type = BlockType.NAMED,
                            Flags = 0,
                            KeyAsStr = key,
                            ValueAsStr = actId.ToString(),
                            Data = kmsData
                        }
                    });
                }
                else
                {
                    byte[] hwidBlock = Constants.UniversalHWIDBlock;
                    byte[] kmsResp;

                    switch (version)
                    {
                        case PSVersion.Win7:
                            kmsResp = Constants.KMSv4Response;
                            break;
                        case PSVersion.Win8:
                            kmsResp = Constants.KMSv5Response;
                            break;
                        case PSVersion.WinBlue:
                        case PSVersion.WinModern:
                            kmsResp = Constants.KMSv6Response;
                            break;
                        default:
                            throw new NotSupportedException("Unsupported PSVersion.");
                    }

                    VariableBag kmsBinding = new VariableBag(version);

                    kmsBinding.Blocks.AddRange(new[]
                    {
                    new CRCBlockModern
                    {
                        DataType = CRCBlockType.BINARY,
                        Key = new byte[] { },
                        Value = kmsResp
                    },
                    new CRCBlockModern
                    {
                        DataType = CRCBlockType.STRING,
                        Key = new byte[] { },
                        ValueAsStr = "msft:rm/algorithm/hwid/4.0"
                    },
                    new CRCBlockModern
                    {
                        DataType = CRCBlockType.BINARY,
                        KeyAsStr = "SppBindingLicenseData",
                        Value = hwidBlock
                    }
                    });

                    if (version == PSVersion.WinModern)
                    {
                        kmsBinding.Blocks.AddRange(new[]
                        {
                        new CRCBlockModern
                        {
                            DataType = CRCBlockType.STRING,
                            Key = new byte[] { },
                            ValueAsStr = "massgrave.dev"
                        },
                        new CRCBlockModern
                        {
                            DataType = CRCBlockType.STRING,
                            Key = new byte[] { },
                            ValueAsStr = "6969"
                        }
                        });
                    }

                    byte[] kmsBindingData = kmsBinding.Serialize();

                    Timer kmsTimer = new Timer
                    {
                        Unknown = unknown,
                        Time1 = time1,
                        Time2 = time2,
                        Expiry = expiry
                    };

                    string storeVal = string.Format("msft:spp/kms/bind/2.0/store/{0}/{1}", appId, actId);
                    string timerVal = string.Format("msft:spp/kms/bind/2.0/timer/{0}/{1}", appId, actId);

                    store.DeleteBlock(key, storeVal);
                    store.DeleteBlock(key, timerVal);

                    store.AddBlocks(new[]
                    {
                    new PSBlock
                    {
                        Type = BlockType.NAMED,
                        Flags = (version == PSVersion.WinModern) ? (uint)0x400 : 0,
                        KeyAsStr = key,
                        ValueAsStr = storeVal,
                        Data = kmsBindingData
                    },
                    new PSBlock
                    {
                        Type = BlockType.TIMER,
                        Flags = (version == PSVersion.Win7) ? (uint)0 : 0x4,
                        KeyAsStr = key,
                        ValueAsStr = timerVal,
                        Data = kmsTimer.CastToArray()
                    }
                    });
                }
            }

            SPPUtils.RestartSPP(version);
            SLApi.FireStateChangedEvent(appId);
            Logger.WriteLine("Activated using KMS4k successfully.");
        }
    }
}


// Activators/ZeroCID.cs
namespace LibTSforge.Activators
{
    using System;
    using System.IO;
    using System.Linq;
    using Crypto;
    using PhysicalStore;
    using SPP;

    public static class ZeroCID
    {
        private static void Deposit(Guid actId, string instId)
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

            if (version == PSVersion.Vista || version == PSVersion.Win7)
            {
                Deposit(actId, instId);
            }

            SPPUtils.KillSPP(version);

            Logger.WriteLine("Writing TrustedStore data...");

            using (IPhysicalStore store = SPPUtils.GetStore(version, production))
            {
                byte[] hwidBlock = Constants.UniversalHWIDBlock;

                Logger.WriteLine("Activation ID: " + actId);
                Logger.WriteLine("Installation ID: " + instId);
                Logger.WriteLine("Product Key ID: " + pkeyId);

                byte[] iidHash;

                if (version == PSVersion.Vista)
                {
                    iidHash = CryptoUtils.SHA256Hash(Utils.EncodeString(instId)).Take(0x10).ToArray();
                }
                else if (version == PSVersion.Win7)
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

                VariableBag pkb = new VariableBag(keyBlock.Data, version);

                byte[] pkeyData;

                if (version == PSVersion.Vista)
                {
                    pkeyData = pkb.GetBlock("PKeyBasicInfo").Value;
                    string uniqueId = Utils.DecodeString(pkeyData.Skip(0x120).Take(0x80).ToArray());
                    string extPid = Utils.DecodeString(pkeyData.Skip(0x1A0).Take(0x80).ToArray());

                    uint group;
                    uint.TryParse(extPid.Split('-')[1], out group);

                    if (group == 0)
                    {
                        throw new FormatException("Extended PID has invalid format.");
                    }

                    ulong shortauth;

                    try
                    {
                        shortauth = BitConverter.ToUInt64(Convert.FromBase64String(uniqueId.Split('&')[1]), 0);
                    } 
                    catch
                    {
                        throw new FormatException("Key Unique ID has invalid format.");
                    }

                    shortauth |= (ulong)group << 41;
                    pkeyData = BitConverter.GetBytes(shortauth);
                }
                else if (version == PSVersion.Win7)
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
                writer.Write(iidHash.Length);
                writer.Write(iidHash);
                writer.Write(hwidBlock.Length);
                writer.Write(hwidBlock);
                byte[] tsHwidData = writer.GetBytes();

                writer = new BinaryWriter(new MemoryStream());
                writer.Write(iidHash.Length);
                writer.Write(iidHash);
                writer.Write(pkeyData.Length);
                writer.Write(pkeyData);
                byte[] tsPkeyInfoData = writer.GetBytes();

                string phoneVersion = version == PSVersion.Vista ? "6.0" : "7.0";
                Guid indexSlid = version == PSVersion.Vista ? actId : pkeyId;
                string hwidBlockName = string.Format("msft:Windows/{0}/Phone/Cached/HwidBlock/{1}", phoneVersion, indexSlid);
                string pkeyInfoName = string.Format("msft:Windows/{0}/Phone/Cached/PKeyInfo/{1}", phoneVersion, indexSlid);

                store.DeleteBlock(key, hwidBlockName);
                store.DeleteBlock(key, pkeyInfoName);

                store.AddBlocks(new[] {
                    new PSBlock
                    {
                        Type = BlockType.NAMED,
                        Flags = 0,
                        KeyAsStr = key,
                        ValueAsStr = hwidBlockName,
                        Data = tsHwidData
                    }, 
                    new PSBlock
                    {
                        Type = BlockType.NAMED,
                        Flags = 0,
                        KeyAsStr = key,
                        ValueAsStr = pkeyInfoName,
                        Data = tsPkeyInfoData
                    }
                });
            }

            if (version != PSVersion.Vista && version != PSVersion.Win7)
            {
                Deposit(actId, instId);
            }

            SPPUtils.RestartSPP(version);
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
        public readonly Dictionary<string, string> Data = new Dictionary<string, string>();

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

        private void Deserialize(byte[] data)
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
    using Crypto;

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
        private readonly FileStream TokensFile;

        public void Deserialize()
        {
            if (TokensFile.Length < BLOCK_SIZE) return;

            TokensFile.Seek(0x24, SeekOrigin.Begin);
            uint nextBlock;

            BinaryReader reader = new BinaryReader(TokensFile);
            do
            {
                reader.ReadUInt32();
                nextBlock = reader.ReadUInt32();

                for (int i = 0; i < ENTRIES_PER_BLOCK; i++)
                {
                    uint curOffset = reader.ReadUInt32();
                    bool populated = reader.ReadUInt32() == 1;
                    uint contentOffset = reader.ReadUInt32();
                    uint contentLength = reader.ReadUInt32();
                    uint allocLength = reader.ReadUInt32();
                    byte[] contentData = { };

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
                        byte[] blockData = new byte[BLOCK_SIZE - 0x20];

                        tokens.Read(blockData, 0, BLOCK_SIZE - 0x20);
                        byte[] blockHash = CryptoUtils.SHA256Hash(blockData);

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

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct VistaTimer
    {
        public ulong Time;
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
    using Crypto;

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
        private byte[] PreHeaderBytes = { };
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
            if (!Data.ContainsKey(key))
            {
                return;
            }

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

        public void DeleteBlock(string key, uint value)
        {
            if (!Data.ContainsKey(key))
            {
                return;
            }

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

        public PhysicalStoreModern(string tsPath, bool production, PSVersion version)
        {
            TSFile = File.Open(tsPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            Deserialize(PhysStoreCrypto.DecryptPhysicalStore(TSFile.ReadAllBytes(), production, version));
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
            byte[] data = PhysStoreCrypto.DecryptPhysicalStore(TSFile.ReadAllBytes(), Production, Version);
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


// PhysicalStore/PhysicalStoreVista.cs
namespace LibTSforge.PhysicalStore
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using Crypto;

    public class VistaBlock
    {
        public BlockType Type;
        public uint Flags;
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
            writer.Write(Value.Length);
            writer.Write(Data.Length);
            writer.Write(Value);
            writer.Write(Data);
        }

        internal static VistaBlock Decode(BinaryReader reader)
        {
            uint type = reader.ReadUInt32();
            uint flags = reader.ReadUInt32();

            int valueLen = reader.ReadInt32();
            int dataLen = reader.ReadInt32();

            byte[] value = reader.ReadBytes(valueLen);
            byte[] data = reader.ReadBytes(dataLen);
            return new VistaBlock
            {
                Type = (BlockType)type,
                Flags = flags,
                Value = value,
                Data = data,
            };
        }
    }

    public sealed class PhysicalStoreVista : IPhysicalStore
    {
        private byte[] PreHeaderBytes = { };
        private readonly List<VistaBlock> Blocks = new List<VistaBlock>();
        private readonly FileStream TSPrimary;
        private readonly FileStream TSSecondary;
        private readonly bool Production;

        public byte[] Serialize()
        {
            BinaryWriter writer = new BinaryWriter(new MemoryStream());
            writer.Write(PreHeaderBytes);

            foreach (VistaBlock block in Blocks)
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
                Blocks.Add(VistaBlock.Decode(reader));
                reader.Align(4);
            }
        }

        public void AddBlock(PSBlock block)
        {
            Blocks.Add(new VistaBlock
            {
                Type = block.Type,
                Flags = block.Flags,
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
            foreach (VistaBlock block in Blocks)
            {
                if (block.ValueAsStr == value)
                {
                    return new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = new byte[0],
                        Value = block.Value,
                        Data = block.Data
                    };
                }
            }

            return null;
        }

        public PSBlock GetBlock(string key, uint value)
        {
            foreach (VistaBlock block in Blocks)
            {
                if (block.ValueAsInt == value)
                {
                    return new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = new byte[0],
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
                VistaBlock block = Blocks[i];

                if (block.ValueAsStr == value)
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
                VistaBlock block = Blocks[i];

                if (block.ValueAsInt == value)
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
            foreach (VistaBlock block in Blocks)
            {
                if (block.ValueAsStr == value)
                {
                    Blocks.Remove(block);
                    return;
                }
            }
        }

        public void DeleteBlock(string key, uint value)
        {
            foreach (VistaBlock block in Blocks)
            {
                if (block.ValueAsInt == value)
                {
                    Blocks.Remove(block);
                    return;
                }
            }
        }

        public PhysicalStoreVista(string primaryPath, bool production)
        {
            TSPrimary = File.Open(primaryPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            TSSecondary = File.Open(primaryPath.Replace("-0.", "-1."), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            Production = production;

            Deserialize(PhysStoreCrypto.DecryptPhysicalStore(TSPrimary.ReadAllBytes(), production, PSVersion.Vista));
            TSPrimary.Seek(0, SeekOrigin.Begin);
        }

        public void Dispose()
        {
            if (TSPrimary.CanWrite && TSSecondary.CanWrite)
            {
                byte[] data = PhysStoreCrypto.EncryptPhysicalStore(Serialize(), Production, PSVersion.Vista);

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
            byte[] data = PhysStoreCrypto.DecryptPhysicalStore(TSPrimary.ReadAllBytes(), Production, PSVersion.Vista);
            TSPrimary.Seek(0, SeekOrigin.Begin);
            return data;
        }

        public void WriteRaw(byte[] data)
        {
            byte[] encrData = PhysStoreCrypto.EncryptPhysicalStore(data, Production, PSVersion.Vista);

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

            foreach (VistaBlock block in Blocks)
            {
                if (block.ValueAsStr.Contains(valueSearch))
                {
                    results.Add(new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = new byte[0],
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

            foreach (VistaBlock block in Blocks)
            {
                if (block.ValueAsInt == valueSearch)
                {
                    results.Add(new PSBlock
                    {
                        Type = block.Type,
                        Flags = block.Flags,
                        Key = new byte[0],
                        Value = block.Value,
                        Data = block.Data
                    });
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
    using Crypto;

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
        private byte[] PreHeaderBytes = { };
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

            Deserialize(PhysStoreCrypto.DecryptPhysicalStore(TSPrimary.ReadAllBytes(), production, PSVersion.Win7));
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
            byte[] data = PhysStoreCrypto.DecryptPhysicalStore(TSPrimary.ReadAllBytes(), Production, PSVersion.Win7);
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

    public abstract class CRCBlock
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

        public abstract void Encode(BinaryWriter writer);
        public abstract void Decode(BinaryReader reader);
        public abstract uint CRC();
    }

    public class CRCBlockVista : CRCBlock
    {
        public override void Encode(BinaryWriter writer)
        {
            uint crc = CRC();
            writer.Write((uint)DataType);
            writer.Write(0);
            writer.Write(Key.Length);
            writer.Write(Value.Length);
            writer.Write(crc);

            writer.Write(Key);

            writer.Write(Value);
        }

        public override void Decode(BinaryReader reader)
        {
            uint type = reader.ReadUInt32();
            reader.ReadUInt32();
            uint lenName = reader.ReadUInt32();
            uint lenVal = reader.ReadUInt32();
            uint crc = reader.ReadUInt32();

            byte[] key = reader.ReadBytes((int)lenName);
            byte[] value = reader.ReadBytes((int)lenVal);

            DataType = (CRCBlockType)type;
            Key = key;
            Value = value;

            if (CRC() != crc)
            {
                throw new InvalidDataException("Invalid CRC in variable bag.");
            }
        }

        public override uint CRC()
        {
            return Utils.CRC32(Value);
        }
    }

    public class CRCBlockModern : CRCBlock
    {
        public override void Encode(BinaryWriter writer)
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

        public override void Decode(BinaryReader reader)
        {
            uint crc = reader.ReadUInt32();
            uint type = reader.ReadUInt32();
            uint lenName = reader.ReadUInt32();
            uint lenVal = reader.ReadUInt32();

            byte[] key = reader.ReadBytes((int)lenName);
            reader.Align(8);

            byte[] value = reader.ReadBytes((int)lenVal);
            reader.Align(8);

            DataType = (CRCBlockType)type;
            Key = key;
            Value = value;

            if (CRC() != crc)
            {
                throw new InvalidDataException("Invalid CRC in variable bag.");
            }
        }

        public override uint CRC()
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
        private readonly PSVersion Version;

        private void Deserialize(byte[] data)
        {
            int len = data.Length;

            BinaryReader reader = new BinaryReader(new MemoryStream(data));

            while (reader.BaseStream.Position < len - 0x10)
            {
                CRCBlock block;

                if (Version == PSVersion.Vista)
                {
                    block = new CRCBlockVista();
                }
                else
                {
                    block = new CRCBlockModern();
                }

                block.Decode(reader);
                Blocks.Add(block);
            }
        }

        public byte[] Serialize()
        {
            BinaryWriter writer = new BinaryWriter(new MemoryStream());

            foreach (CRCBlock block in Blocks)
            {
                if (Version == PSVersion.Vista)
                {
                    ((CRCBlockVista)block).Encode(writer);
                } else
                {
                    ((CRCBlockModern)block).Encode(writer);
                }
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

        public VariableBag(byte[] data, PSVersion version)
        {
            Version = version;
            Deserialize(data);
        }

        public VariableBag(PSVersion version)
        {
            Version = version;
        }
    }
}
'@
$ErrorActionPreference = 'Stop'
$binPath = "$env:_work\BIN\LibTSforge.dll"
$psMajorVer = (Get-Host).Version.Major

if (Test-Path -LiteralPath $binPath) {
    Write-Host "LibTSforge.dll found in BIN folder. Loading the DLL..."
    Add-Type -Path $binPath
}
else {
    $cp = [CodeDom.Compiler.CompilerParameters] [string[]]@("System.dll", "System.Core.dll", "System.ServiceProcess.dll", "System.Xml.dll", "System.Xml.Linq.dll")
    if ($psMajorVer -le 2) { $cp.CompilerOptions = "/define:POWERSHELL2 /unsafe" } else { $cp.CompilerOptions = "/unsafe" }
    $lang = if ($psMajorVer -gt 2) { "CSharp" } else { "CSharpVersion3" }

    $ctemp = "$env:SystemRoot\Temp\"
    if (-Not (Test-Path -Path $ctemp)) { New-Item -Path $ctemp -ItemType Directory > $null }
    $env:TMP = $ctemp
    $env:TEMP = $ctemp

    $cp.GenerateInMemory = $true
    Add-Type -Language $lang -TypeDefinition $src -CompilerParameters $cp
}

if ($env:_debug -eq '0') {
    [LibTSforge.Logger]::HideOutput = $true
}
$ver = [LibTSforge.Utils]::DetectVersion()
$prod = [LibTSforge.SPP.SPPUtils]::DetectCurrentKey()
$tsactids = @($args)

function Get-WmiInfo {
    param ([string]$tsactid, [string]$property)
    
    $query = "SELECT ID, $property FROM SoftwareLicensingProduct WHERE ID='$tsactid'"
    $record = Get-WmiObject -Query $query
    if ($record) {
        return $record.$property
    }
}

function slGetSkuInfo($SkuId) {
    $t = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0)
    $t.DefinePInvokeMethod('SLOpen', 'slc.dll', 22, 1, [Int32], @([IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
    $t.DefinePInvokeMethod('SLClose', 'slc.dll', 22, 1, [IntPtr], @([IntPtr]), 1, 3).SetImplementationFlags(128)
    $t.DefinePInvokeMethod('SLGetProductSkuInformation', 'slc.dll', 22, 1, [Int32], @([IntPtr], [Guid].MakeByRefType(), [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
    $w = $t.CreateType()
    $hSLC = 0
    try {
        [void]$w::SLOpen([ref]$hSLC)
        $c = 0; $b = 0
        $r = $w::SLGetProductSkuInformation($hSLC, [ref][Guid]$SkuId, "msft:sl/EUL/PHONE/PUBLIC", [ref]$null, [ref]$c, [ref]$b)
        return ($r -eq 0)
    }
    finally {
        [void]$w::SLClose($hSLC)
    }
}

if (-not $env:resetstuff) {
    foreach ($tsactid in $tsactids) {
        try {
            $activated = $null
            $prodDes = Get-WmiInfo -tsactid $tsactid -property "Description"
            $prodName = Get-WmiInfo -tsactid $tsactid -property "Name"
            if ($prodName) {
                $nameParts = $prodName -split ',', 2
                $prodName = if ($nameParts.Count -gt 1) { ($nameParts[1].Trim() -split '[ ,]')[0] } else { $null }
            }
            if (-not $env:_vis -and -not $env:oldks) {
                [LibTSforge.Modifiers.GenPKeyInstall]::InstallGenPKey($ver, $prod, $tsactid)
            }
            if ($prodName -match 'Office' -and $prodDes -notmatch 'KMS' -and -not (slGetSkuInfo($tsactid))) {
                $licenseStatus = Get-WmiInfo -tsactid $tsactid -property "LicenseStatus"
                if ($licenseStatus -eq 1) {
                    Write-Host "[$prodName] is already permanently activated." -ForegroundColor White -BackgroundColor DarkGreen
                    continue
                }
            }
            if ($env:tsmethod -eq "StaticCID") {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $attempts = @(
                    @(100055, 1000043, 1338662172562478),
                    @(1345, 1003020, 6311608238084405)
                )
                foreach ($params in $attempts) {
                    [LibTSforge.Modifiers.SetIIDParams]::SetParams($ver, $prod, $tsactid, [LibTSforge.SPP.PKeyAlgorithm]::PKEY2009, $params[0], $params[1], $params[2])
                    $instId = [LibTSforge.SPP.SLApi]::GetInstallationID($tsactid)
                    $confId = [ActivationWs.ActivationHelper]::CallWebService(1, $instId, "31337-42069-123-456789-04-1337-2600.0000-2542001")
                    $result = [LibTSforge.SPP.SLApi]::DepositConfirmationID($tsactid, $instId, $confId)
                    if ($result -eq 0) { break }
                }
                [LibTSforge.SPP.SPPUtils]::RestartSPP($ver)
            }
            elseif ($env:tsmethod -eq "KMS4k") {
                [LibTSforge.Activators.KMS4k]::Activate($ver, $prod, $tsactid)
            }
            else {
                [LibTSforge.Activators.ZeroCID]::Activate($ver, $prod, $tsactid)
            }
            if ($env:tsmethod -eq "KMS4k") {
                $GracePeriodStatus = Get-WmiInfo -tsactid $tsactid -property "GracePeriodRemaining"
                if ($GracePeriodStatus -gt 259200) { $activated = 1 }
            }
            else {
                $licenseStatus = Get-WmiInfo -tsactid $tsactid -property "LicenseStatus"
                if ($licenseStatus -eq 1) { $activated = 1 }
            }
            if ($activated) {
                if ($prodDes -match 'KMS' -and $prodDes -notmatch 'CLIENT') {
                    [LibTSforge.Modifiers.KMSHostCharge]::Charge($ver, $prod, $tsactid)
                    Write-Host "[$prodName] CSVLK is permanently activated with $env:tsmethod." -ForegroundColor White -BackgroundColor DarkGreen
                    Write-Host "[$prodName] CSVLK is charged with 25 clients for 30 days." -ForegroundColor White -BackgroundColor DarkGreen
                }
                else {
                    if ($env:tsmethod -eq "KMS4k") {
                        Write-Host "[$prodName] is activated till $([DateTime]::Now.AddMinutes($GracePeriodStatus).ToString('yyyy-MM-dd HH:mm:ss')) with $env:tsmethod." -ForegroundColor White -BackgroundColor DarkGreen
                    }
                    else {
                        Write-Host "[$prodName] is permanently activated with $env:tsmethod." -ForegroundColor White -BackgroundColor DarkGreen
                    }
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

if ($env:resetstuff) {
    try {
        if (-not $env:_vis) { [LibTSforge.Modifiers.TamperedFlagsDelete]::DeleteTamperFlags($ver, $prod) }
        [LibTSforge.SPP.SLApi]::RefreshLicenseStatus()
        [LibTSforge.Modifiers.RearmReset]::Reset($ver, $prod)
        [LibTSforge.Modifiers.GracePeriodReset]::Reset($ver, $prod)
        if (-not $env:_vis) { [LibTSforge.Modifiers.KeyChangeLockDelete]::Delete($ver, $prod) }
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

    $filterPreview = $filteredConfigs | Where-Object { $_.ProductDescription -notmatch 'preview|c2r' }

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
::  3rd column = Product ID from branding.xml
::  4th column = Edition
::  5th column = Other Edition IDs if they are part of the same primary product (For reference only)
::  Separator  = "_"

:msiofficedata

for %%# in (
14_4d463c2c-0505-4626-8cdb-a4da82e2d8ed_0015_AccessR
14_745fb377-0a59-4ca9-b9a9-c359557a2c4e_001C_AccessRuntimeR
14_95ab3ec8-4106-4f9d-b632-03c019d1d23f_0015_AccessVL
14_4eaff0d0-c6cb-4187-94f3-c7656d49a0aa_0016_ExcelR_[HSExcelR]
14_71dc86ff-f056-40d0-8ffb-9592705c9b76_0016_ExcelVL
14_7004b7f0-6407-4f45-8eac-966e5f868bde_00BA_GrooveR
14_fdad0dfa-417d-4b4f-93e4-64ea8867b7fd_00BA_GrooveVL
14_7b7d1f17-fdcb-4820-9789-9bec6e377821_0013_HomeBusinessR_[HomeBusinessDemoR]
14_19316117-30a8-4773-8fd9-7f7231f4e060_011E_HomeBusinessSubR
14_09e2d37e-474b-4121-8626-58ad9be5776f_002F_HomeStudentR_[HomeStudentDemoR]
14_ef1da464-01c8-43a6-91af-e4e5713744f9_0044_InfoPathR
14_85e22450-b741-430c-a172-a37962c938af_0044_InfoPathVL
14_14f5946a-debc-4716-babc-7e2c240fec08_000F_MondoR
14_533b656a-4425-480b-8e30-1a2358898350_000F_MondoVL
14_c1ceda8b-c578-4d5d-a4aa-23626be4e234_003D_ProfessionalR_[OEM-SingleImage]Exception
14_3f7aa693-9a7e-44fc-9309-bb3d8e604925_00A1_OneNoteR_[HSOneNoteR]
14_6860b31f-6a67-48b8-84b9-e312b3485c4b_00A1_OneNoteVL
14_fbf4ac36-31c8-4340-8666-79873129cf40_001A_OutlookR
14_a9aeabd8-63b8-4079-a28e-f531807fd6b8_001A_OutlookVL
14_acb51361-c0db-4895-9497-1831c41f31a6_0033_PersonalR_[PersonalDemoR,PersonalPrepaidR]
14_133c8359-4e93-4241-8118-30bb18737ea0_0018_PowerPointR_[HSPowerPointR]
14_38252940-718c-4aa6-81a4-135398e53851_0018_PowerPointVL
14_8b559c37-0117-413e-921b-b853aeb6e210_0014_ProfessionalR_[ProfessionalAcadR,ProfessionalDemoR]
14_725714d7-d58f-4d12-9fa8-35873c6f7215_003B_ProjectProR_[ProjectProMSDNR]
14_4d06f72e-fd50-4bc2-a24b-d448d7f17ef2_011F_ProjectProSubR
14_1cf57a59-c532-4e56-9a7d-ffa2fe94b474_003B_ProjectProVL
14_688f6589-2bd9-424e-a152-b13f36aa6de1_003A_ProjectStdR
14_11b39439-6b93-4642-9570-f2eb81be2238_003A_ProjectStdVL
14_71af7e84-93e6-4363-9b69-699e04e74071_0011_ProPlusR_[ProPlusAcadR,ProPlusMSDNR,Sub4R]
14_e98ef0c0-71c4-42ce-8305-287d8721e26c_011D_ProPlusSubR
14_fdf3ecb9-b56f-43b2-a9b8-1b48b6bae1a7_0011_ProPlusVL_[ProPlusAcadVL]
14_98677603-a668-4fa4-9980-3f1f05f78f69_0019_PublisherR
14_3d014759-b128-4466-9018-e80f6320d9d0_0019_PublisherVL
14_dbe3aee0-5183-4ff7-8142-66050173cb01_008B_SmallBusBasicsR_[SmallBusBasicsMSDNR]
14_8090771e-d41a-4482-929e-de87f1f47e46_008B_SmallBusBasicsVL
14_b78df69e-0966-40b1-ae85-30a5134dedd0_0017_SPDR
14_d3422cfb-8d8b-4ead-99f9-eab0ccd990d7_0012_StandardR
14_1f76e346-e0be-49bc-9954-70ec53a4fcfe_0012_StandardVL_[StandardAcadVL]
14_2745e581-565a-4670-ae90-6bf7c57ffe43_0066_StarterR
14_66cad568-c2dc-459d-93ec-2f3cb967ee34_0057_VisioSIR_Prem[Pro,Std]Exception
14_36756cb8-8e69-4d11-9522-68899507cd6a_0057_VisioSIVL_Prem[Pro,Std]Exception
14_db3bbc9c-ce52-41d1-a46f-1a1d68059119_001B_WordR_[HSWordR]
14_98d4050e-9c98-49bf-9be1-85e12eb3ab13_001B_WordVL
:: Office 2013
15_ab4d047b-97cf-4126-a69f-34df08e2f254_0015_AccessRetail
15_259de5be-492b-44b3-9d78-9645f848f7b0_001C_AccessRuntimeRetail
15_4374022d-56b8-48c1-9bb7-d8f2fc726343_0015_AccessVolume
15_1b1d9bd5-12ea-4063-964c-16e7e87d6e08_0016_ExcelRetail
15_ac1ae7fd-b949-4e04-a330-849bc40638cf_0016_ExcelVolume
15_cfaf5356-49e3-48a8-ab3c-e729ab791250_00BA_GrooveRetail
15_4825ac28-ce41-45a7-9e6e-1fed74057601_00BA_GrooveVolume
15_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_00E7_HomeBusinessPipcRetail
15_cd256150-a898-441f-aac0-9f8f33390e45_0013_HomeBusinessRetail
15_1fdfb4e4-f9c9-41c4-b055-c80daf00697d_00CE_HomeStudentARMRetail
15_ebef9f05-5273-404a-9253-c5e252f50555_00DA_HomeStudentPlusARMRetail
15_98685d21-78bd-4c62-bc4f-653344a63035_002F_HomeStudentRetail
15_44984381-406e-4a35-b1c3-e54f499556e2_0044_InfoPathRetail
15_9e016989-4007-42a6-8051-64eb97110cf2_0044_InfoPathVolume
15_9103f3ce-1084-447a-827e-d6097f68c895_00EA_LyncAcademicRetail
15_ff693bf4-0276-4ddb-bb42-74ef1a0c9f4d_012D_LyncEntryRetail
15_fada6658-bfc6-4c4e-825a-59a89822cda8_012C_LyncRetail
15_e1264e10-afaf-4439-a98b-256df8bb156f_012C_LyncVolume
15_3169c8df-f659-4f95-9cc6-3115e6596e83_000F_MondoRetail
15_f33485a0-310b-4b72-9a0e-b1d605510dbd_000F_MondoVolume
15_3391e125-f6e4-4b1e-899c-a25e6092d40d_00A1_OneNoteFreeRetail
15_8b524bcc-67ea-4876-a509-45e46f6347e8_00A1_OneNoteRetail
15_b067e965-7521-455b-b9f7-c740204578a2_00A1_OneNoteVolume
15_12004b48-e6c8-4ffa-ad5a-ac8d4467765a_001A_OutlookRetail
15_8d577c50-ae5e-47fd-a240-24986f73d503_001A_OutlookVolume
15_5aab8561-1686-43f7-9ff5-2c861da58d17_00E6_PersonalPipcRetail
15_17e9df2d-ed91-4382-904b-4fed6a12caf0_0033_PersonalRetail
15_31743b82-bfbc-44b6-aa12-85d42e644d5b_0018_PowerPointRetail
15_e40dcb44-1d5c-4085-8e8f-943f33c4f004_0018_PowerPointVolume
15_4e26cac1-e15a-4467-9069-cb47b67fe191_00E8_ProfessionalPipcRetail
15_44bc70e2-fb83-4b09-9082-e5557e0c2ede_0014_ProfessionalRetail
15_f2435de4-5fc0-4e5b-ac97-34f515ec5ee7_003B_ProjectProRetail
15_ed34dc89-1c27-4ecd-8b2f-63d0f4cedc32_003B_ProjectProVolume
15_5517e6a2-739b-4822-946f-7f0f1c5934b1_003A_ProjectStdRetail
15_2b9e4a37-6230-4b42-bee2-e25ce86c8c7a_003A_ProjectStdVolume
15_064383fa-1538-491c-859b-0ecab169a0ab_0011_ProPlusRetail
15_2b88c4f2-ea8f-43cd-805e-4d41346e18a7_0011_ProPlusVolume
15_c3a0814a-70a4-471f-af37-2313a6331111_0019_PublisherRetail
15_38ea49f6-ad1d-43f1-9888-99a35d7c9409_0019_PublisherVolume
15_ba3e3833-6a7e-445a-89d0-7802a9a68588_0017_SPDRetail
15_32255c0a-16b4-4ce2-b388-8a4267e219eb_0012_StandardRetail
15_a24cca51-3d54-4c41-8a76-4031f5338cb2_0012_StandardVolume
15_15d12ad4-622d-4257-976c-5eb3282fb93d_0051_VisioProRetail
15_3e4294dd-a765-49bc-8dbd-cf8b62a4bd3d_0051_VisioProVolume
15_dae597ce-5823-4c77-9580-7268b93a4b23_0053_VisioStdRetail
15_44a1f6ff-0876-4edb-9169-dbb43101ee89_0053_VisioStdVolume
15_191509f2-6977-456f-ab30-cf0492b1e93a_001B_WordRetail
15_9cedef15-be37-4ff0-a08a-13a045540641_001B_WordVolume
:: Office 365 - 15.0 version
15_befee371-a2f5-4648-85db-a2c55fdf324c_00E9_O365BusinessRetail
15_537ea5b5-7d50-4876-bd38-a53a77caca32_00D6_O365HomePremRetail
15_149dbce7-a48e-44db-8364-a53386cd4580_00D4_O365ProPlusRetail
15_bacd4614-5bef-4a5e-bafc-de4c788037a2_00D5_O365SmallBusPremRetail
:: Office 365 - 16.0 version
16_6337137e-7c07-4197-8986-bece6a76fc33_00E9_O365BusinessRetail
16_2f5c71b4-5b7a-4005-bb68-f9fac26f2ea3_00D6_O365EduCloudRetail
16_537ea5b5-7d50-4876-bd38-a53a77caca32_00D6_O365HomePremRetail
16_149dbce7-a48e-44db-8364-a53386cd4580_00D4_O365ProPlusRetail
16_bacd4614-5bef-4a5e-bafc-de4c788037a2_00D5_O365SmallBusPremRetail
:: Office 2016
16_bfa358b0-98f1-4125-842e-585fa13032e6_0015_AccessRetail
16_9d9faf9e-d345-4b49-afce-68cb0a539c7c_001C_AccessRuntimeRetail
16_3b2fa33f-cd5a-43a5-bd95-f49f3f546b0b_0015_AccessVolume
16_424d52ff-7ad2-4bc7-8ac6-748d767b455d_0016_ExcelRetail
16_685062a7-6024-42e7-8c5f-6bb9e63e697f_0016_ExcelVolume
16_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_00E7_HomeBusinessPipcRetail
16_86834d00-7896-4a38-8fae-32f20b86fa2b_0013_HomeBusinessRetail
16_090896a0-ea98-48ac-b545-ba5da0eb0c9c_00CE_HomeStudentARMRetail
16_6bbe2077-01a4-4269-bf15-5bf4d8efc0b2_00DA_HomeStudentPlusARMRetail
16_c28acdb8-d8b3-4199-baa4-024d09e97c99_002F_HomeStudentRetail
16_e2127526-b60c-43e0-bed1-3c9dc3d5a468_002F_HomeStudentVNextRetail
16_b21367df-9545-4f02-9f24-240691da0e58_000F_MondoRetail
16_2cd0ea7e-749f-4288-a05e-567c573b2a6c_000F_MondoVolume
16_436366de-5579-4f24-96db-3893e4400030_00A3_OneNoteFreeRetail
16_83ac4dd9-1b93-40ed-aa55-ede25bb6af38_00A1_OneNoteRetail
16_23b672da-a456-4860-a8f3-e062a501d7e8_00A1_OneNoteVolume
16_5a670809-0983-4c2d-8aad-d3c2c5b7d5d1_001A_OutlookRetail
16_50059979-ac6f-4458-9e79-710bcb41721a_001A_OutlookVolume
16_5aab8561-1686-43f7-9ff5-2c861da58d17_00E6_PersonalPipcRetail
16_a9f645a1-0d6a-4978-926a-abcb363b72a6_0033_PersonalRetail
16_f32d1284-0792-49da-9ac6-deb2bc9c80b6_0018_PowerPointRetail
16_9b4060c9-a7f5-4a66-b732-faf248b7240f_0018_PowerPointVolume
16_4e26cac1-e15a-4467-9069-cb47b67fe191_00E8_ProfessionalPipcRetail
16_d64edc00-7453-4301-8428-197343fafb16_0014_ProfessionalRetail
16_0f42f316-00b1-48c5-ada4-2f52b5720ad0_003B_ProjectProRetail
16_82f502b5-b0b0-4349-bd2c-c560df85b248_003B_ProjectProVolume
16_16728639-a9ab-4994-b6d8-f81051e69833_003B_ProjectProXVolume
16_e9f0b3fc-962f-4944-ad06-05c10b6bcd5e_003A_ProjectStdRetail
16_82e6b314-2a62-4e51-9220-61358dd230e6_003A_ProjectStdVolume
16_431058f0-c059-44c5-b9e7-ed2dd46b6789_003A_ProjectStdXVolume
16_de52bd50-9564-4adc-8fcb-a345c17f84f9_0011_ProPlusRetail
16_c47456e3-265d-47b6-8ca0-c30abbd0ca36_0011_ProPlusVolume
16_6e0c1d99-c72e-4968-bcb7-ab79e03e201e_0019_PublisherRetail
16_fcc1757b-5d5f-486a-87cf-c4d6dedb6032_0019_PublisherVolume
16_971cd368-f2e1-49c1-aedd-330909ce18b6_012D_SkypeforBusinessEntryRetail
16_418d2b9f-b491-4d7f-84f1-49e27cc66597_012C_SkypeforBusinessRetail
16_03ca3b9a-0869-4749-8988-3cbc9d9f51bb_012C_SkypeforBusinessVolume
16_9103f3ce-1084-447a-827e-d6097f68c895_012C_SkypeServiceBypassRetail
16_4a31c291-3a12-4c64-b8ab-cd79212be45e_0012_StandardRetail
16_0ed94aac-2234-4309-ba29-74bdbb887083_0012_StandardVolume
16_2dfe2075-2d04-4e43-816a-eb60bbb77574_0051_VisioProRetail
16_295b2c03-4b1c-4221-b292-1411f468bd02_0051_VisioProVolume
16_0594dc12-8444-4912-936a-747ca742dbdb_0051_VisioProXVolume
16_c76dbcbc-d71b-4f45-b5b3-b7494cb4e23e_0053_VisioStdRetail
16_44151c2d-c398-471f-946f-7660542e3369_0053_VisioStdVolume
16_1d1c6879-39a3-47a5-9a6d-aceefa6a289d_0053_VisioStdXVolume
16_cacaa1bf-da53-4c3b-9700-11738ef1c2a5_001B_WordRetail
16_c3000759-551f-4f4a-bcac-a4b42cbf1de2_001B_WordVolume
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do (

if "%oVer%"=="%%A" (
reg query "%1\Registration\{%%B}" /v ProductCode %nul2% | find /i "-%%C-" %nul% && (
reg query "%1\Common\InstalledPackages" %nul2% | find /i "-%%C-" %nul% && (
if defined _oIds (set _oIds=!_oIds! %%D) else (set _oIds=%%D)
if /i 003D==%%C set SingleImage=1
)
)
)

)
)
exit /b

::========================================================================================================================================

::  1st column = Office version
::  2nd column = Volume or free retail product
::  3rd column = Retail product names that needs to be converted to the Volume product mentioned in 2nd column
::  Separator  = "_"

:tsksdata

set f=
for %%# in (
:: Office 2013
15_AccessVolume_-AccessRetail-
15_AccessRuntimeRetail
15_ExcelVolume_-ExcelRetail-
15_GrooveVolume_-GrooveRetail-
15_InfoPathVolume_-InfoPathRetail-
15_LyncAcademicRetail
15_LyncEntryRetail
15_LyncVolume_-LyncRetail-
15_MondoRetail
15_MondoVolume_-O365BusinessRetail-O365HomePremRetail-O365ProPlusRetail-O365SmallBusPremRetail-
15_OneNoteFreeRetail
15_OneNoteVolume_-OneNoteRetail-
15_OutlookVolume_-OutlookRetail-
15_PowerPointVolume_-PowerPointRetail-
15_ProjectProVolume_-ProjectProRetail-
15_ProjectStdVolume_-ProjectStdRetail-
15_ProPlusVolume_-ProPlusRetail-ProfessionalPipcRetail-ProfessionalRetail-
15_PublisherVolume_-PublisherRetail-
15_SPDRetail
15_StandardVolume_-StandardRetail-HomeBusinessPipcRetail-HomeBusinessRetail-HomeStudentARMRetail-HomeStudentPlusARMRetail-HomeStudentRetail-PersonalPipcRetail-PersonalRetail-
15_VisioProVolume_-VisioProRetail-
15_VisioStdVolume_-VisioStdRetail-
15_WordVolume_-WordRetail-
:: Office 2016
16_AccessRuntimeRetail
16_AccessVolume_-AccessRetail-
16_ExcelVolume_-ExcelRetail-
16_MondoRetail
16_MondoVolume_-O365AppsBasicRetail-O365BusinessRetail-O365EduCloudRetail-O365HomePremRetail-O365ProPlusRetail-O365SmallBusPremRetail-
16_OneNoteFreeRetail
16_OneNoteVolume_-OneNoteRetail-OneNote2021Retail-
16_OutlookVolume_-OutlookRetail-
16_PowerPointVolume_-PowerPointRetail-
16_ProjectProVolume_-ProjectProRetail-
16_ProjectProXVolume
16_ProjectStdVolume_-ProjectStdRetail-
16_ProjectStdXVolume
16_ProPlusVolume_-ProPlusRetail-ProfessionalPipcRetail-ProfessionalRetail-
16_PublisherVolume_-PublisherRetail-
16_SkypeServiceBypassRetail
16_SkypeforBusinessEntryRetail
16_SkypeforBusinessVolume_-SkypeforBusinessRetail-
16_StandardVolume_-StandardRetail-HomeBusinessPipcRetail-HomeBusinessRetail-HomeStudentARMRetail-HomeStudentPlusARMRetail-HomeStudentRetail-HomeStudentVNextRetail-PersonalPipcRetail-PersonalRetail-
16_VisioProVolume_-VisioProRetail-
16_VisioProXVolume
16_VisioStdVolume_-VisioStdRetail-
16_VisioStdXVolume
16_WordVolume_-WordRetail-
:: Office 2019
16_AccessRuntime2019Retail
16_Access2019Volume_-Access2019Retail-
16_Excel2019Volume_-Excel2019Retail-
16_Outlook2019Volume_-Outlook2019Retail-
16_PowerPoint2019Volume_-PowerPoint2019Retail-
16_ProjectPro2019Volume_-ProjectPro2019Retail-
16_ProjectStd2019Volume_-ProjectStd2019Retail-
16_ProPlus2019Volume_-ProPlus2019Retail-Professional2019Retail-
16_Publisher2019Volume_-Publisher2019Retail-
16_SkypeforBusiness2019Volume_-SkypeforBusiness2019Retail-
16_SkypeforBusinessEntry2019Retail
16_Standard2019Volume_-Standard2019Retail-HomeBusiness2019Retail-HomeStudentARM2019Retail-HomeStudentPlusARM2019Retail-HomeStudent2019Retail-Personal2019Retail-
16_VisioPro2019Volume_-VisioPro2019Retail-
16_VisioStd2019Volume_-VisioStd2019Retail-
16_Word2019Volume_-Word2019Retail-
:: Office 2021
:: OneNote2021Volume KMS license is not available
16_AccessRuntime2021Retail
16_Access2021Volume_-Access2021Retail-
16_Excel2021Volume_-Excel2021Retail-
16_Outlook2021Volume_-Outlook2021Retail-
16_OneNoteFree2021Retail
16_PowerPoint2021Volume_-PowerPoint2021Retail-
16_ProjectPro2021Volume_-ProjectPro2021Retail-
16_ProjectStd2021Volume_-ProjectStd2021Retail-
16_ProPlus2021Volume_-ProPlus2021Retail-Professional2021Retail-
16_Publisher2021Volume_-Publisher2021Retail-
16_SkypeforBusiness2021Volume_-SkypeforBusiness2021Retail-
16_Standard2021Volume_-Standard2021Retail-HomeBusiness2021Retail-HomeStudent2021Retail-Personal2021Retail-
16_VisioPro2021Volume_-VisioPro2021Retail-
16_VisioStd2021Volume_-VisioStd2021Retail-
16_Word2021Volume_-Word2021Retail-
:: Office 2024
16_Access2024Volume_-Access2024Retail-
16_Excel2024Volume_-Excel2024Retail-
16_Outlook2024Volume_-Outlook2024Retail-
16_PowerPoint2024Volume_-PowerPoint2024Retail-
16_ProjectPro2024Volume_-ProjectPro2024Retail-
16_ProjectStd2024Volume_-ProjectStd2024Retail-
16_ProPlus2024Volume_-ProPlus2024Retail-
16_SkypeforBusiness2024Volume
16_Standard2024Volume_-Home2024Retail-HomeBusiness2024Retail-
16_VisioPro2024Volume_-VisioPro2024Retail-
16_VisioStd2024Volume_-VisioStd2024Retail-
16_Word2024Volume_-Word2024Retail-
) do (
for /f "tokens=1-3 delims=_" %%A in ("%%#") do (

if %1==chkprod if "%oVer%"=="%%A" if not defined foundprod (
if /i "%%B"=="%2" set foundprod=1
)

if %1==getinfo if "%oVer%"=="%%A" (
echo: %%C | find /i "-%2-" %nul% && (
set _License=%%B
set _altoffid=%%B
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

::========================================================================================================================================
:: Leave empty line below
