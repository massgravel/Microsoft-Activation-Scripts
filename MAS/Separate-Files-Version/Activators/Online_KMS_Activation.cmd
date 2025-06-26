@set masver=3.4
@echo off



::============================================================================
::
::   Homepage: mass grave[.]dev
::      Email: mas.help@outlook.com
::
::============================================================================



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

::  Debug Mode:
::  To run the script in debug mode, change 0 to any parameter above that you want to run, in below line
set "_debug=0"

::  Script will run in unattended mode if parameters are used OR value is changed in above lines FOR activation or uninstall.



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
title  Online %KS% Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args set _args=%_args:re1=%
if defined _args set _args=%_args:re2=%
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

set "_serv=%_slser% Winmgmt"

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
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%evaluation_editions"
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
if %winbuild% LSS 7600 if exist "%SysPath%\licensing\skus\Security-Licensing-SLC-Component-SKU-%osedition%\*KMS*.xrm-ms" set sppks=1
if %winbuild% LSS 7600 if exist "%SysPath%\licensing\skus\Security-Licensing-SLC-Component-SKU-%osedition%\*VL-BYPASS*.xrm-ms" set sppks=1

if defined skunotfound (
call :dk_color %Red% "Required license files not found."
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
)

if defined sppks (
call :dk_color %Red% "%KS% activation is supported but failed to find the %KS% key."
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
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
set generickey=1
call :dk_inskey "[%key%]"
)

::========================================================================================================================================

:ks_office

if not %_actoff%==1 goto :ks_activate

call :oh_setspp

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

call :oh_getpath

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
echo You only have the Office Dashboard app installed. You need to install the full version of Office.
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

call :oh_expiredpreview 2013
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

call :oh_expiredpreview 2016 2019 2021 2024
if "%_actprojvis%"=="0" call :oh_fixprids
call :ks_process

::========================================================================================================================================

:ks_startmsi

if defined o14msi call :oh_setspp 14
if defined o14msi call :ks_processmsi 14 %o14msi_reg%
call :oh_setspp
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
::  https://learn.microsoft.com/en-us/windows/privacy/manage-connections-from-windows-operating-system-components-to-microsoft-services#19-software-protection-platform

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

::  https://learn.microsoft.com/en-us/azure/virtual-desktop/windows-10-multisession-faq

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
set _oMSI=
exit /b

::========================================================================================================================================

:oh_getpath

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

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o16msi=1&set o16msi_reg=%_86%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o16msi=1&set o16msi_reg=%_68%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o15msi=1&set o15msi_reg=%_86%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o15msi=1&set o15msi_reg=%_68%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o14msi=1&set o14msi_reg=%_86%\14.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\*Picker.dll" (set o14msi=1&set o14msi_reg=%_68%\14.0)

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

:oh_setspp

set isOspp=
if %winbuild% GEQ 9200 (
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
) else (
set isOspp=1
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
)
if "%1"=="14" (
set isOspp=1
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
)
exit /b

::========================================================================================================================================

:ks_process

for %%# in (%_oIds%) do (

set skipprocess=
if %_NoEditionChange%==1 if not defined _oMSI (
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
echo %%# | find /i "O365" %nul% && (
if "%oVer%"=="15" (call :dk_color %Gray% "Mondo 2013 is equivalent to O365 [15.0 version] in terms of the latest features.")
if "%oVer%"=="16" (call :dk_color %Gray% "Mondo 2016 is equivalent to O365 in terms of the latest features.")
)
set _prod=!_altoffid!
call :ks_osppready
)

if not "!key!"=="" (
echo "!allapps!" | find /i "!_actid!" %nul1% || call :oh_installlic
set generickey=1
call :dk_inskey "[!key!] [!_prod!]"
) else (
if not defined _oMSI (
set error=1
call :dk_color %Red% "Checking Product In Script              [Office %oVer%.0 !_prod! not found in script]"
call :dk_color %Blue% "Make sure you are using Latest MAS script."
) else (
call :dk_color %Red% "Checking Product In Script              [!_prod! MSI Retail is not supported]"
call :dk_color %Blue% "Use Ohook option to activate it. To activate with %KS%, you need to install Volume version of Office."
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
set _oMSI=1

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

call :msiofficedata %2

echo:
echo Processing Office...                    [MSI ^| %_version% ^| %_oArch%]

if not defined _oIds (
set error=1
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
exit /b
)

call :ks_process
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
set psc=powershell.exe -nop -c

set _slexe=sppsvc.exe& set _slser=sppsvc
if %winbuild% LEQ 6300 (set _slexe=SLsvc.exe& set _slser=SLsvc)
if %winbuild% LSS 7600 if exist "%SysPath%\SLsvc.exe" (set _slexe=SLsvc.exe& set _slser=SLsvc)
if %_slexe%==SLsvc.exe set _vis=1

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
if %winbuild% LSS 9200 set _wmic=1

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
net start %_slser% /y %nul%
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
set "srvlist=%srvlist% kms.id%-%ina.cn kms.moe%-%yuuko.com xinch%-%eng213618.cn kms.lol%-%i.best kms.mc%-%06.net"
set "srvlist=%srvlist% kms.0%-%t.net.cn win.k%-%ms.pub kms.wx%-%lost.com kms.moe%-%yuuko.top kms.gh%-%xi.com"

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
if %winbuild% LSS 7600 (set _vista=_vista)
if exist "%_temp%\.*" rmdir /s /q "%_temp%\" %nul%
md "%_temp%\" %nul%
call :ks_RenExport renewal%_vista% "%_temp%\Renewal.xml" Unicode
if not defined _int (call :ks_RenExport run_once%_vista% "%_temp%\Run_Once.xml" Unicode)
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

:renewal_vista:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>1999-01-01T12:00:00.34375</Date>
    <Author>WindowsAddict</Author>
    <Description>Online K-M-S Activation-Renewal - Weekly Task</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>1999-01-01T12:00:00.34375</StartBoundary>
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
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <IdleSettings>
      <Duration>PT10M</Duration>
      <WaitTimeout>PT1H</WaitTimeout>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <NetworkSettings />
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT5M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>%ProgramFiles%\Activation-Renewal\Activation_task.cmd</Command>
      <Arguments>Task</Arguments>
    </Exec>
  </Actions>
</Task>
:renewal_vista:

:run_once_vista:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>1999-01-01T12:00:00.34375</Date>
    <Author>WindowsAddict</Author>
    <Description>Online K-M-S Activation Run Once - Run and Delete itself on first Internet Contact</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <IdleSettings>
      <Duration>PT10M</Duration>
      <WaitTimeout>PT1H</WaitTimeout>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <NetworkSettings />
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT5M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>%ProgramFiles%\Activation-Renewal\Activation_task.cmd</Command>
      <Arguments>Task</Arguments>
    </Exec>
  </Actions>
</Task>
:run_once_vista:

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

::  Get Windows installed key channel

:k_channel

set _gvlk=
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path %spp% where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL and Description like '%%KMSCLIENT%%'" Get Name /value %nul6%') do (echo %%# findstr /i "Windows" %nul1% && set _gvlk=1)
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT Name FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL and Description like ''%%KMSCLIENT%%''').Get()).Name | %% {echo ('Name='+$_)}" %nul6%') do (echo %%# findstr /i "Windows" %nul1% && set _gvlk=1)
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
:: Windows Vista
cfd8ff08-c0d7-452b-9f60-ef5c70c32094_VKK3X-68KWM-X2YGT-QR4M6-4B%f%WMV___4_Enterprise
4f3d1606-3fea-4c01-be3c-8d671c401e3b_YFKBB-PQJJV-G996G-VWGXY-2V%f%3X8___6_Business
2c682dc2-8b68-4f63-a165-ae291d4cf138_HMBQG-8H2RH-C77VX-27R82-VM%f%QBT__16_BusinessN
d4f54950-26f2-4fb4-ba21-ffab16afcade_VTC42-BM838-43QHV-84HX6-XJ%f%XKV__27_EnterpriseN
:: Windows Server 2008
ad2542d4-9154-4c6d-8a44-30f11ee96989_TM24T-X9RMF-VWXK6-X8JC9-BF%f%GM2___7_ServerStandard
68b6e220-cf09-466b-92d3-45cd964b9509_7M67G-PC374-GR742-YH8V4-TC%f%BY3___8_ServerDatacenter
c1af4d90-d1bc-44ca-85d4-003ba33db3b9_YQGMW-MPWTJ-34KDK-48M3W-X4%f%Q6V__10_ServerEnterprise
01ef176b-3e0d-422a-b4f8-4ea880035e8f_4DWFP-JF3DJ-B7DTH-78FJB-PD%f%RHK__15_ServerEnterpriseIA64
ddfa9f7c-f09e-40b9-8c1a-be877a9a7f4b_WYR28-R7TFJ-3X2YQ-YCY4H-M2%f%49D__17_ServerWeb
7afb1156-2c1d-40fc-b260-aab7442b62fe_RCTX3-KWVHP-BR6TB-RB6DM-6X%f%7HP__18_ServerComputeCluster
2401e3d0-c50a-4b58-87b2-7e794b7d2607_W7VD6-7JFBR-RX26B-YKQ3Y-6F%f%FFJ__36_ServerStandardV
fd09ef77-5647-4eff-809c-af2b64659a45_22XQ2-VRXRG-P8D42-K34TD-G3%f%QQC__37_ServerDatacenterV
8198490a-add0-47b2-b3ba-316b12d647b4_39BXF-X8Q23-P2WWT-38T2F-G3%f%FPG__38_ServerEnterpriseV
::=======================================================================================================================================
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
9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce_HFTND-W9MK4-8B7MJ-B6C4G-XQ%f%BR2__16_MondoVolume_-O365AppsBasicRetail-O365BusinessRetail-O365EduCloudRetail-O365HomePremRetail-O365ProPlusRetail-O365SmallBusPremRetail-
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
) else if not defined _oMSI if %_NoEditionChange%==0 (
echo: %%E | find /i "-%2-" %nul% && (
set key=%%B
set _altoffid=%%D
set _actid=%%A
set _allactid=!_allactid! %%A
)
)
)

)
)
exit /b

::========================================================================================================================================

::  Below code is used to get alternate edition name and key if current edition doesn't support K-M-S activation

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

::========================================================================================================================================
:: Leave empty line below
