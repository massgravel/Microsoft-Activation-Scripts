@set masver=2.8
@echo off



::============================================================================
::
::   Homepage: mass grave[.]dev
::      Email: mas.help@outlook.com
::
::============================================================================



::  To activate, run the script with "/HWID" parameter or change 0 to 1 in below line
set _act=0

::  To disable changing edition if current edition doesn't support HWID activation, change the value to 1 from 0 or run the script with "/HWID-NoEditionChange" parameter
set _NoEditionChange=0

::  To run the script in debug mode, change 0 to "/HWID" in below line
set "_debug=0"

::  If value is changed in above lines or parameter is used then script will run in unattended mode



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
title  HWID Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args set _args=%_args:re1=%
if defined _args set _args=%_args:re2=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/HWID"                  set _act=1
if /i "%%A"=="/HWID-NoEditionChange"  set _NoEditionChange=1
if /i "%%A"=="-el"                    set _elev=1
)
)

for %%A in (%_act% %_NoEditionChange%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

call :dk_setvar

if %winbuild% LSS 10240 (
%eline%
echo Unsupported OS version detected [%winbuild%].
echo HWID Activation is only supported on Windows 10/11.
echo:
call :dk_color %Blue% "Use Online KMS activation option."
goto dk_done
)

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo HWID Activation is not supported on Windows Server.
call :dk_color %Blue% "Use KMS38 or Online KMS activation option."
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

REM :PowerShellTest: $ExecutionContext.SessionState.LanguageMode :PowerShellTest:

cmd /c "%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':PowerShellTest:\s*';iex ($f[1])"" | find /i "FullLanguage" %nul1% || (
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

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% && (
%eline%
echo [%winos% ^| %winbuild%]
echo:
echo Evaluation editions cannot be activated outside of their evaluation period. 
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
echo If you are, then try KMS38 activation option.
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

if defined _int (
reg delete "%_ident%" /f %nul%
reg query "%_ident%" %nul% && (
echo:
set error=1
call :dk_color %Red% "Deleting IdentityCRL Registry           [Failed] [%_ident%]"
)
for %%# in (wlidsvc LicenseManager sppsvc) do (%psc% "Start-Job { Restart-Service %%# } | Wait-Job -Timeout 20 | Out-Null")
call :dk_refresh
call :dk_act
call :dk_checkperm
)

::==========================================================================================================================================

::  Extended licensing servers tests incase error not found and activation failed

if %keyerror% EQU 0 if not defined _perm if defined _int (
set resfail=
ipconfig /flushdns %nul%
set "tls=[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;"

for %%# in (
login.live.com/ppsecure/deviceaddcredential.srf
purchase.mp.microsoft.com/v7.0/users/me/orders
) do if not defined resfail (
set "d1=Add-Type -AssemblyName System.Net.Http;"
set "d1=!d1! $client = [System.Net.Http.HttpClient]::new();"
set "d1=!d1! $response = $client.GetAsync('https://%%#').GetAwaiter().GetResult();"
set "d1=!d1! $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()"
%psc% "!tls! !d1!" %nul2% | findstr /i "PurchaseFD DeviceAddResponse" %nul1% || set resfail=1
)

if not defined resfail (
%psc% "!tls! irm https://licensing.mp.microsoft.com/v7.0/licenses/content -Method POST" | find /i "traceId" %nul1% || set resfail=1
)

if defined resfail (
set error=1
echo:
call :dk_color %Red% "Checking Licensing Servers              [Failed to Connect]"
set fixes=%fixes% %mas%licensing-servers-issue
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%licensing-servers-issue"
)
)

::==========================================================================================================================================

if %keyerror% EQU 0 if not defined _perm if defined _int (

reg query "%_ident%" %nul% || (
set error=1
echo:
call :dk_color %Red% "Generating New IdentityCRL Registry     [Failed] [%_ident%]"
)

reg query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /v DisableWindowsUpdateAccess %nul2% | find /i "0x1" %nul% && set wublock=1
reg query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /v DoNotConnectToWindowsUpdateInternetLocations %nul2% | find /i "0x1" %nul% && set wublock=1
if defined wublock call :dk_color %Red% "Checking Update Blocker In Registry     [Found]"

reg query "HKLM\SOFTWARE\Policies\Microsoft\WindowsStore" /v DisableStoreApps %nul2% | find /i "0x1" %nul% && (
set storeblock=1
call :dk_color %Red% "Checking Store Blocker In Registry      [Found]"
)

for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type ServiceSidType RequiredPrivileges FailureActions) do if not defined wucorrupt (
reg query HKLM\SYSTEM\CurrentControlSet\Services\wuauserv /v %%G %nul% || set wucorrupt=1
)

for %%G in (Parameters Security TriggerInfo) do if not defined wucorrupt (
reg query HKLM\SYSTEM\CurrentControlSet\Services\wuauserv\%%G %nul% || set wucorrupt=1
)

if defined wucorrupt (
call :dk_color %Red% "Checking Windows Update Registry        [Corruption Found]"
) else (
%psc% "Start-Job { Start-Service wuauserv } | Wait-Job -Timeout 20 | Out-Null"
sc query wuauserv | find /i "RUNNING" %nul% || (
set wuerror=1
sc start wuauserv %nul%
call :dk_color %Red% "Starting Windows Update Service         [Failed] [!errorlevel!]"
)
)

REM Check Internet related error codes

if not defined wucorrupt if not defined wublock if not defined wuerror if not defined storeblock (
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
call :dk_color %Blue% "Use KMS38 activation option instead."
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
set error=1
set showfix=1
call :dk_color %Gray% "Checking Windows Setup State            [%imagestate%]"
echo "%imagestate%" | find /i "RESEAL" %nul% && (
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


if not defined officeact if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% || (
set error=1
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

if not defined officeact (
if %osedition%==0 (
call :dk_color %Red% "Checking Edition Name                   [Not Found In Registry]"
) else (

if not exist "%SysPath%\spp\tokens\skus\%osedition%\%osedition%*.xrm-ms" if not exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%osedition%\*-%osedition%-*.xrm-ms" (
set error=1
set skunotfound=1
call :dk_color %Red% "Checking License Files                  [Not Found] [%osedition%]"
)

if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*-%osedition%-*.mum" (
set error=1
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


if not defined officeact (
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
if %winbuild% LSS 9200 (
echo Checking Eval WLMS Service              [Found]
) else (
call :dk_color %Red% "Checking Eval WLMS Service              [Found]"
)
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
set error=1
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Plugins\Objects\msft:rm/algorithm/hwid/4.0" /f ba02fed39662 /d %nul% || (
call :dk_color %Red% "Checking SPP Registry Key               [Incorrect ModuleId Found]"
set fixes=%fixes% %mas%issues_due_to_gaming_spoofers
call :dk_color2 %Blue% "Most likely caused by HWID spoofers. Help - " %_Yellow% " %mas%issues_due_to_gaming_spoofers"
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


call :dk_actid 55c92734-d682-4d71-983e-d6ec3f16059f
if not defined apps (
%psc% "Start-Job { Stop-Service sppsvc -force } | Wait-Job -Timeout 20 | Out-Null; $sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
call :dk_actid 55c92734-d682-4d71-983e-d6ec3f16059f
if not defined apps (
set "_notfoundids=Key Not Installed / Act ID Not Found"
call :dk_actids 55c92734-d682-4d71-983e-d6ec3f16059f
if not defined allapps (
set "_notfoundids=Not found"
)
set error=1
call :dk_color %Red% "Checking Activation IDs                 [!_notfoundids!]"
)
)


if exist "%tokenstore%\" if not exist "%tokenstore%\tokens.dat" (
set error=1
call :dk_color %Red% "Checking SPP tokens.dat                 [Not Found] [%tokenstore%\]"
)


if %winbuild% GEQ 9200 if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
for /f "delims=" %%a in ('%psc% "(Get-ScheduledTask -TaskName 'SvcRestartTask' -TaskPath '\Microsoft\Windows\SoftwareProtectionPlatform\').State" %nul6%') do (set taskinfo=%%a)
echo !taskinfo! | find /i "Ready" %nul% || (
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "actionlist" /f %nul%
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Microsoft\Windows\SoftwareProtectionPlatform\SvcRestartTask" %nul% || set taskinfo=Removed
if "!taskinfo!"=="" set "taskinfo=Not Found"
call :dk_color %Red% "Checking SvcRestartTask Status          [!taskinfo!, System might deactivate later]"
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

::========================================================================================================================================
:: Leave empty line below
