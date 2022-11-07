@setlocal DisableDelayedExpansion
@echo off



::============================================================================
::
::   This script is a part of 'Microsoft Activation Scripts' (MAS) project.
::
::   Homepage: massgrave.dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================



::  To activate, run the script with "/HWID" parameter or change 0 to 1 in below line
set _act=0

::  To disable changing edition if current edition doesn't support HWID activation, change the value to 1 from 0 or run the script with "/HWID-NoEditionChange" parameter
set _NoEditionChange=0

::  If value is changed in above lines or parameter is used then script will run in unattended mode



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
echo Error: This is not a correct file. It has LF line ending issue.
echo:
ping 127.0.0.1 -n 6 > nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title  HWID Activation

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
set    "Gray="100;97m""
set   "Green="42;97m""
set "Magenta="45;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set "Magenta="Darkmagenta" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
if %~z0 GEQ 200000 (set "_exitmsg=Go back") else (set "_exitmsg=Exit")

::========================================================================================================================================

if %winbuild% LSS 10240 (
%eline%
echo Unsupported OS version detected.
echo Project is supported for Windows 10/11.
goto dk_done
)

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo HWID Activation is not supported for Windows Server.
echo Use KMS38 or KMS Activation.
goto dk_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto dk_done
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%temp%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" 1>nul && (
if /i not "!_work!"=="!_ttemp!" (
%eline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto dk_done
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

>nul fltmc || (
if not defined _elev %nul% %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto dk_done
)

::========================================================================================================================================

cls
mode 102, 33
title  HWID Activation

echo:
echo Initializing...
call :dk_product
call :dk_ckeckwmic

::  Show info for potential script stuck scenario

sc start sppsvc %nul%
if %errorlevel% NEQ 1056 if %errorlevel% NEQ 0 (
echo:
echo Error code: %errorlevel%
call :dk_color %Red% "Failed to start [sppsvc] service, rest of the process may take a long time..."
echo:
)

::========================================================================================================================================

::  Check if system is permanently activated or not

call :dk_checkperm
if defined _perm (
cls
echo ___________________________________________________________________________________________
echo:
call :dk_color2 %_White% "     " %Green% "Checking: %winos% is Permanently Activated."
call :dk_color2 %_White% "     " %Gray% "Activation is not required."
echo ___________________________________________________________________________________________
if %_unattended%==1 goto dk_done
echo:
choice /C:10 /N /M ">    [1] Activate [0] %_exitmsg% : "
if errorlevel 2 exit /b
)
cls

::========================================================================================================================================

::  Check Evaluation version

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID 2>nul | find /i "Eval" 1>nul && (
%eline%
echo [%winos% ^| %winbuild%]
echo Evaluation Editions cannot be activated. Download ^& Install full version of Windows OS.
echo:
echo https://massgrave.dev/
goto dk_done
)
)

::========================================================================================================================================

::  Check SKU value / Check in multiple places to find Edition change corruption

set osSKU=
set regSKU=
set wmiSKU=

for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do set "regSKU=%%a"
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" 2^>nul') do if not errorlevel 1 set "wmiSKU=%%a"
if %_wmic% EQU 0 for /f "tokens=1" %%a in ('%psc% "([WMI]'Win32_OperatingSystem=@').OperatingSystemSKU" 2^>nul') do if not errorlevel 1 set "wmiSKU=%%a"

set osSKU=%wmiSKU%
if not defined osSKU set osSKU=%regSKU%

if not defined osSKU (
%eline%
echo SKU value was not detected properly. Aborting...
goto dk_done
)

::========================================================================================================================================

set error=

::  Check Internet connection

cls
echo:
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b
echo Checking OS Info                        [%winos% ^| %winbuild% ^| %arch%]

set _intcon=
for /f "delims=[] tokens=2" %%# in ('ping -n 1 licensing.mp.microsoft.com') do if not [%%#]==[] set _intcon=1

%psc% "$t = New-Object Net.Sockets.TcpClient;try{$t.Connect("""licensing.mp.microsoft.com""", 443)}catch{};$t.Connected" | findstr /i true 1>nul
if %errorlevel% EQU 0 (
echo Checking Internet Connection            [Connected]
) else (
set error=1
if defined _intcon (
call :dk_color %Red% "Checking Internet Connection            [Internet Found But Cant Connect licensing.mp.microsoft.com]"
call :dk_color %Magenta% "Make sure restricted Internet [Office/College] is not connected and URL is not blocked in the system"
) else (
call :dk_color %Red% "Checking Internet Connection            [Not Connected]"
)
)

::========================================================================================================================================

::  Check Windows Script Host

set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)

if %_WSH% EQU 0 (
reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
reg add "HKCU\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
if not "%arch%"=="x86" reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f /reg:32 %nul%
echo Enabling Windows Script Host            [Successful]
)

::========================================================================================================================================

echo Initiating Diagnostic Tests...

set "_serv=ClipSVC wlidsvc sppsvc LicenseManager Winmgmt wuauserv"

::  Client License Service (ClipSVC)
::  Microsoft Account Sign-in Assistant
::  Software Protection
::  Windows License Manager Service
::  Windows Management Instrumentation
::  Windows Update

call :dk_errorcheck

::========================================================================================================================================

::  Detect Key

set key=
set altkey=
set changekey=
set curedition=
set altedition=
set notworking=
set actidnotfound=

for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildBranch 2^>nul') do set "branch=%%b"

if defined applist call :hwiddata key attempt1
if not defined key call :hwiddata key attempt2

if defined notworking call :hwidfallback
if not defined key call :hwidfallback

if defined altkey (set key=%altkey%&set changekey=1&set notworking=)

if defined notworking if defined notfoundaltactID (
call :dk_color %Red% "Checking Alternate Edition For HWID     [%altedition% Activation ID Not Found]"
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
call :dk_color %Magenta% "Evaluation Windows Found. Install Full version of Windows. https://massgrave.dev/"
)
)

if not defined key (
%eline%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
echo Unable to find this product in the supported product list.
echo Make sure you are using updated version of the script.
echo https://massgrave.dev
echo:
goto dk_done
)

::========================================================================================================================================

::  Install key

echo:
if defined changekey (
call :dk_color %Magenta% "[%altedition%] Edition product key will be used to enable HWID activation."
echo:
)

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%key%')" %nul%
if not %errorlevel%==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%
set errorcode=%errorlevel%
cmd /c exit /b %errorcode%
if %errorcode% NEQ 0 set "errorcode=[0x%=ExitCode%]"

if %errorcode% EQU 0 (
call :dk_refresh
echo Installing Generic Product Key          [%key%] [Successful]
) else (
set error=1
call :dk_color %Red% "Installing Generic Product Key          [%key%] [Failed] %errorcode%"
if defined applist if defined actidnotfound call :dk_color %Red% "Activation ID not found for this key. Make sure you are using updated version of MAS."
)

::========================================================================================================================================

::  Change Windows region to USA to avoid activation issues as Windows store license is not available in many countries 

for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Control Panel\International\Geo" /v Name 2^>nul') do set "name=%%b"
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Control Panel\International\Geo" /v Nation 2^>nul') do set "nation=%%b"

set regionchange=
if not "%name%"=="US" (
set regionchange=1
%psc% Set-WinHomeLocation -GeoId 244
if !errorlevel! EQU 0 (
echo Changing Windows Region To USA          [Successful]
) else (
call :dk_color %Red% "Changing Windows Region To USA          [Failed]"
)
)

::==========================================================================================================================================

::  Generate GenuineTicket.xml and apply
::  Most correct way to apply a ticket is by restarting ClipSVC service but we can not check the log details in this way
::  To get the log details and also to correctly apply ticket, script will install tickets two times (service restart + clipup -v -o)

set "tdir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if not exist "%tdir%\" md "%tdir%\" %nul%

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%

call :hwiddata ticket

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "Generating GenuineTicket.xml            [Failed]"
echo [%encoded%]
if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
goto :dl_final
) else (
echo Generating GenuineTicket.xml            [Successful]
)

set "_xmlexist=if exist "%tdir%\GenuineTicket.xml""

%_xmlexist% (
net stop ClipSVC /y %nul%
net start ClipSVC /y %nul%
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

%_xmlexist% (
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
call :dk_color %Red% "Installing GenuineTicket.xml            [Failed With ClipSVC Service Restart, Wait...]"
)
)

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%
clipup -v -o
if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%

::==========================================================================================================================================

call :dk_product

echo:
echo Activating...
echo:

call :dk_act
call :dk_checkperm
if defined _perm (
call :dk_color %Green% "%winos% is permanently activated."
goto :dl_final
)


if not defined error (

REM Clear store ID related registry to fix activation incase if there is any corruption

set "_ident=HKU\S-1-5-19\SOFTWARE\Microsoft\IdentityCRL"
reg delete "!_ident!" /f %nul%
reg query "!_ident!" %nul% && (
call :dk_color %Red% "Deleting a Registry                     [Failed] [!_ident!]"
) || (
echo Deleting a Registry                     [Successful] [!_ident!]
)

REM Refresh some services and license status

for %%# in (wlidsvc LicenseManager sppsvc) do (net stop %%# /y %nul% & net start %%# /y %nul%)
call :dk_refresh
call :dk_act
call :dk_checkperm
)

if defined _perm (
call :dk_color %Green% "%winos% is permanently activated."
) else (
call :dk_color %Red% "Activation Failed %error_code%"
if defined notworking (
call :dk_color %Magenta% "At the time of writing this, HWID Activation was not supported for this product."
) else (
call :dk_color2 %Magenta% "Check this page for help" %_Yellow% " https://massgrave.dev/troubleshoot"
)
)

::========================================================================================================================================

:dl_final

echo:

if defined regionchange (
%psc% Set-WinHomeLocation -GeoId %nation%
if !errorlevel! EQU 0 (
echo Restoring Windows Region                [Successful]
) else (
call :dk_color %Red% "Restoring Windows Region                [Failed] [%name%-%nation%]"
)
)

if %osSKU%==175 call :dk_color %Red% "ServerRdsh Editon does not officially support activation on non-azure platforms."

goto :dk_done

::========================================================================================================================================

::  Get Windows permanent activation status

:dk_checkperm

if %_wmic% EQU 1 wmic path SoftwareLicensingProduct where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name /value 2>nul | findstr /i "Windows" 1>nul && set _perm=1||set _perm=
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Name FROM SoftwareLicensingProduct WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL').Get()).Name | %% {echo ('Name='+$_)}" 2>nul | findstr /i "Windows" 1>nul && set _perm=1||set _perm=
exit /b

::  Refresh license status

:dk_refresh

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call RefreshLicenseStatus %nul%
if %_wmic% EQU 0 %psc% "$null=(([WMICLASS]'SoftwareLicensingService').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::  Activation command

:dk_act

set error_code=
if %_wmic% EQU 1 wmic path SoftwareLicensingProduct where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null" call Activate %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL').Get()).Activate()" %nul%
if not %errorlevel%==0 cscript //nologo %windir%\system32\slmgr.vbs /ato %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 (set "error_code=[Error Code: 0x%=ExitCode%]") else (set error_code=)
exit /b

::  Get Windows Activation IDs

:dk_actids

set applist=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f') get ID /VALUE" 2^>nul')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f''').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))
exit /b

::  Get Product name (WMI/REG methods are not reliable in all conditions, hence winbrand.dll method is used)

:dk_product

set winos=
set d1=[DllImport(\"winbrand\",CharSet=CharSet.Unicode)]public static extern string BrandingFormatString(string s);
set d2=$AP=Add-Type -Member '%d1%' -Name D1 -PassThru; $AP::BrandingFormatString('%%WINDOWS_LONG%%')
for /f "delims=" %%s in ('"%psc% %d2%"') do if not errorlevel 1 (set winos=%%s)
echo "%winos%" | find /i "Windows" 1>nul || (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "winos=%%b"
if %winbuild% GEQ 22000 (
set winos=!winos:Windows 10=Windows 11!
)
)
exit /b

::  Check wmic.exe

:dk_ckeckwmic

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul && set _wmic=1
)
exit /b

::========================================================================================================================================

:dk_errorcheck

::  Check disabled services

set serv_ste=
for %%# in (%_serv%) do (
set serv_dis=
reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v Start %nul% || set serv_dis=1
for /f "skip=2 tokens=2*" %%a in ('reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v Start 2^>nul') do if /i %%b equ 0x4 set serv_dis=1
if defined serv_dis (if defined serv_ste (set "serv_ste=!serv_ste! %%#") else (set "serv_ste=%%#"))
)

::  Change disabled services startup type to default

set serv_csts=
set serv_cste=

if defined serv_ste (
for %%# in (%serv_ste%) do (
if /i %%#==ClipSVC        (reg add "HKLM\SYSTEM\CurrentControlSet\Services\%%#" /v "Start" /t REG_DWORD /d "3" /f %nul% & sc config %%# start= demand %nul%)
if /i %%#==wlidsvc        sc config %%# start= demand %nul%
if /i %%#==sppsvc         (reg add "HKLM\SYSTEM\CurrentControlSet\Services\%%#" /v "Start" /t REG_DWORD /d "2" /f %nul% & sc config %%# start= delayed-auto %nul%)
if /i %%#==LicenseManager sc config %%# start= demand %nul%
if /i %%#==Winmgmt        sc config %%# start= auto %nul%
if /i %%#==wuauserv       sc config %%# start= demand %nul%
if !errorlevel!==0 (
if defined serv_csts (set "serv_csts=!serv_csts! %%#") else (set "serv_csts=%%#")
) else (
set error=1
if defined serv_cste (set "serv_cste=!serv_cste! %%#") else (set "serv_cste=%%#")
)
)
)

if defined serv_csts echo Enabling Disabled Services              [Successful] [%serv_csts%]

if defined serv_cste (
echo %serv_cste% | findstr /i "ClipSVC sppsvc" %nul% && (
call :dk_color %Red% "Enabling Disabled Services              [Failed] [%serv_cste%] [Restart System]"
) || (
call :dk_color %Red% "Enabling Disabled Services              [Failed] [%serv_cste%]"
)
)

::========================================================================================================================================

::  Check if the services are able to run or not
::  Workarounds are added to get correct status and error code because sc query doesn't output correct results in some conditions

set serv_e=
for %%# in (%_serv%) do (
set errorcode=
set checkerror=
net start %%# /y %nul%
sc query %%# | find /i "4  RUNNING" %nul% || set checkerror=1

sc start %%# %nul%
set errorcode=!errorlevel!
if !errorcode! NEQ 1056 if !errorcode! NEQ 0 set checkerror=1
if defined checkerror if defined serv_e (set "serv_e=!serv_e!, %%#-!errorcode!") else (set "serv_e=%%#-!errorcode!")
)

if defined serv_e (
set error=1
call :dk_color %Red% "Starting Services                       [Failed] [%serv_e%]"
)

::========================================================================================================================================

::  Various error checks

for %%# in (wmic.exe) do @if "%%~$PATH:#"=="" (
call :dk_color %Gray% "Checking WMIC.exe                       [Not Found]"
)


%psc% $ExecutionContext.SessionState.LanguageMode 2>nul | find /i "Full" 1>nul || (
set error=1
call :dk_color %Red% "Checking Powershell                     [Not Responding]"
)


if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" 2>nul | find /i "computersystem" 1>nul
if %errorlevel% NEQ 0 (
set error=1
call :dk_color %Red% "Checking WMI                            [Not Responding] %_wmic%"
)


if not "%regSKU%"=="%wmiSKU%" (
set error=1
call :dk_color %Red% "Checking WMI/REG SKU                    [Difference Found - WMI:%wmiSKU% Reg:%regSKU%]"
)


DISM /English /Online /Get-CurrentEdition %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 set "error_code=[0x%=ExitCode%]"
if %error_code% NEQ 0 (
call :dk_color %Red% "Checking DISM                           [Not Responding] %error_code%"
)


if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
call :dk_color %Red% "Checking Eval Packages                  [Non-Eval Licenses are installed in Eval Windows]"
)


cscript //nologo %windir%\system32\slmgr.vbs /dlv %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 set "error_code=0x%=ExitCode%"
if %error_code% NEQ 0 (
set error=1
call :dk_color %Red% "Checking slmgr /dlv                     [Not Responding] %error_code%"
)


reg query "HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\PersistedTSReArmed" %nul% && (
set error=1
call :dk_color %Red% "Checking Rearm                          [System Restart Is Required]"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState" %nul% && (
set error=1
call :dk_color %Red% "Checking ClipSVC                        [System Restart Is Required]"
)


for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" 2^>nul') do if /i %%b NEQ 0x0 (
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" /t REG_DWORD /d "0" /f %nul%
call :dk_color %Red% "Checking SkipRearm                      [Default 0 Value Not Found, Changing To 0]"
net stop sppsvc /y %nul%
net start sppsvc /y %nul%
set error=1
)


call :dk_actids
if not defined applist (
net stop sppsvc /y %nul%
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if !errorlevel! NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
call :dk_refresh
call :dk_actids
if not defined applist (
set error=1
call :dk_color %Red% "Checking Activation IDs                 [Not Found]"
)
)


set token=0
if exist %Systemdrive%\Windows\System32\spp\store\2.0\tokens.dat set token=1
if exist %Systemdrive%\Windows\System32\spp\store_test\2.0\tokens.dat set token=1
if %token%==0 (
set error=1
call :dk_color %Red% "Checking SPP tokens.dat                 [Not Found]"
)

if not exist %SystemRoot%\system32\sppsvc.exe (
set error=1
call :dk_color %Red% "Checking sppsvc.exe File                [Not Found]"
)

if /i %error_code% EQU 0xc0000022 (
echo "%serv_e%" | find /i "sppsvc" %nul% && (
call :dk_color %Magenta% "Looks like you may have used a Gaming spoofer. Check Activation Troubleshoot option in MAS."
)
)
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

:dk_done

echo:
if %_unattended%==1 timeout /t 2 & exit /b
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
pause >nul
exit /b

::========================================================================================================================================

::  1st column = Activation ID
::  2nd column = Generic Retail/OEM/MAK Key
::  3rd column = SKU ID
::  4th column = Key part number
::  5th column = Ticket signature value. It's as it is, it's not encoded. (Check https://massgrave.dev/hwid.html#Manual_Activation to see how it's generated)
::  6th column = 1 = activation is not working (at the time of writing this), 0 = activation is working
::  7th column = Key Type
::  8th column = WMI Edition ID
::  9th column = Version name incase same Edition ID is used in different OS versions with different key
::  Separator  = _


:hwiddata

for %%# in (
8b351c9c-f398-4515-9900-09df49427262_XGVPP-NMH47-7TTHJ-W3FW7-8HV2C___4_X19-99683_X9J5T0gPQprYpz2euPvoJGlkurIO9h6N8ypE0KWYVpy0nbCKYnqSUCD7u8ReXAmc085jX2uM5PKurSee9Yq/PxesgiysQHDBsOhr98MXZZiIgy4ssnz2gZF70KB8tO3X7kk9LHwxXfz3rlquYPod9swe90nqvVaJMWCpQK0InUw_0_OEM:NONSLP_Enterprise
c83cef07-6b72-4bbc-a28f-a00386872839_3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT__27_X19-98746_WFZBjlVtHQumoaVE28/NHsRvv1lgkkfav6NPHqr6OC2u4vxkjjJkkl9OTF6DpHJu0IFrrQv+HYcdZ/WC5EzhOMqMxcujTBSAN7xLIVEbs72Db0Bi5iDAbOltJpk8QKKe18otQJ6vajW5WOPXjbgSJfDFaZQfiwvIJ1ICXt+stog_0_Volume:MAK_EnterpriseN
4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7JG-NPHTM-C97JM-9MPGT-3V66T__48_X19-98841_K3qev/5gQpX1RK1F9M9beEWWv/di1GsRF7OUcEMGTGDTYnaRenRcJaO8zOHQQvKDc57fon/v77ZpHQHT/jWWhWnLm7Ssory+s8tOs72fPjivVBDwpSPIEC1v+8Vpb4a3XCZet2e/Z5wmpCq9XDkowys3IcxYM0mHWBaNPu8gIe4_0_____Retail_Professional
9fbaf5d6-4d83-4422-870d-fdda6e5858aa_2B87N-8KFHP-DKV6R-Y2C8J-PKCKT__49_X19-98859_WcAcor6kQgxgkTRzcoxnb8UIoo5/ueYeaOKqy9/xAzlruHAKxhatXeGtSI58lXcCK5hxXkDmcyrRFwWSwdvg0txwTi7VusYcTNCLdmNWU/62iDrBhzMrCYtuhW9EV/g4+TlbjSm4PBJ0HMlI4YzAEnyJiBgKPDgBQ8Gj9LRbEgU_0_____Retail_ProfessionalN
f742e4ff-909d-4fe9-aacb-3231d24a0c58_4CPRK-NM3K3-X6XXQ-RXX86-WXCHW__98_X19-98877_MBDSEqlayxtVVEgIeAl8milgjS/BVHow6+MmpCyh9nweuctlT1+LbEHmDlnqDeLr9FQrN2FpEJtNr26rE0niMdvcAP51MfJsREyhWOEbrWwWyMH0KwDAci2WxWZTJp/SEZnq5HYYT1pPPLMWAkKRHJksJJFtg4zBtoyHvLjc35c_0_____Retail_CoreN
1d1bac85-7365-4fea-949a-96978ec91ae0_N2434-X9D7W-8PF6X-8DV9T-8TYMD__99_X19-99652_mpjCoh6soA/rwJutsjekZpA9vDUD8znR20V/c8FwSjuCcSbPhmP6bpJR9rfptAZqpagliMxA/OUZsx0Knt0n/hgOy2mv8pr24gI9uYXK8EfhG74bVdsyvZz1tyA6CaVR02ZahQvbKYzCmXUvsI+Wge3bHbKbVpn9Mvl+itn2a4g_0_____Retail_CoreCountrySpecific
3ae2cc14-ab2d-41f4-972f-5e20142771dc_BT79Q-G7N6G-PGBYW-4YWX6-6F4BT_100_X19-99661_KaUs6KwvtthPOsxd3x0tU/baKSv1DWSFOqbq7PbU/uYEY95p0Skzv3y4aXq+xVmfwSt8STL/4vSfFIAlsaRh7Vnq6Y/Ael8joeqI8hBN461fykoHxSELRMJ+eed50T0cJUS79ol6OTBOCCVeHgmtGVbHuL88TMWW69fGNdIMM3U_0_____Retail_CoreSingleLanguage
2b1f36bb-c1cd-4306-bf5c-a0367c2d97d8_YTMG3-N6DKC-DKB77-7M9GH-8HVX7_101_X19-98868_NpHxrAtA+GL6kawAP5Z2UdfUVcKFvf9UzEe6FIV/HztZqxpMBDFv2hdxCjD9+T8PKcW8j3n04McelOAgr3lD37Fu+wrvJIGX0dG3xEtU/MG9L9X5baBS8H6AmC6rq2+w5NUY8EchK9W2oatBflFb8IcfCSeAyOfsJei6bdu4mp8_0_____Retail_Core
2a6137f3-75c0-4f26-8e3e-d83d802865a4_XKCNC-J26Q9-KFHD2-FKTHY-KD72Y_119_X19-99606_gtywgqIP3j+bliKdunuseeZWtsOzWhj+DmSBq7nqeNarHutgbWEwvcRiGo+nwxONt9Ak/VyuO76ZWH/db3iRVTk1y61vFv15gVlOy1ovLjVHBvmPVdQXIne2N+pIMb0eBhZWHRX63mYdkZRZ0wg/+bj4xsjJv+qLpWhVCzNMge4_0_OEM:NONSLP_PPIPro
e558417a-5123-4f6f-91e7-385c1c7ca9d4_YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY_121_X19-98886_VuBmoSUdF63Cvwm9wNlc2yhD2tP9B72iVVWFNcbAwDGXF6o06oNMsIJ0VqGJDdBzZjVGw2wHokMabxZNDyIl90CO7trwgV8S0lLJVLymxyUaE3ThvN3YUsi9Q3H+5Kr0RpsojCWb+UQd/GY4bSXfyStXFylj6im7yv0db/ZWGbw_0_____Retail_Education
c5198a66-e435-4432-89cf-ec777c9d0352_84NGF-MHBT6-FXBX8-QWJK7-DRR8H_122_X19-98892_jQ6S2bbNoVrp/zvi8BEUwCf7fge1nAdspcjXyTeTySUiR+hXPiKQEWgyLqAdZ5Or+X2JGT/LZN1/eZ9P+REmzG/WQotZ+fyyPguoSsES+d312RkfmQoI5gVanEkGjZSU4YohREM/Vyf9MOO7dbH9MMEpFm2mje6OnhyJo2gux0g_0_____Retail_EducationN
cce9d2de-98ee-4ce2-8113-222620c64a27_KCNVH-YKWX8-GJJB9-H9FDT-6F7W2_125_X22-66075_wJ/BPDFz+13PVJtBqBo+E4LCm3LoMVALCQUun9kXGBULr7V8FQ5nKUudUGHDLNNVIIicdw9Uh26BKAt0/hnE7BpBkzwdi4qAdZgKXQ1t06Ek4+zXmoT225NvpaHsuhDkE687TtCB1ZWvAulA8G9ehE3HTJSoNm4wCFOQyIQQtqQ_1_Volume:MAK_EnterpriseS_VB
d06934ee-5448-4fd1-964a-cd077618aa06_43TBQ-NH92J-XKTM7-KT3KK-P39PB_125_X21-83233_V+y0SFmAnGwRwgNz+0sO0mj+XxSjbdRDpom1Iqx2BJcsf96Q5ittJOcMhKiNswyKuq5suM5vy60tA/AUdb1mrnnrnXfmz7nFam/BIOOfa18GA7vd1aNFufhpmCiMWxoGSewH/T1pnCZrsvGYIj//qC7aiQVKYBngO7UYWGaytgc_0_OEM:NONSLP_EnterpriseS_RS5
706e0cfd-23f4-43bb-a9af-1a492b9f1302_NK96Y-D9CD8-W44CQ-R8YTK-DYJWX_125_X21-05035_U2DIv+LAhSGz0rNbTiMQYaP3M41+0+ZioF7vh0COeeJSIruDFCZ3Li7ZM3dSleg6QTCxG04uZ3i3r1bCZv0+WAfU9rG+3BqLAwKlJS/31rETeRWvrxB1UK4mTMHwAJc9txDAc15ureqF+2b9pIIpwLljmFer6fI7z0iI6I/ZuTU_0_OEM:NONSLP_EnterpriseS_RS1
faa57748-75c8-40a2-b851-71ce92aa8b45_FWN7H-PF93Q-4GGP8-M8RF3-MDWWW_125_X19-99617_0frpwr4N/wBVRA/nOvAMqkxmRj6Vv9mA+jVNtnurAL1TjkPN/y+6YVUd5MP/Y4As4kddHoHiZXI+2siKHJsaV95ppXoHKR8d7FRVitr1F+82TbB7OVvdCclGrRZymnq25HvtSC3BROHt7ZXTgSCWMyB7MlbLiqHiTymOj5OMX1g_0_OEM:NONSLP_EnterpriseS_TH
2c060131-0e43-4e01-adc1-cf5ad1100da8_RQFNW-9TPM3-JQ73T-QV4VQ-DV9PT_126_X22-66108_UeA6O2iIW6zFMJzLMCQjVA7gUHOGRTiFB6LPrgjhgfJEXSZnDjxw8wsR+tp+JQWeaQDsVt06c2byH3z7Ft2wNk8n3gcXUknIjlcCckNjw05WDI64/wCqz+gtf1RajMEoV/mODpBx7rdLtCg03FyV7Z9LOib4/WLSmnxjDPKMG7s_1_Volume:MAK_EnterpriseSN_VB
e8f74caa-03fb-4839-8bcc-2e442b317e53_M33WV-NHY3C-R7FPM-BQGPT-239PG_126_X21-83264_NtP6sMWmOTCdABAbgIZfxZzRs8zaqzfaabLeFXQJvfJvQPLQPk2UxMliASJG+7YwwbTD8pyhUoQqUYrlCzJZ6jDSDyUTJkXgo9akR4fBOg6Z5wn5fW8NGAMDcLND5d9XxHl0gWH/HZNIs/GZaPJsCVVqPr7X8bk/y0DeIofxICU_1_Volume:MAK_EnterpriseSN_RS5
3d1022d8-969f-4222-b54b-327f5a5af4c9_2DBW3-N2PJG-MVHW3-G7TDK-9HKR4_126_X21-04921_WeNSkuiC3iyNT9tDqlj6KvM17UYMsYjEelyyMEyPEXSAbYA08lYtYJjCzxSE9T30p9dxqPIuj370OwHhAxG8a51/HoLNWR0grj08HmdOXUA8Ap4clEivxKM0zRvwPR6L2M2HQP0nN54c9It7ikzweJ0X2HHOb58oEw9LbMeUM/Y_0_Volume:MAK_EnterpriseSN_RS1
60c243e1-f90b-4a1b-ba89-387294948fb6_NTX6B-BRYC2-K6786-F6MVQ-M7V2X_126_X19-98770_QLG40WW/TtUqtir9K6FJCQXU1mfn27uutdOunHJ3gXk6v0Mbxaqu9GKqpg5xFzdFiOPb/8Bmk/ylwceXgoaUx1nKcBGb/Bg+jICiNMEYIbGyMuYiHb0iJeVbjbBLLfWuAAuUPftfnKPH3dAu1YvhaS5nv7a5wICrXdJWeVNpBxk_0_Volume:MAK_EnterpriseSN_TH
eb6d346f-1c60-4643-b960-40ec31596c45_DXG7C-N36C4-C4HTG-X4T3X-2YV77_161_X21-43626_vHO/5UEtrsDzGC30A2Ya5DYXlNMs7hVYiLvM7X31xkaFMxogbiy3ZDxBbjRku3VXyW+TYsFX/D/wdJgFmMrhsNrObkxqzYMMRjx+BpwOx2PspKpS2RyzovyRl8v93SvHB5IyoO2/3pm2YqJDK1hXLhms6+DDPuiofQt36q47reQ_0_____Retail_ProfessionalWorkstation
89e87510-ba92-45f6-8329-3afa905e3e83_WYPNQ-8C467-V2W6J-TX4WX-WT2RQ_162_X21-43644_phlxNLr+sk8cCCmAVU3k3XrtD6sFDeoaODc+21soKqePbVQbzPHgokS73ccok6/gDfu/u5UKc7omL8pm2IhIhf70oC+8M/FFp0zRFeC/ZFXdF2tL23oKWI9kZbvcaoZBiqaDGc1bNYi5KAZYaJU8wwqw16ZnohQJZ7QR9cgUfFQ_0_____Retail_ProfessionalWorkstationN
62f0c100-9c53-4e02-b886-a3528ddfe7f6_8PTT6-RNW4C-6V7J2-C2D3X-MHBPB_164_X21-04955_Px7QWdfy0esrMzQoydKlmIcGdfV0pQvbnumyrh4evDNF9gpENm8OIfZfljIynury0qZAkw4AG3uGyp+5IxZGIh6U3dz41uNVfEcA9NZ34OEBXMtjEOU1ZbJ8wp8JecQKwlORclvsri9OOi0GbGc0TYRanlci2jJL/3x/gSuWXCs_0_____Retail_ProfessionalEducation
13a38698-4a49-4b9e-8e83-98fe51110953_GJTYN-HDMQY-FRR76-HVGC7-QPF8P_165_X21-04956_GRSYno4+yqU/JMxHLDKdvdFWRz1uT90n5JkTvSqztDvXMf/mBhSV/OpppJWGo6UL0FwqYcu9oXl+Vx336pLAE5/EDzQHh+QCwOCDJiTKnd3hW/zrGMe6Sb0OAIkNNML9gcOBbr1IHFWhN99r8ZWl5JjpzMs2nPjejB1Ec8NCcpE_0_____Retail_ProfessionalEducationN
df96023b-dcd9-4be2-afa0-c6c871159ebe_NJCF7-PW8QT-3324D-688JX-2YV66_175_X21-41295_kkJyX1AwYgDYcGK1eIRdocybkbAfEtQkDxhRUhY89X2i2PSD9jcsGQgHWyD3KUKWb3bzR8QkDS3MTeieOw3EzD0RyAQhHc6lRR+rk18lh5UOVCgrZ6byxn29Ur+jAh0LJXImggC9JMGb2cTYaZckxzm3ICoAKwrmI9JnrzBTVmY_0_____Retail_ServerRdsh
d4ef7282-3d2c-4cf0-9976-8854e64a8d1e_V3WVW-N2PV2-CGWC3-34QGF-VMJ2C_178_X21-32983_YIMgXu2dZ9x1r1NLs3egTc/8EYc1RndYDvoX7QquQQLnhnhbSNBw3hmlqrQ0zNsTLut3EKpGZK2CwPspJJWE60lecdxI4211K748P6vkuqHPL4uFqXyKxTG3qRrtDIra5nnMn4GqG2fWuguzTXaumu8cJU3H1uTOsR1E/DQnJJ0_0_____Retail_Cloud
af5c9381-9240-417d-8d35-eb40cd03e484_NH9J3-68WK7-6FB93-4K3DF-DJ4F6_179_X21-32987_H0qrFdf+FQxcSRJDtEwd8OfwC4iH/25Q01jz3QuB9yhEqB0W1i83u0WDpVK04pvU1EDCCRRI/DhXynbkWpLC0chdTOW4k5jIy+aa0cD3fccz9ChSjVHMzyTg3abEVFAvy9rttUyxcFIOKcINXHTxTRp5cZPwOa393tlJyBiliAo_0_____Retail_CloudN
8ab9bdd1-1f67-4997-82d9-8878520837d9_XQQYW-NFFMW-XJPBH-K8732-CKFFD_188_X21-99378_Bwx3E7qmE6M8UR6+KPqLnnavI6ThNHHUO717RJY9di2YI9rzC3O0LceXOHjshSKwfwxosqFsD/p/inrJmabed1yA/ZWwISyGtAIGTtRgpuSE4TAfW6KEW0v7rcr2wwwDq7DHSuz4QN4odEGe9bvtx4zIZKufQzzN4TN2rd/BJkE_0_____OEM:DM_IoTEnterprise
ed655016-a9e8-4434-95d9-4345352c2552_QPM6N-7J2WJ-P88HH-P3YRH-YY74H_191_X21-99682_lE8qL1p4m68mv9wcxU2sdKZPIccybtOjr+aMAdV+sLHs9wzE26oz5GiSZ3UzpU7yoYrNMqwGkKX6mrCEGRLh+XR2Ricp7ELA1PkzaGm0FLUqaK2GNVQ00i+s6KcA2XRr/gWOhhGTqSCjpSi9cMiqMbftf9Bo/BJVK3ib9xU4OQw_0_OEM:NONSLP_IoTEnterpriseS_VB
d4bdc678-0a4b-4a32-a5b3-aaa24c3b0f24_K9VKN-3BGWV-Y624W-MCRMQ-BHDCD_202_X22-53884_hPcIn0dF9Dq6zlXd3RxBqVDPDnf5sTasTjUqhD6lGc9IkTc8476NHd1PV1Ds++VO34/dw2H2PWk33LT5Es6PnUi32Ypva4POy4QJo5W3qyduiJiHUOM5GS9yAkKfdHFgUXaUVwopYKq+EwmgxFmEvHYdWgREHgIMyNoKAZQK0Ok_0_____Retail_CloudEditionN
92fb8726-92a8-4ffc-94ce-f82e07444653_KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W_203_X22-53847_DCP6QzPj+BD1EEmlBelBt7x9AmvQOfd7kdkUB0b0x6/TNHRnZtdyix3pNX2IDQtJbLnNLc2ZlMmupbZQrtyxe3xl8+xlCnHByXZpzFty9sGzq3MozHHA9u9WsJEf5R7tnFDplNM1UitlTVTAyuCGk83brY4zjmz/52pUQyQHzjI_0_____Retail_CloudEdition
d4f9b41f-205c-405e-8e08-3d16e88e02be_J7NJW-V6KBM-CC8RW-Y29Y4-HQ2MJ_205_X23-15027_U9eyfIBXrs++lyP6OjHHaF/wjieAxQeSKwzSkGBeTTpyCDcenq8t4cKvqDHnauSZzaVPWNoVcASkMCdlJi3EkR29KSgvx9/K2OB8LVH2PPpqvwjm1ZZdrvLMGhW83A/KRrtN9AOx7bnPC8MNLErnzbRRS9/aOrmp4Uzo8EIVagI_0_OEM:NONSLP_IoTEnterpriseSK
) do (
for /f "tokens=1-10 delims=_" %%A in ("%%#") do (

if %1==key if %osSKU%==%%C (

REM Detect key attempt 1

if "%2"=="attempt1" if not defined key (
echo "!applist!" | find /i "%%A" 1>nul && (
if %%F==1 set notworking=1
set key=%%B
)
)

REM Detect key attempt 2

if "%2"=="attempt2" if not defined key (
set actidnotfound=1
set 9th=%%I
if not defined 9th (
if %%F==1 set notworking=1
set key=%%B
) else (
echo "%branch%" | find /i "%%I" 1>nul && (
if %%F==1 set notworking=1
set key=%%B
)
)
)
)

REM Generate ticket

if %1==ticket if "%key%"=="%%B" (
set _nil=
set "string=OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;Pfn=Microsoft.Windows.%%C.%%D_8wekyb3d8bbwe;DownlevelGenuineState=1;$([char]0)"
for /f "tokens=* delims=" %%i in ('powershell [convert]::ToBas!_nil!e64String([Text.Encoding]::Unicode.GetBytes("""!string!"""^)^)') do set "encoded=%%i"
echo "!encoded!" | find "AAAA" 1>nul || exit /b

<nul set /p "=<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=!encoded!;TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">%%E=</signature></signatures></genuineProperties></genuineAuthorization>" >"%tdir%\GenuineTicket"
)

)
)
exit /b

::========================================================================================================================================

::  Below code is used to get alternate edition name and key if current edition doesn't support HWID activation

::  ProfessionalCountrySpecific won't be converted because it's not a good idea to change CountrySpecific editions

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
125_EnterpriseS-2021___________cce9d2de-98ee-4ce2-8113-222620c64a27_ed655016-a9e8-4434-95d9-4345352c2552_QPM6N-7J2WJ-P88HH-P3YRH-YY74H_IoTEnterpriseS-2021
191_IoTEnterpriseS-Win11_______59eb965c-9150-42b7-a0ec-22151b9897c5_d4f9b41f-205c-405e-8e08-3d16e88e02be_J7NJW-V6KBM-CC8RW-Y29Y4-HQ2MJ_IoTEnterpriseSK-Win11
138_ProfessionalSingleLanguage_a48938aa-62fa-4966-9d44-9f04da3f72f2_4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7JG-NPHTM-C97JM-9MPGT-3V66T_Professional
) do (
for /f "tokens=1-6 delims=_" %%A in ("%%#") do if %osSKU%==%%A (
echo "!applist!" | find /i "%%C" 1>nul && (
echo "!applist!" | find /i "%%D" 1>nul && (
set altkey=%%E
set curedition=%%B
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