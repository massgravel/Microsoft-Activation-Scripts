@setlocal DisableDelayedExpansion
@echo off



::============================================================================
::
::   This script is a part of 'Microsoft_Activation_Scripts' (MAS) project.
::
::   Homepage: massgrave[.]dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================



::  To activate, run the script with "/KMS38" parameter or change 0 to 1 in below line
set _act=0

::  To remove KMS38 protection, run the script with /KMS38-RemoveProtection parameter or change 0 to 1 in below line
set _rem=0

::  To disable changing edition if current edition doesn't support KMS38 activation, change the value to 1 from 0 or run the script with "/KMS38-NoEditionChange" parameter
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
title  KMS38 Activation

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

set _k38=
set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
if %~z0 GEQ 200000 (set "_exitmsg=Go back") else (set "_exitmsg=Exit")
set "specific_kms=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f"

::========================================================================================================================================

if %winbuild% LSS 14393 (
%eline%
echo Unsupported OS version detected.
echo KMS38 Activation is supported for Windows 10/11/Server, build 14393 and later.
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
echo This script require admin privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto dk_done
)

::========================================================================================================================================

if %_rem%==1 goto :k_uninstall

:k_menu

if %_unattended%==0 (
cls
mode 76, 25
title  KMS38 Activation

echo:
echo:
echo:
echo:
echo         ____________________________________________________________
echo:
echo                 [1] KMS38 Activation
echo                 ____________________________________________
echo:
echo                 [2] Remove KM38 Protection
echo:
echo                 [0] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "              " %_Green% "Enter a menu option in the Keyboard [1,2,0]"
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
mode 102, 33
title  KMS38 Activation

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

set _eval=
set _evalserv=

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _eval=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" set _evalserv=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" set _eval=1 & set _evalserv=1

if defined _eval (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID 2>nul | find /i "Eval" 1>nul && (
%eline%
echo [%winos% ^| %winbuild%]
if defined _evalserv (
echo Server Evaluation cannot be activated. Convert it to full Server OS.
echo:
echo In MAS, goto Extras and use 'Change Edition' option.
) else (
echo Evaluation Editions cannot be activated. 
echo You need to install full version of %winos%
echo:
echo Download it from here,
echo https://massgrave.dev/genuine-installation-media.html
)
goto dk_done
)
)

::========================================================================================================================================

:: Check clipup.exe for the detection and activation of server cor and acor editions

set a_cor=
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" if not exist "%systemroot%\System32\clipup.exe" set a_cor=1

if defined a_cor (
if not exist "!_work!\clipup.exe" (
%eline%
echo clipup.exe doesn't exist in Server Cor/Acor [No GUI] version.
echo It's required for KMS38 Activation.
echo Check below page on how to activate it.
echo https://massgrave.dev/kms38.html
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

cls
echo:
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b
echo Checking OS Info                        [%winos% ^| %winbuild% ^| %arch%]

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

set "_serv=ClipSVC sppsvc KeyIso Winmgmt"

::  Client License Service (ClipSVC)
::  Software Protection
::  CNG Key Isolation
::  Windows Management Instrumentation

call :dk_errorcheck

::========================================================================================================================================

::  Check if GVLK (KMS key) is already installed or not

set _gvlk=
call :dk_channel
if /i "Volume:GVLK"=="%_channel%" set _gvlk=1

::  Detect Key

set key=
set pkey=
set altkey=
set changekey=
set curedition=
set altedition=

if defined applist call :kms38data getkey

if not defined key call :dk_gvlk %nul%
if defined applist if not defined key call :kms38fallback

if defined altkey (set key=%altkey%&set changekey=1)

if not defined key if defined notfoundaltactID (
call :dk_color %Red% "Checking Alternate Edition For KMS38    [%altedition% Activation ID Not Found]"
)

if not defined key if not defined _gvlk (
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
call :dk_color %Magenta% "[%altedition%] Edition product key will be used to enable KMS38 activation."
echo:
)

set _partial=
if not defined key (
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path SoftwareLicensingProduct where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null" Get PartialProductKey /value 2^>nul') do set "_partial=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT PartialProductKey FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL').Get()).PartialProductKey | %% {echo ('PartialProductKey='+$_)}" 2^>nul') do set "_partial=%%#"
call echo Checking Installed Product Key          [Partial Key - %%_partial%%] [Volume:GVLK]
)

set error_code=
if defined key (
if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%key%')" %nul%
if not !errorlevel!==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%
set error_code=!errorlevel!
cmd /c exit /b !error_code!
if !error_code! NEQ 0 set "error_code=[0x!=ExitCode!]"

if !error_code! EQU 0 (
call :dk_refresh
echo Installing KMS Client Setup Key         [%key%] [Successful]
) else (
call :dk_color %Red% "Installing KMS Client Setup Key         [%key%] [Failed] !error_code!"
if not defined error (
call :dk_color %Magenta% "In MAS, Goto Troubleshoot and run Fix Licensing option."
set showfix=1
)
set error=1
)
)

::========================================================================================================================================

::  Check activation ID for setting specific KMS host

set app=
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get ID /VALUE" 2^>nul') do call set "app=%%a"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL').Get()).ID | %% {echo ('ID='+$_)}" 2^>nul') do call set "app=%%a"

if not defined app (
call :dk_color %Red% "Checking Installed GVLK Activation ID   [Not Found] Aborting..."
goto :dk_done
)

::========================================================================================================================================

::  Set specific KMS host to Local Host
::  By doing this, global KMS IP can not replace KMS38 activation but can be used with Office and other Windows Editions

echo:
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f

%nul% reg query "HKLM\%specific_kms%" && (
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regdel\:.*';iex ($f[1]);"
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
::  Check https://massgrave.dev/kms38.html#Manual_Activation to see how it's generated

set "signature=C52iGEoH+1VqzI6kEAqOhUyrWuEObnivzaVjyef8WqItVYd/xGDTZZ3bkxAI9hTpobPFNJyJx6a3uriXq3HVd7mlXfSUK9ydeoUdG4eqMeLwkxeb6jQWJzLOz41rFVSMtBL0e+ycCATebTaXS4uvFYaDHDdPw2lKY8ADj3MLgsA="
set "sessionId=TwBTAE0AYQBqAG8AcgBWAGUAcgBzAGkAbwBuAD0ANQA7AE8AUwBNAGkAbgBvAHIAVgBlAHIAcwBpAG8AbgA9ADEAOwBPAFMAUABsAGEAdABmAG8AcgBtAEkAZAA9ADIAOwBQAFAAPQAwADsARwBWAEwASwBFAHgAcAA9ADIAMAAzADgALQAwADEALQAxADkAVAAwADMAOgAxADQAOgAwADcAWgA7AEQAbwB3AG4AbABlAHYAZQBsAEcAZQBuAHUAaQBuAGUAUwB0AGEAdABlAD0AMQA7AAAA"
<nul set /p "=<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=%sessionId%;TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">%signature%</signature></signatures></genuineProperties></genuineAuthorization>" >"%tdir%\GenuineTicket"

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "Generating GenuineTicket.xml            [Failed]"
if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
goto :k_final
) else (
echo Generating GenuineTicket.xml            [Successful]
)

set "_xmlexist=if exist "%tdir%\GenuineTicket.xml""

::  Stop sppsvc

net stop sppsvc /y %nul%
net stop sppsvc /y %nul%
net stop sppsvc /y %nul%

sc query sppsvc | find /i "1  STOPPED" %nul% && (
echo Stopping sppsvc Service                 [Successful]
) || (
call :dk_color %Red% "Stopping sppsvc Service                 [Failed]"
)

%_xmlexist% (
net stop ClipSVC /y %nul%
net start ClipSVC /y %nul%
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

%_xmlexist% (
set error=1
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
call :dk_color %Red% "Installing GenuineTicket.xml            [Failed With ClipSVC Service Restart, Wait...]"
)
)

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%
clipup -v -o

set rebuildinfo=

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

if defined applist if not defined showfix if defined rebuildinfo (
set showfix=1
call :dk_color %Magenta% "In MAS, Goto Troubleshoot and run Fix Licensing option."
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

if %_wmic% EQU 1 wmic path SoftwareLicensingProduct where ID='%app%' call ReArmsku %nul%
if %_wmic% EQU 0 %psc% "$null=([WMI]'SoftwareLicensingProduct=''%app%''').ReArmsku()" %nul%

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
if not defined error call :dk_color %Magenta% "In MAS, Goto Troubleshoot and run Fix Licensing option."
call :dk_color2 %Magenta% "Check this page for help" %_Yellow% " https://massgrave.dev/troubleshoot"

::========================================================================================================================================

:k_final

::  Remove the added Specific KMS Host (Local Host) if activation is not completed

echo:
if not defined _k38 (
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f
%nul% reg query "HKLM\%specific_kms%" && (
call :dk_color %Red% "Removing The Added Specific KMS Host    [Failed]"
) || (
echo Removing The Added Specific KMS Host    [Successful]
)
)

::  Protect KMS38 if opted by the user and conditions are correct

if defined _k38 (
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regdel\:.*';& ([ScriptBlock]::Create($f[1])) -protect;"
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg query "HKLM\%specific_kms%" && (
call :dk_color %Magenta% "Protect KMS38 By KMS                    [Successful] [Locked A Registry Key]"
) || (
call :dk_color %Red% "Protect KMS38 By KMS                    [Failed To Lock A Registry Key]"
)
)

::  clipup.exe does not exist in server cor and acor editions by default, it was copied there with this script

if defined a_cor if exist "%_clipup%" del /f /q "%_clipup%" %nul%

if defined a_cor (
if exist "%_clipup%" (
call :dk_color %Red% "Deleting copied clipup.exe file         [Failed]"
) else (
echo Deleting copied clipup.exe file         [Successful]
)
)

for %%# in (175 407) do if %osSKU%==%%# (
call :dk_color %Red% "%winos% does not support activation on non-azure platforms."
)

goto :dk_done

::========================================================================================================================================

:k_uninstall

cls
mode 99, 28
title  Remove KMS38 Protection

%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f

%nul% reg query "HKLM\%specific_kms%" && (
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regdel\:.*';iex ($f[1]);"
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
for /f "tokens=* delims=" %%# in ('%psc% "$([DateTime]::Now.addMinutes(%gpr%)).ToString('yyyy-MM-dd HH:mm:ss')" 2^>nul') do set "xpr=%%#"
call :dk_color %Green% "%winos% is activated till !xpr!"
exit /b

::  Check remaining KMS activation grace period

:k_checkexp

set gpr=0
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" 2^>nul') do set "gpr=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT GracePeriodRemaining FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL').Get()).GracePeriodRemaining | %% {echo ('GracePeriodRemaining='+$_)}" 2^>nul') do set "gpr=%%#"
if %gpr% GTR 259200 (set _k38=1) else (set _k38=)
exit /b

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

::  Get Windows installed key channel

:dk_channel

if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path SoftwareLicensingProduct where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null" Get ProductKeyChannel /value 2^>nul') do set "_channel=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT ProductKeyChannel FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL').Get()).ProductKeyChannel | %% {echo ('ProductKeyChannel='+$_)}" 2^>nul') do set "_channel=%%#"
exit /b

::  Get Windows Activation IDs

:dk_actids

set applist=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f') get ID /VALUE" 2^>nul')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f''').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))
exit /b

::  Check wmic.exe

:dk_ckeckwmic

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul && set _wmic=1
)
exit /b

::  Get Product name (WMI/REG methods are not reliable in all conditions, hence winbrand.dll method is used)

:dk_product

call :dk_reflection

set d1=%ref% $meth = $TypeBuilder.DefinePInvokeMethod('BrandingFormatString', 'winbrand.dll', 'Public, Static', 1, [String], @([String]), 1, 3);
set d1=%d1% $meth.SetImplementationFlags(128); $TypeBuilder.CreateType()::BrandingFormatString('%%WINDOWS_LONG%%')

set winos=
for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set winos=%%s)
echo "%winos%" | find /i "Windows" 1>nul || (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "winos=%%b"
if %winbuild% GEQ 22000 (
set winos=!winos:Windows 10=Windows 11!
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

:dk_pkey

call :dk_reflection

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SkuGetProductKeyForEdition', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([int], [String], [String].MakeByRefType(), [String].MakeByRefType()), 1, 3);
set d1=%d1% $out = ''; [void]$TypeBuilder.CreateType()::SkuGetProductKeyForEdition(%1, %2, [ref]$out, [ref]$null); $out

set pkey=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkey=%%a)
exit /b

::  Get channel name for the key which was extracted from pkeyhelper.dll

:dk_pkeychannel

set k=%1
set m=[Runtime.InteropServices.Marshal]
set p=%SystemRoot%\System32\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('PidGenX', 'pidgenx.dll', 'Public, Static', 1, [int], @([String], [String], [String], [int], [IntPtr], [IntPtr], [IntPtr]), 1, 3);
set d1=%d1% $r = [byte[]]::new(0x04F8); $r[0] = 0xF8; $r[1] = 0x04; $f = %m%::AllocHGlobal(0x04F8); %m%::Copy($r, 0, $f, 0x04F8);
set d1=%d1% [void]$TypeBuilder.CreateType()::PidGenX('%k%', '%p%', '00000', 0, 0, 0, $f); %m%::Copy($f, $r, 0, 0x04F8); %m%::FreeHGlobal($f); [Text.Encoding]::Unicode.GetString($r, 1016, 128)

set pkeychannel=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkeychannel=%%a)
exit /b

:dk_gvlk

for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b
for %%# in (Volume:GVLK) do (
call :dk_pkey %osSKU% '%%#'
if defined pkey call :dk_pkeychannel !pkey!
if /i [!pkeychannel!]==[%%#] (
set key=!pkey!
exit /b
)
)
exit /b

::========================================================================================================================================

:dk_errorcheck

::  Check disabled services

set serv_ste=
for %%# in (%_serv%) do (
set serv_dis=
reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v ImagePath %nul% || set serv_dis=1
for /f "skip=2 tokens=2*" %%a in ('reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v Start 2^>nul') do if /i %%b equ 0x4 set serv_dis=1
sc start %%# %nul%
if !errorlevel! EQU 1058 set serv_dis=1
sc query %%# %nul% || set serv_dis=1
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
if /i %%#==KeyIso         sc config %%# start= demand %nul%
if /i %%#==LicenseManager sc config %%# start= demand %nul%
if /i %%#==Winmgmt        sc config %%# start= auto %nul%
if /i %%#==wuauserv       sc config %%# start= demand %nul%
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
net start %%# /y %nul%
set errorcode=!errorlevel!
sc query %%# | find /i "4  RUNNING" %nul% || set checkerror=1

sc start %%# %nul%
if !errorlevel! NEQ 1056 if !errorlevel! NEQ 0 (set errorcode=!errorlevel!&set checkerror=1)
if defined checkerror if defined serv_e (set "serv_e=!serv_e!, %%#-!errorcode!") else (set "serv_e=%%#-!errorcode!")
)

if defined serv_e (
set error=1
call :dk_color %Red% "Starting Services                       [Failed] [%serv_e%]"
echo %serv_e% | findstr /i "ClipSVC-1058 sppsvc-1058" %nul% && (
call :dk_color %Magenta% "Restart the system to fix disabled service error 1058."
)
)

::========================================================================================================================================

::  Various error checks

if defined safeboot_option (
set error=1
call :dk_color2 %Red% "Checking Boot Mode                      " %Magenta% "[System is running in safe mode. Run in normal mode.]"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State" 2>nul | find /i "IMAGE_STATE_COMPLETE" 1>nul || (
set error=1
call :dk_color2 %Red% "Checking Audit Mode                     " %Magenta% "[System is running in Audit mode. Run in normal mode.]"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinPE" /v InstRoot %nul% && (
set error=1
call :dk_color2 %Red% "Checking WinPE                          " %Magenta% "[System is running in WinPE mode. Run in normal mode.]"
)


for %%# in (wmic.exe) do @if "%%~$PATH:#"=="" (
call :dk_color %Gray% "Checking WMIC.exe                       [Not Found]"
)


%psc% $ExecutionContext.SessionState.LanguageMode 2>nul | find /i "Full" 1>nul || (
set error=1
call :dk_color %Red% "Checking Powershell                     [Not Responding]"
)


set wmifailed=
if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" 2>nul | find /i "computersystem" 1>nul
if %errorlevel% NEQ 0 (
set error=1
set wmifailed=1
call :dk_color %Red% "Checking WMI                            [Not Responding] %_wmic%"
call :dk_color %Magenta% "In MAS, Goto Troubleshoot and run Fix WMI option."
)


if not "%regSKU%"=="%wmiSKU%" (
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
set error=1
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
call :dk_color2 %Red% "Checking Rearm                          " %Magenta% "[System Restart Is Required]"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState" %nul% && (
set error=1
call :dk_color2 %Red% "Checking ClipSVC                        " %Magenta% "[System Restart Is Required]"
)


for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" 2^>nul') do if /i %%b NEQ 0x0 (
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" /t REG_DWORD /d "0" /f %nul%
call :dk_color %Red% "Checking SkipRearm                      [Default 0 Value Not Found, Changing To 0]"
net stop sppsvc /y %nul%
net start sppsvc /y %nul%
set error=1
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Plugins\Objects\msft:rm/algorithm/hwid/4.0" /f ba02fed39662 /d %nul% || (
call :dk_color %Red% "Checking SPP Registry Key               [Incorrect ModuleId Found]"
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


set tokenstore=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore 2^>nul') do call set "tokenstore=%%b"
if not exist "%tokenstore%\" (
set error=1
call :dk_color %Red% "Checking SPP Token Folder               [Not Found] [%tokenstore%\]"
)


if exist "%tokenstore%\" if not exist "%tokenstore%\tokens.dat" (
set error=1
call :dk_color %Red% "Checking SPP tokens.dat                 [Not Found] [%tokenstore%\]"
)


if not exist %ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat (
set error=1
call :dk_color %Red% "Checking ClipSVC tokens.dat             [Not Found]"
)


if not exist %SystemRoot%\system32\sppsvc.exe (
set error=1
call :dk_color %Red% "Checking sppsvc.exe File                [Not Found]"
)


::  Below checks are performed if required services are not disabled + slmgr /dlv errorlevel is not Zero + Rearm restart is not required + WMI is working fine

set showfix=
set wpaerror=
set permerror=
if not defined serv_cste if /i not %error_code%==0 if /i not %error_code%==0xC004D302 if not defined wmifailed (

REM  This code checks for invalid registry keys in HKLM\SYSTEM\WPA. This issue may appear even on healthy systems.

if %winbuild% GEQ 14393 (
set /a count=0
for /f %%a in ('reg query "HKLM\SYSTEM\WPA" 2^>nul') do set /a count+=1
for /L %%# in (1,1,!count!) do (
reg query "HKLM\SYSTEM\WPA\8DEC0AF1-0341-4b93-85CD-72606C2DF94C-7P-%%#" /ve /t REG_BINARY %nul% || set wpaerror=1
)
if defined wpaerror call :dk_color %Red% "Checking WPA Registry Keys              [Error Found] [Registry Count - !count!]"
)

REM  This code checks if NT SERVICE\sppsvc has permission access to tokens folder and required registry keys. It's often caused by gaming spoofers. 

if not exist "%tokenstore%\" set permerror=1

for %%# in (
"%tokenstore%"
"HKLM:\SYSTEM\WPA"
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
) do if not defined permerror (
%psc% "$acl = Get-Acl '%%#'; if ($acl.Access.Where{ $_.IdentityReference -eq 'NT SERVICE\sppsvc' -and $_.AccessControlType -eq 'Deny' -or $acl.Access.IdentityReference -notcontains 'NT SERVICE\sppsvc'}) {Exit 2}" %nul%
if !errorlevel!==2 set permerror=1
)
if defined permerror call :dk_color %Red% "Checking SPP Permissions                [Error Found]"

set showfix=1
call :dk_color %Magenta% "In MAS, Goto Troubleshoot and run Fix Licensing option."
if not defined permerror call :dk_color %Magenta% "If activation still fails then run Fix WPA Registry option."
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
::  2nd column = GVLK (Generic volume licensing key)
::  3rd column = SKU ID
::  4th column = WMI Edition ID (For reference only)
::  5th column = Build Branch name incase same Edition ID is used in different OS versions with different key (For reference only)
::  Separator  = "_"

:kms38data

set f=
for %%# in (
73111121-5638-40f6-bc11-f1d7b0d64300_NPP%f%R9-FW%f%DCX-D2C%f%8J-H8%f%72K-2Y%f%T43___4_Enterprise
9bd77860-9b31-4b7b-96ad-2564017315bf_VDY%f%BN-27%f%WPP-V4H%f%QT-9V%f%MD4-VM%f%K7H___7_ServerStandard_FE
de32eafd-aaee-4662-9444-c1befb41bde2_N69%f%G4-B8%f%9J2-4G8%f%F4-WW%f%YCC-J4%f%64C___7_ServerStandard_RS5
8c1c5410-9f39-4805-8c9d-63a07706358f_WC2%f%BQ-8N%f%RM3-FDD%f%YY-2B%f%FGV-KH%f%KQY___7_ServerStandard_RS1
ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03_WX4%f%NM-KY%f%WYW-QJJ%f%R4-XV%f%3QB-6V%f%M33___8_ServerDatacenter_FE
34e1ae55-27f8-4950-8877-7a03be5fb181_WMD%f%GN-G9%f%PQG-XVV%f%XX-R3%f%X43-63%f%DFG___8_ServerDatacenter_RS5
21c56779-b449-4d20-adfc-eece0e1ad74b_CB7%f%KF-BW%f%N84-R7R%f%2Y-79%f%3K2-8X%f%DDG___8_ServerDatacenter_RS1
e272e3e2-732f-4c65-a8f0-484747d0d947_DPH%f%2V-TT%f%NVB-4X9%f%Q3-TJ%f%R4H-KH%f%JW4__27_EnterpriseN
2de67392-b7a7-462a-b1ca-108dd189f588_W26%f%9N-WF%f%GWX-YVC%f%9B-4J%f%6C9-T8%f%3GX__48_Professional
a80b5abf-76ad-428b-b05d-a47d2dffeebf_MH3%f%7W-N4%f%7XK-V7X%f%M9-C7%f%227-GC%f%QG9__49_ProfessionalN
034d3cbb-5d4b-4245-b3f8-f84571314078_WVD%f%HN-86%f%M7X-466%f%P6-VH%f%XV7-YY%f%726__50_ServerSolution_RS5
2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283_JCK%f%RF-N3%f%7P4-C2D%f%82-9Y%f%XRT-4M%f%63B__50_ServerSolution_RS1
7b9e1751-a8da-4f75-9560-5fadfe3d8e38_3KH%f%Y7-WN%f%T83-DGQ%f%KR-F7%f%HPR-84%f%4BM__98_CoreN
a9107544-f4a0-4053-a96a-1479abdef912_PVM%f%JN-6D%f%FY6-9CC%f%P6-7B%f%KTT-D3%f%WVR__99_CoreCountrySpecific
cd918a57-a41b-4c82-8dce-1a538e221a83_7HN%f%RX-D7%f%KGG-3K4%f%RQ-4W%f%PJ4-YT%f%DFH_100_CoreSingleLanguage
58e97c99-f377-4ef1-81d5-4ad5522b5fd8_TX9%f%XD-98%f%N7V-6WM%f%Q6-BX%f%7FG-H8%f%Q99_101_Core
7b4433f4-b1e7-4788-895a-c45378d38253_QN4%f%C6-GB%f%JD2-FB4%f%22-GH%f%WJK-GJ%f%G2R_110_ServerCloudStorage
8de8eb62-bbe0-40ac-ac17-f75595071ea3_GRF%f%BW-QN%f%DC4-6QB%f%HG-CC%f%K3B-2P%f%R88_120_ServerARM64_RS5
43d9af6e-5e86-4be8-a797-d072a046896c_K9F%f%YF-G6%f%NCK-73M%f%32-XM%f%VPY-F9%f%DRR_120_ServerARM64_RS4
e0c42288-980c-4788-a014-c080d2e1926e_NW6%f%C2-QM%f%PVW-D7K%f%KK-3G%f%KT6-VC%f%FB2_121_Education
3c102355-d027-42c6-ad23-2e7ef8a02585_2WH%f%4N-8Q%f%GBV-H22%f%JP-CT%f%43Q-MD%f%WWJ_122_EducationN
32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7X%f%TQ-FN%f%8P6-TTK%f%YV-9D%f%4CC-J4%f%62D_125_EnterpriseS_RS5,VB
2d5a5a60-3040-48bf-beb0-fcd770c20ce0_DCP%f%HK-NF%f%MTC-H88%f%MJ-PF%f%HPY-QJ%f%4BJ_125_EnterpriseS_RS1
7b51a46c-0c04-4e8f-9af4-8496cca90d5e_WNM%f%TR-4C%f%88C-JK8%f%YV-HQ%f%7T2-76%f%DF9_125_EnterpriseS_TH1
7103a333-b8c8-49cc-93ce-d37c09687f92_92N%f%FX-8D%f%JQP-P6B%f%BQ-TH%f%F9C-7C%f%G2H_126_EnterpriseSN_RS5,VB
9f776d83-7156-45b2-8a5c-359b9c9f22a3_QFF%f%DN-GR%f%T3P-VKW%f%WX-X7%f%T3R-8B%f%639_126_EnterpriseSN_RS1
87b838b7-41b6-4590-8318-5797951d8529_2F7%f%7B-TN%f%FGY-69Q%f%QF-B8%f%YKP-D6%f%9TJ_126_EnterpriseSN_TH1
39e69c41-42b4-4a0a-abad-8e3c10a797cc_QFN%f%D9-D3%f%Y9C-J3K%f%KY-6R%f%PVP-2D%f%PYV_145_ServerDatacenterACor_FE
90c362e5-0da1-4bfd-b53b-b87d309ade43_6NM%f%RW-2C%f%8FM-D24%f%W7-TQ%f%WMY-CW%f%H2D_145_ServerDatacenterACor_RS5
e49c08e7-da82-42f8-bde2-b570fbcae76c_2HX%f%DN-KR%f%XHB-GPY%f%C7-YC%f%KFJ-7F%f%VDG_145_ServerDatacenterACor_RS3
f5e9429c-f50b-4b98-b15c-ef92eb5cff39_67K%f%N8-4F%f%YJW-248%f%7Q-MQ%f%2J7-4C%f%4RG_146_ServerStandardACor_FE
73e3957c-fc0c-400d-9184-5f7b6f2eb409_N2K%f%JX-J9%f%4YW-TQV%f%FB-DG%f%9YT-72%f%4CC_146_ServerStandardACor_RS5
61c5ef22-f14f-4553-a824-c4b31e84b100_PTX%f%N8-JF%f%HJM-4WC%f%78-MP%f%CBR-9W%f%4KR_146_ServerStandardACor_RS3
82bbc092-bc50-4e16-8e18-b74fc486aec3_NRG%f%8B-VK%f%K3Q-CXV%f%CJ-9G%f%2XF-6Q%f%84J_161_ProfessionalWorkstation
4b1571d3-bafb-4b40-8087-a961be2caf65_9FN%f%HH-K3%f%HBT-3W4%f%TD-63%f%83H-6X%f%YWF_162_ProfessionalWorkstationN
3f1afc82-f8ac-4f6c-8005-1d233e606eee_6TP%f%4R-GN%f%PTD-KYY%f%HQ-7B%f%7DP-J4%f%47Y_164_ProfessionalEducation
5300b18c-2e33-4dc2-8291-47ffcec746dd_YVW%f%GF-BX%f%NMC-HTQ%f%YQ-CP%f%Q99-66%f%QFC_165_ProfessionalEducationN
8c8f0ad3-9a43-4e05-b840-93b8d1475cbc_6N3%f%79-GG%f%TMK-23C%f%6M-XV%f%VTC-CK%f%FRQ_168_ServerAzureCor_FE
a99cc1f0-7719-4306-9645-294102fbff95_FDN%f%H6-VW%f%9RW-BXP%f%J7-4X%f%TYG-23%f%9TB_168_ServerAzureCor_RS5
3dbf341b-5f6c-4fa7-b936-699dce9e263f_VP3%f%4G-4N%f%PPG-79J%f%TQ-86%f%4T4-R3%f%MQX_168_ServerAzureCor_RS1
e0b2d383-d112-413f-8a80-97f373a5820c_YYV%f%X9-NT%f%FWV-6MD%f%M3-9P%f%T4T-4M%f%68B_171_EnterpriseG
e38454fb-41a4-4f59-a5dc-25080e354730_44R%f%PN-FT%f%Y23-9VT%f%TB-MP%f%9BX-T8%f%4FV_172_EnterpriseGN
ec868e65-fadf-4759-b23e-93fe37f2cc29_CPW%f%HC-NT%f%2C7-VYW%f%78-DH%f%DB2-PG%f%3GK_175_ServerRdsh_RS5
e4db50ea-bda1-4566-b047-0ca50abc6f07_7NB%f%T4-WG%f%BQX-MP4%f%H7-QX%f%FF8-YP%f%3KX_175_ServerRdsh_RS3
0df4f814-3f57-4b8b-9a9d-fddadcd69fac_NBT%f%WJ-3D%f%R69-3C4%f%V8-C2%f%6MC-GQ%f%9M6_183_CloudE
59eb965c-9150-42b7-a0ec-22151b9897c5_KBN%f%8V-HF%f%GQ4-MGX%f%VD-34%f%7P6-PD%f%QGT_191_IoTEnterpriseS_NI
d30136fc-cb4b-416e-a23d-87207abc44a9_6XN%f%7V-PC%f%BDC-BDB%f%RH-8D%f%QY7-G6%f%R44_202_CloudEditionN
ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69_37D%f%7F-N4%f%9CB-WQR%f%8W-TB%f%J73-FM%f%8RX_203_CloudEdition
19b5e0fb-4431-46bc-bac1-2f1873e4ae73_NTB%f%V8-9K%f%7Q8-V27%f%C6-M2%f%BTV-KH%f%MXV_407_ServerTurbine
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do if %osSKU%==%%C (
if %1==getkey if not defined key echo "!applist!" | find /i "%%A" >nul && set key=%%B
)
)
exit /b

::========================================================================================================================================

::  Below code is used to get alternate edition name and key if current edition doesn't support KMS38 activation
::  ProfessionalCountrySpecific won't be converted because it's not a good idea to change CountrySpecific editions

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
188_IoTEnterprise_______________8ab9bdd1-1f67-4997-82d9-8878520837d9_73111121-5638-40f6-bc11-f1d7b0d64300_NPP%f%R9-FWD%f%CX-D2%f%C8J-H872%f%K-2Y%f%T43_Enterprise
191_IoTEnterpriseS-2021_________ed655016-a9e8-4434-95d9-4345352c2552_32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7X%f%TQ-FN8%f%P6-TT%f%KYV-9D4C%f%C-J4%f%62D_EnterpriseS-2021
205_IoTEnterpriseSK_____________d4f9b41f-205c-405e-8e08-3d16e88e02be_59eb965c-9150-42b7-a0ec-22151b9897c5_KBN%f%8V-HFG%f%Q4-MG%f%XVD-347P%f%6-PD%f%QGT_IoTEnterpriseS-Win11
138_ProfessionalSingleLanguage__a48938aa-62fa-4966-9d44-9f04da3f72f2_2de67392-b7a7-462a-b1ca-108dd189f588_W26%f%9N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX_Professional
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