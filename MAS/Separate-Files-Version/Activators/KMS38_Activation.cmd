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



::  To activate, run the script with "/KMS38" parameter or change 0 to 1 in below line
set _act=0

::  To remove KMS38 protection, run the script with /KMS38-RemoveProtection parameter or change 0 to 1 in below line
set _rem=0

::  To disable changing edition if current edition doesn't support KMS38 activation, change the value to 1 from 0 or run the script with "/KMS38-NoEditionChange" parameter
set _NoEditionChange=0

::  If value is changed in above lines or parameter is used then script will run in unattended mode



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
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "1" /f 1>nul
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

set _k38=
set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
if %~z0 GEQ 200000 (
set "_exitmsg=Go back"
set "_fixmsg=Go back to Main Menu, select Troubleshoot and run Fix Licensing option."
) else (
set "_exitmsg=Exit"
set "_fixmsg=In MAS folder, run Troubleshoot script and select Fix Licensing option."
)

set "specific_kms=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f"

::========================================================================================================================================

if %winbuild% LSS 14393 (
%eline%
echo Unsupported OS version detected [%winbuild%].
echo KMS38 Activation is supported for Windows 10/11/Server, build 14393 and later.
goto dk_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto dk_done
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
goto dk_done
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo To do so, right click on this script and select 'Run as administrator'.
goto dk_done
)

::========================================================================================================================================

::  This code disables QuickEdit for this cmd.exe session only without making permanent changes to the registry
::  It is added because clicking on the script window pauses the operation and leads to the confusion that script stopped due to an error

if %_unattended%==1 set quedit=1
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
if not %_unattended%==1 (
echo [1] Get Latest MAS
echo [0] Continue Anyway
echo:
call :dk_color %_Green% "Enter a menu option in the Keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
)
)
cls

::========================================================================================================================================

if %_rem%==1 goto :k_uninstall

:k_menu

if %_unattended%==0 (
cls
mode 76, 25
title  KMS38 Activation %masver%

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
mode 110, 34
if exist "%Systemdrive%\Windows\System32\spp\store_test\" mode 134, 34
title  KMS38 Activation %masver%

echo:
echo Initializing...

::  Check PowerShell

%psc% $ExecutionContext.SessionState.LanguageMode %nul2% | find /i "Full" %nul1% || (
%eline%
%psc% $ExecutionContext.SessionState.LanguageMode
echo:
echo PowerShell is not working. Aborting...
echo If you have applied restrictions on Powershell then undo those changes.
echo:
echo Check this page for help. %mas%troubleshoot
goto dk_done
)

::========================================================================================================================================

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
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% && (
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
echo %mas%genuine-installation-media.html
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
echo %mas%kms38.html
goto dk_done
)
)

::========================================================================================================================================

call :dk_checksku

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
for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if "%%j"=="" (set fullbuild=%%i) else (set fullbuild=%%i.%%j)
echo Checking OS Info                        [%winos% ^| %fullbuild% ^| %arch%]

::========================================================================================================================================

::  Check Windows Script Host

set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled %nul2% | find /i "0x0" %nul1% && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled %nul2% | find /i "0x0" %nul1% && (set _WSH=0)

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
set skufound=
set changekey=
set altedition=

call :kms38data getkey
if not defined key call :dk_gvlk %nul%
if defined applist if not defined key call :kms38fallback

if defined altkey (set key=%altkey%&set changekey=1)

set /a UBR=0
if %osSKU%==191 if defined altkey if defined altedition (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR 2^>nul') do if not errorlevel 1 set /a UBR=%%b
if %winbuild% GEQ 19044 if !UBR! LSS 2788 (
call :dk_color %Blue% "Windows must to be updated to build 19044.2788 or higher for IotEnterpriseS KMS38 activation."
)
)

if not defined key if defined notfoundaltactID (
call :dk_color %Red% "Checking Alternate Edition For KMS38    [%altedition% Activation ID Not Found]"
)

if not defined key if not defined _gvlk (
%eline%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
if not defined skufound (
echo Unable to find this product in the supported product list.
) else (
echo Required License files not installed.
)
echo Make sure you are using updated version of the script.
echo %mas%
echo:
goto dk_done
)

::========================================================================================================================================

::  Install key

echo:
if defined changekey (
call :dk_color %Blue% "[%altedition%] Edition product key will be used to enable KMS38 activation."
echo:
)

set _partial=
if not defined key (
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path SoftwareLicensingProduct where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null" Get PartialProductKey /value %nul6%') do set "_partial=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT PartialProductKey FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL').Get()).PartialProductKey | %% {echo ('PartialProductKey='+$_)}" %nul6%') do set "_partial=%%#"
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
call :dk_color %Blue% "%_fixmsg%"
set showfix=1
)
set error=1
)
)

::========================================================================================================================================

::  Check activation ID for setting specific KMS host

set app=
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get ID /VALUE" %nul6%') do call set "app=%%a"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL').Get()).ID | %% {echo ('ID='+$_)}" %nul6%') do call set "app=%%a"

if not defined app (
call :dk_color %Red% "Checking Installed GVLK Activation ID   [Not Found] Aborting..."
call :dk_color2 %Blue% "Check this page for help" %_Yellow% " %mas%troubleshoot"
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

if not exist %SystemRoot%\system32\ClipUp.exe (
call :dk_color %Red% "Checking ClipUp.exe File                [Not found, aborting the process]"
call :dk_color2 %Blue% "Check this page for help" %_Yellow% " %mas%troubleshoot"
goto :k_final
)

set "tdir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if not exist "%tdir%\" md "%tdir%\" %nul%

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" del /f /q "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" %nul%

::  Signature value is as it is, it's not encoded
::  Session ID is in Base64 encoded format. It's decoded value is "OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;GVLKExp=2038-01-19T03:14:07Z;DownlevelGenuineState=1;"
::  Check mass grave[.]dev/kms38.html#Manual_Activation to see how it's generated

set "signature=C52iGEoH+1VqzI6kEAqOhUyrWuEObnivzaVjyef8WqItVYd/xGDTZZ3bkxAI9hTpobPFNJyJx6a3uriXq3HVd7mlXfSUK9ydeoUdG4eqMeLwkxeb6jQWJzLOz41rFVSMtBL0e+ycCATebTaXS4uvFYaDHDdPw2lKY8ADj3MLgsA="
set "sessionId=TwBTAE0AYQBqAG8AcgBWAGUAcgBzAGkAbwBuAD0ANQA7AE8AUwBNAGkAbgBvAHIAVgBlAHIAcwBpAG8AbgA9ADEAOwBPAFMAUABsAGEAdABmAG8AcgBtAEkAZAA9ADIAOwBQAFAAPQAwADsARwBWAEwASwBFAHgAcAA9ADIAMAAzADgALQAwADEALQAxADkAVAAwADMAOgAxADQAOgAwADcAWgA7AEQAbwB3AG4AbABlAHYAZQBsAEcAZQBuAHUAaQBuAGUAUwB0AGEAdABlAD0AMQA7AAAA"
<nul set /p "=<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=%sessionId%;TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">%signature%</signature></signatures></genuineProperties></genuineAuthorization>" >"%tdir%\GenuineTicket"

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "Generating GenuineTicket.xml            [Failed, aborting the process]"
if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
goto :k_final
) else (
echo Generating GenuineTicket.xml            [Successful]
)

set "_xmlexist=if exist "%tdir%\GenuineTicket.xml""

::  Stop sppsvc

%psc% Stop-Service sppsvc %nul%

sc query sppsvc | find /i "STOPPED" %nul% && (
echo Stopping sppsvc Service                 [Successful]
) || (
call :dk_color %Gray% "Stopping sppsvc Service                 [Failed]"
)

%_xmlexist% (
%psc% Restart-Service ClipSVC %nul%
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

if defined applist if not defined showfix if defined rebuildinfo (
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
if not defined error call :dk_color %Blue% "%_fixmsg%"
call :dk_color2 %Blue% "Check this page for help" %_Yellow% " %mas%troubleshoot"

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
echo Protect KMS38 From KMS                  [Successful] [Locked A Registry Key]
) || (
call :dk_color %Red% "Protect KMS38 From KMS                  [Failed To Lock A Registry Key]"
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
title  Remove KMS38 Protection %masver%

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

::  Check SKU value

:dk_checksku

set osSKU=
set slcSKU=
set wmiSKU=
set regSKU=

if %winbuild% GEQ 14393 (set info=Kernel-BrandingInfo) else (set info=Kernel-ProductInfo)
set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SLGetWindowsInformationDWORD', 'slc.dll', 'Public, Static', 1, [int], @([String], [int].MakeByRefType()), 1, 3);
set d1=%d1% $Sku = 0; [void]$TypeBuilder.CreateType()::SLGetWindowsInformationDWORD('%info%', [ref]$Sku); $Sku
for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set slcSKU=%%s)
if "%slcSKU%"=="0" set slcSKU=
if 1%slcSKU% NEQ +1%slcSKU% set slcSKU=

for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn %nul6%') do set "regSKU=%%a"
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" %nul6%') do if not errorlevel 1 set "wmiSKU=%%a"
if %_wmic% EQU 0 for /f "tokens=1" %%a in ('%psc% "([WMI]'Win32_OperatingSystem=@').OperatingSystemSKU" %nul6%') do if not errorlevel 1 set "wmiSKU=%%a"

set osSKU=%slcSKU%
if not defined osSKU set osSKU=%wmiSKU%
if not defined osSKU set osSKU=%regSKU%
exit /b

::  Check KMS activation status

:k_actinfo

set xpr=
for /f "tokens=* delims=" %%# in ('%psc% "$([DateTime]::Now.addMinutes(%gpr%)).ToString('yyyy-MM-dd HH:mm:ss')" %nul6%') do set "xpr=%%#"
call :dk_color %Green% "%winos% is activated till !xpr!"
exit /b

::  Check remaining KMS activation grace period

:k_checkexp

set gpr=0
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" %nul6%') do set "gpr=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT GracePeriodRemaining FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL').Get()).GracePeriodRemaining | %% {echo ('GracePeriodRemaining='+$_)}" %nul6%') do set "gpr=%%#"
if %gpr% GTR 259200 (set _k38=1) else (set _k38=)
exit /b

::  Get Windows permanent activation status

:dk_checkperm

if %_wmic% EQU 1 wmic path SoftwareLicensingProduct where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name /value %nul2% | findstr /i "Windows" %nul1% && set _perm=1||set _perm=
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Name FROM SoftwareLicensingProduct WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL').Get()).Name | %% {echo ('Name='+$_)}" %nul2% | findstr /i "Windows" %nul1% && set _perm=1||set _perm=
exit /b

::  Refresh license status

:dk_refresh

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call RefreshLicenseStatus %nul%
if %_wmic% EQU 0 %psc% "$null=(([WMICLASS]'SoftwareLicensingService').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::  Get Windows installed key channel

:dk_channel

if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path SoftwareLicensingProduct where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null" Get ProductKeyChannel /value %nul6%') do set "_channel=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT ProductKeyChannel FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL').Get()).ProductKeyChannel | %% {echo ('ProductKeyChannel='+$_)}" %nul6%') do set "_channel=%%#"
exit /b

::  Get Windows Activation IDs

:dk_actids

set applist=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f') get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f''').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))
exit /b

::  Check wmic.exe

:dk_ckeckwmic

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value %nul2% | find /i "computersystem" %nul1% && set _wmic=1
)
exit /b

::  Get Product name (WMI/REG methods are not reliable in all conditions, hence winbrand.dll method is used)

:dk_product

call :dk_reflection

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

set showfix=

::  Check corrupt services

set serv_cor=
for %%# in (%_serv%) do (
set _corrupt=
sc start %%# %nul%
if !errorlevel! EQU 1060 set _corrupt=1
sc query %%# %nul% || set _corrupt=1
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type) do if not defined _corrupt (
reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v %%G %nul% || set _corrupt=1
if /i %%#==TrustedInstaller if /i %%G==DependOnService set _corrupt=
)

if defined _corrupt (if defined serv_cor (set "serv_cor=!serv_cor! %%#") else (set "serv_cor=%%#"))
)

if defined serv_cor (
set error=1
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
if /i %%#==DoSvc            sc config %%# start= delayed-auto %nul%
if /i %%#==UsoSvc           sc config %%# start= delayed-auto %nul%
if /i %%#==CryptSvc         sc config %%# start= auto %nul%
if /i %%#==BITS             sc config %%# start= delayed-auto %nul%
if /i %%#==wuauserv         sc config %%# start= demand %nul%
if /i %%#==WaaSMedicSvc     sc config %%# start= demand %nul%
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
%psc% Start-Service %%# %nul%
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
call :dk_color %Blue% "Restart the system to fix this error."
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


for /f "skip=2 tokens=2*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State" /v ImageState') do (set imagestate=%%B)
if /i not "%imagestate%"=="IMAGE_STATE_COMPLETE" (
set error=1
call :dk_color %Red% "Checking Windows Setup State            [%imagestate%]"
echo "%imagestate%" | find /i "RESEAL" %nul% && (
set showfix=1
call :dk_color %Blue% "You need to run it in normal mode in case you are running it in Audit Mode."
)
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinPE" /v InstRoot %nul% && (
set error=1
set showfix=1
call :dk_color2 %Red% "Checking WinPE                          " %Blue% "[WinPE mode found. Run in normal mode.]"
)


set wpainfo=
set wpaerror=
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':wpatest\:.*';iex ($f[1]);" %nul6%') do (set wpainfo=%%a)
echo "%wpainfo%" | find /i "Error Found" %nul% && (
set error=1
set wpaerror=1
call :dk_color %Red% "Checking WPA Registry Error             [%wpainfo%]"
) || (
echo Checking WPA Registry Count             [%wpainfo%]
)


DISM /English /Online /Get-CurrentEdition %nul%
set dism_error=%errorlevel%
cmd /c exit /b %dism_error%
if %dism_error% NEQ 0 set "dism_error=0x%=ExitCode%"
if %dism_error% NEQ 0 (
call :dk_color %Red% "Checking DISM                           [Not Responding] [%dism_error%]"
)


if not defined officeact if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
set error=1
set showfix=1
call :dk_color %Red% "Checking Eval Packages                  [Non-Eval Licenses are installed in Eval Windows]"
call :dk_color %Blue% "Evaluation Windows can not be activated and different License install may lead to errors."
call :dk_color %Blue% "It is recommended to install full version of %winos%."
call :dk_color %Blue% "You can download it from %mas%genuine-installation-media.html"
)


set osedition=
for /f "skip=2 tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul6%') do set "osedition=%%a"

::  Workaround for an issue in builds between 1607 and 1709 where ProfessionalEducation is shown as Professional

if "%osSKU%"=="164" set osedition=ProfessionalEducation
if "%osSKU%"=="165" set osedition=ProfessionalEducationN

if not defined officeact (
if not defined osedition (
call :dk_color %Red% "Checking Edition Name                   [Not Found In Registry]"
) else (

if not exist "%SystemRoot%\System32\spp\tokens\skus\%osedition%\%osedition%*.xrm-ms" (
set error=1
call :dk_color %Red% "Checking License Files                  [Not Found] [%osedition%]"
)

if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*-%osedition%-*.mum" (
set error=1
call :dk_color %Red% "Checking Package File                   [Not Found] [%osedition%]"
)
)
)


cscript //nologo %windir%\system32\slmgr.vbs /dlv %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 set "error_code=0x%=ExitCode%"
if %error_code% NEQ 0 (
set error=1
call :dk_color %Red% "Checking slmgr /dlv                     [Not Responding] %error_code%"
)


for %%# in (wmic.exe) do @if "%%~$PATH:#"=="" (
call :dk_color %Gray% "Checking WMIC.exe                       [Not Found]"
)


set wmifailed=
if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value %nul2% | find /i "computersystem" %nul1%
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" %nul2% | find /i "computersystem" %nul1%

if %errorlevel% NEQ 0 set wmifailed=1
echo "%error_code%" | findstr /i "0x800410 0x800440" %nul1% && set wmifailed=1& ::  https://learn.microsoft.com/en-us/windows/win32/wmisdk/wmi-error-constants
if defined wmifailed (
set error=1
call :dk_color %Red% "Checking WMI                            [Not Responding]"
call :dk_color %Blue% "In MAS, Goto Troubleshoot and run Fix WMI option."
set showfix=1
)


%nul% set /a "sum=%slcSKU%+%regSKU%+%wmiSKU%"
set /a "sum/=3"
if not defined officeact if not "%sum%"=="%slcSKU%" (
call :dk_color %Red% "Checking SLC/WMI/REG SKU                [Difference Found - SLC:%slcSKU% WMI:%wmiSKU% Reg:%regSKU%]"
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


for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" %nul6%') do if /i %%b NEQ 0x0 (
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" /t REG_DWORD /d "0" /f %nul%
call :dk_color %Red% "Checking SkipRearm                      [Default 0 Value Not Found. Changing To 0]"
%psc% Restart-Service sppsvc %nul%
set error=1
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Plugins\Objects\msft:rm/algorithm/hwid/4.0" /f ba02fed39662 /d %nul% || (
call :dk_color %Red% "Checking SPP Registry Key               [Incorrect ModuleId Found]"
call :dk_color %Blue% "Possibly Caused By Gaming Spoofers. Help: %mas%troubleshoot"
set error=1
set showfix=1
)


set tokenstore=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore %nul6%') do call set "tokenstore=%%b"
if not exist "%tokenstore%\" (
set error=1
REM This code creates token folder only if it's missing and sets default permission for it
mkdir "%tokenstore%" %nul%
set "d=$sddl = 'O:BAG:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICIIO;GR;;;BU)(A;;FR;;;BU)(A;OICI;FA;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)';"
set "d=!d! $AclObject = New-Object System.Security.AccessControl.DirectorySecurity;"
set "d=!d! $AclObject.SetSecurityDescriptorSddlForm($sddl);"
set "d=!d! Set-Acl -Path %tokenstore% -AclObject $AclObject;"
%psc% "!d!" %nul%
call :dk_color %Gray% "Checking SPP Token Folder               [Not Found. Creating Now] [%tokenstore%\]"
)


call :dk_actids
if not defined applist (
%psc% Stop-Service sppsvc %nul%
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if !errorlevel! NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
call :dk_refresh
call :dk_actids
if not defined applist (
set error=1
call :dk_color %Red% "Checking Activation IDs                 [Not Found]"
)
)


if exist "%tokenstore%\" if not exist "%tokenstore%\tokens.dat" (
set error=1
call :dk_color %Red% "Checking SPP tokens.dat                 [Not Found] [%tokenstore%\]"
)


if not exist %SystemRoot%\system32\sppsvc.exe (
set error=1
set showfix=1
call :dk_color %Red% "Checking sppsvc.exe File                [Not Found]"
)


::  This code checks if NT SERVICE\sppsvc has permission access to tokens folder and required registry keys. It's often caused by gaming spoofers. 

set permerror=
if not exist "%tokenstore%\" set permerror=1

for %%# in (
"%tokenstore%"
"HKLM:\SYSTEM\WPA"
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
) do if not defined permerror (
%psc% "$acl = Get-Acl '%%#'; if ($acl.Access.Where{ $_.IdentityReference -eq 'NT SERVICE\sppsvc' -and $_.AccessControlType -eq 'Deny' -or $acl.Access.IdentityReference -notcontains 'NT SERVICE\sppsvc'}) {Exit 2}" %nul%
if !errorlevel!==2 set permerror=1
)
if defined permerror (
set error=1
set showfix=1
call :dk_color %Red% "Checking SPP Permissions                [Error Found]"
call :dk_color %Blue% "%_fixmsg%"
)


::  If required services are not disabled or corrupted + if there is any error + slmgr /dlv errorlevel is not Zero + no fix was shown before

if not defined serv_cor if not defined serv_cste if defined error if /i not %error_code%==0 if not defined showfix (
set showfix=1
call :dk_color %Blue% "%_fixmsg%"
if not defined permerror call :dk_color %Blue% "If activation still fails then run Fix WPA Registry option."
)

if not defined showfix if defined wpaerror (
set showfix=1
call :dk_color %Blue% "If activation fails then go back to Main Menu, select Troubleshoot and run Fix WPA Registry option."
)

exit /b

::  This code checks for invalid registry keys in HKLM\SYSTEM\WPA. This issue may appear even on healthy systems

:wpatest:
$wpaKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey("SYSTEM\\WPA")
$count = $wpaKey.SubKeyCount

$osVersion = [System.Environment]::OSVersion.Version
$minBuildNumber = 14393

if ($osVersion.Build -ge $minBuildNumber) {
    $subkeyHashTable = @{}
    foreach ($subkeyName in $wpaKey.GetSubKeyNames()) {
        $keyNumber = $subkeyName -replace '.*-', ''
        $subkeyHashTable[$keyNumber] = $true
    }
    for ($i=1; $i -le $count; $i++) {
        if (-not $subkeyHashTable.ContainsKey("$i")) {
            Write-Host "Total Keys $count. Error Found- $i key does not exist"
			$wpaKey.Close()
            exit
        }
    }
}
$wpaKey.GetSubKeyNames() | ForEach-Object {
    $subkey = $wpaKey.OpenSubKey($_)
    $p = $subkey.GetValueNames()
    if (($p | Where-Object { $subkey.GetValueKind($_) -eq [Microsoft.Win32.RegistryValueKind]::Binary }).Count -eq 0) {
        Write-Host "Total Keys $count. Error Found- Binary Data is corrupt"
		$wpaKey.Close()
        exit
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
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
pause %nul1%
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
73111121-5638-40f6-bc11-f1d7b0d64300_NP%f%PR9-FWD%f%CX-D2%f%C8J-H872%f%K-2Y%f%T43___4_Enterprise
7dc26449-db21-4e09-ba37-28f2958506a6_DP%f%NXD-67Y%f%Y9-WW%f%FJJ-RYH9%f%9-RM%f%832___7_ServerStandard_Ge
9bd77860-9b31-4b7b-96ad-2564017315bf_VD%f%YBN-27W%f%PP-V4%f%HQT-9VMD%f%4-VM%f%K7H___7_ServerStandard_FE
de32eafd-aaee-4662-9444-c1befb41bde2_N6%f%9G4-B89%f%J2-4G%f%8F4-WWYC%f%C-J4%f%64C___7_ServerStandard_RS5
8c1c5410-9f39-4805-8c9d-63a07706358f_WC%f%2BQ-8NR%f%M3-FD%f%DYY-2BFG%f%V-KH%f%KQY___7_ServerStandard_RS1
c052f164-cdf6-409a-a0cb-853ba0f0f55a_CN%f%FDQ-2BW%f%8H-9V%f%4WM-TKCP%f%D-MD%f%2QF___8_ServerDatacenter_Ge
ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03_WX%f%4NM-KYW%f%YW-QJ%f%JR4-XV3Q%f%B-6V%f%M33___8_ServerDatacenter_FE
34e1ae55-27f8-4950-8877-7a03be5fb181_WM%f%DGN-G9P%f%QG-XV%f%VXX-R3X4%f%3-63%f%DFG___8_ServerDatacenter_RS5
21c56779-b449-4d20-adfc-eece0e1ad74b_CB%f%7KF-BWN%f%84-R7%f%R2Y-793K%f%2-8X%f%DDG___8_ServerDatacenter_RS1
e272e3e2-732f-4c65-a8f0-484747d0d947_DP%f%H2V-TTN%f%VB-4X%f%9Q3-TJR4%f%H-KH%f%JW4__27_EnterpriseN
2de67392-b7a7-462a-b1ca-108dd189f588_W2%f%69N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX__48_Professional
a80b5abf-76ad-428b-b05d-a47d2dffeebf_MH%f%37W-N47%f%XK-V7%f%XM9-C722%f%7-GC%f%QG9__49_ProfessionalN
034d3cbb-5d4b-4245-b3f8-f84571314078_WV%f%DHN-86M%f%7X-46%f%6P6-VHXV%f%7-YY%f%726__50_ServerSolution_RS5
2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283_JC%f%KRF-N37%f%P4-C2%f%D82-9YXR%f%T-4M%f%63B__50_ServerSolution_RS1
7b9e1751-a8da-4f75-9560-5fadfe3d8e38_3K%f%HY7-WNT%f%83-DG%f%QKR-F7HP%f%R-84%f%4BM__98_CoreN
a9107544-f4a0-4053-a96a-1479abdef912_PV%f%MJN-6DF%f%Y6-9C%f%CP6-7BKT%f%T-D3%f%WVR__99_CoreCountrySpecific
cd918a57-a41b-4c82-8dce-1a538e221a83_7H%f%NRX-D7K%f%GG-3K%f%4RQ-4WPJ%f%4-YT%f%DFH_100_CoreSingleLanguage
58e97c99-f377-4ef1-81d5-4ad5522b5fd8_TX%f%9XD-98N%f%7V-6W%f%MQ6-BX7F%f%G-H8%f%Q99_101_Core
7b4433f4-b1e7-4788-895a-c45378d38253_QN%f%4C6-GBJ%f%D2-FB%f%422-GHWJ%f%K-GJ%f%G2R_110_ServerCloudStorage
8de8eb62-bbe0-40ac-ac17-f75595071ea3_GR%f%FBW-QND%f%C4-6Q%f%BHG-CCK3%f%B-2P%f%R88_120_ServerARM64_RS5
43d9af6e-5e86-4be8-a797-d072a046896c_K9%f%FYF-G6N%f%CK-73%f%M32-XMVP%f%Y-F9%f%DRR_120_ServerARM64_RS4
e0c42288-980c-4788-a014-c080d2e1926e_NW%f%6C2-QMP%f%VW-D7%f%KKK-3GKT%f%6-VC%f%FB2_121_Education
3c102355-d027-42c6-ad23-2e7ef8a02585_2W%f%H4N-8QG%f%BV-H2%f%2JP-CT43%f%Q-MD%f%WWJ_122_EducationN
32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7%f%XTQ-FN8%f%P6-TT%f%KYV-9D4C%f%C-J4%f%62D_125_EnterpriseS_RS5,VB,Ge
2d5a5a60-3040-48bf-beb0-fcd770c20ce0_DC%f%PHK-NFM%f%TC-H8%f%8MJ-PFHP%f%Y-QJ%f%4BJ_125_EnterpriseS_RS1
7b51a46c-0c04-4e8f-9af4-8496cca90d5e_WN%f%MTR-4C8%f%8C-JK%f%8YV-HQ7T%f%2-76%f%DF9_125_EnterpriseS_TH1
7103a333-b8c8-49cc-93ce-d37c09687f92_92%f%NFX-8DJ%f%QP-P6%f%BBQ-THF9%f%C-7C%f%G2H_126_EnterpriseSN_RS5,VB,Ge
9f776d83-7156-45b2-8a5c-359b9c9f22a3_QF%f%FDN-GRT%f%3P-VK%f%WWX-X7T3%f%R-8B%f%639_126_EnterpriseSN_RS1
87b838b7-41b6-4590-8318-5797951d8529_2F%f%77B-TNF%f%GY-69%f%QQF-B8YK%f%P-D6%f%9TJ_126_EnterpriseSN_TH1
39e69c41-42b4-4a0a-abad-8e3c10a797cc_QF%f%ND9-D3Y%f%9C-J3%f%KKY-6RPV%f%P-2D%f%PYV_145_ServerDatacenterACor_FE
90c362e5-0da1-4bfd-b53b-b87d309ade43_6N%f%MRW-2C8%f%FM-D2%f%4W7-TQWM%f%Y-CW%f%H2D_145_ServerDatacenterACor_RS5
e49c08e7-da82-42f8-bde2-b570fbcae76c_2H%f%XDN-KRX%f%HB-GP%f%YC7-YCKF%f%J-7F%f%VDG_145_ServerDatacenterACor_RS3
f5e9429c-f50b-4b98-b15c-ef92eb5cff39_67%f%KN8-4FY%f%JW-24%f%87Q-MQ2J%f%7-4C%f%4RG_146_ServerStandardACor_FE
73e3957c-fc0c-400d-9184-5f7b6f2eb409_N2%f%KJX-J94%f%YW-TQ%f%VFB-DG9Y%f%T-72%f%4CC_146_ServerStandardACor_RS5
61c5ef22-f14f-4553-a824-c4b31e84b100_PT%f%XN8-JFH%f%JM-4W%f%C78-MPCB%f%R-9W%f%4KR_146_ServerStandardACor_RS3
82bbc092-bc50-4e16-8e18-b74fc486aec3_NR%f%G8B-VKK%f%3Q-CX%f%VCJ-9G2X%f%F-6Q%f%84J_161_ProfessionalWorkstation
4b1571d3-bafb-4b40-8087-a961be2caf65_9F%f%NHH-K3H%f%BT-3W%f%4TD-6383%f%H-6X%f%YWF_162_ProfessionalWorkstationN
3f1afc82-f8ac-4f6c-8005-1d233e606eee_6T%f%P4R-GNP%f%TD-KY%f%YHQ-7B7D%f%P-J4%f%47Y_164_ProfessionalEducation
5300b18c-2e33-4dc2-8291-47ffcec746dd_YV%f%WGF-BXN%f%MC-HT%f%QYQ-CPQ9%f%9-66%f%QFC_165_ProfessionalEducationN
45b5aff2-60a0-42f2-bc4b-ec6e5f7b527e_QN%f%7G3-4RM%f%92-MT%f%6QR-PR96%f%6-FV%f%YV7_168_ServerAzureCor_Ge
8c8f0ad3-9a43-4e05-b840-93b8d1475cbc_6N%f%379-GGT%f%MK-23%f%C6M-XVVT%f%C-CK%f%FRQ_168_ServerAzureCor_FE
a99cc1f0-7719-4306-9645-294102fbff95_FD%f%NH6-VW9%f%RW-BX%f%PJ7-4XTY%f%G-23%f%9TB_168_ServerAzureCor_RS5
3dbf341b-5f6c-4fa7-b936-699dce9e263f_VP%f%34G-4NP%f%PG-79%f%JTQ-864T%f%4-R3%f%MQX_168_ServerAzureCor_RS1
e0b2d383-d112-413f-8a80-97f373a5820c_YY%f%VX9-NTF%f%WV-6M%f%DM3-9PT4%f%T-4M%f%68B_171_EnterpriseG
e38454fb-41a4-4f59-a5dc-25080e354730_44%f%RPN-FTY%f%23-9V%f%TTB-MP9B%f%X-T8%f%4FV_172_EnterpriseGN
ec868e65-fadf-4759-b23e-93fe37f2cc29_CP%f%WHC-NT2%f%C7-VY%f%W78-DHDB%f%2-PG%f%3GK_175_ServerRdsh_RS5
e4db50ea-bda1-4566-b047-0ca50abc6f07_7N%f%BT4-WGB%f%QX-MP%f%4H7-QXFF%f%8-YP%f%3KX_175_ServerRdsh_RS3
0df4f814-3f57-4b8b-9a9d-fddadcd69fac_NB%f%TWJ-3DR%f%69-3C%f%4V8-C26M%f%C-GQ%f%9M6_183_CloudE
59eb965c-9150-42b7-a0ec-22151b9897c5_KB%f%N8V-HFG%f%Q4-MG%f%XVD-347P%f%6-PD%f%QGT_191_IoTEnterpriseS_VB,NI
d30136fc-cb4b-416e-a23d-87207abc44a9_6X%f%N7V-PCB%f%DC-BD%f%BRH-8DQY%f%7-G6%f%R44_202_CloudEditionN
ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69_37%f%D7F-N49%f%CB-WQ%f%R8W-TBJ7%f%3-FM%f%8RX_203_CloudEdition
c2e946d1-cfa2-4523-8c87-30bc696ee584_NQ%f%8HH-FTD%f%TM-6V%f%GY7-TQ3D%f%V-XF%f%BV2_407_ServerTurbine_Ge
19b5e0fb-4431-46bc-bac1-2f1873e4ae73_NT%f%BV8-9K7%f%Q8-V2%f%7C6-M2BT%f%V-KH%f%MXV_407_ServerTurbine_RS5
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do if %osSKU%==%%C (
set skufound=1
if %1==getkey if not defined key echo "!applist!" | find /i "%%A" %nul1% && set key=%%B
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
188_IoTEnterprise__________________8ab9bdd1-1f67-4997-82d9-8878520837d9_73111121-5638-40f6-bc11-f1d7b0d64300_NPP%f%R9-FWD%f%CX-D2%f%C8J-H872%f%K-2Y%f%T43_Enterprise
206_IoTEnterpriseK_________________80083eae-7031-4394-9e88-4901973d56fe_73111121-5638-40f6-bc11-f1d7b0d64300_NPP%f%R9-FWD%f%CX-D2%f%C8J-H872%f%K-2Y%f%T43_Enterprise
191_IoTEnterpriseS-2021____________ed655016-a9e8-4434-95d9-4345352c2552_32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7X%f%TQ-FN8%f%P6-TT%f%KYV-9D4C%f%C-J4%f%62D_EnterpriseS-2021
205_IoTEnterpriseSK________________d4f9b41f-205c-405e-8e08-3d16e88e02be_59eb965c-9150-42b7-a0ec-22151b9897c5_KBN%f%8V-HFG%f%Q4-MG%f%XVD-347P%f%6-PD%f%QGT_IoTEnterpriseS
138_ProfessionalSingleLanguage_____a48938aa-62fa-4966-9d44-9f04da3f72f2_2de67392-b7a7-462a-b1ca-108dd189f588_W26%f%9N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX_Professional
139_ProfessionalCountrySpecific____f7af7d09-40e4-419c-a49b-eae366689ebd_2de67392-b7a7-462a-b1ca-108dd189f588_W26%f%9N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX_Professional
139_ProfessionalCountrySpecific-Zn_01eb852c-424d-4060-94b8-c10d799d7364_2de67392-b7a7-462a-b1ca-108dd189f588_W26%f%9N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX_Professional
) do (
for /f "tokens=1-6 delims=_" %%A in ("%%#") do if %osSKU%==%%A (
echo "!applist!" | find /i "%%C" %nul1% && (
echo "!applist!" | find /i "%%D" %nul1% && (
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
