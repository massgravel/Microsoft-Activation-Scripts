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



::  To activate Office with Ohook activation, run the script with "/Ohook" parameter or change 0 to 1 in below line
set _act=0

::  To remove Ohook activation, run the script with /Ohook-Uninstall parameter or change 0 to 1 in below line
set _rem=0

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
set "mas=mass%blank%grave.dev"

::  Check if Null service is working, it's important for the batch script

sc query Null | find /i "RUNNING"
if %errorlevel% NEQ 0 (
echo:
echo Null service is not running, script may crash...
echo:
echo:
echo Help - https://%mas%/troubleshoot.html
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
title  Ohook Activation

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

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
if %~z0 GEQ 200000 (
set "_exitmsg=Go back"
set "_fixmsg=Go back to Main Menu, select Troubleshoot and run Fix Licensing option."
) else (
set "_exitmsg=Exit"
set "_fixmsg=In MAS folder, run Troubleshoot script and select Fix Licensing option."
)

::========================================================================================================================================

if %winbuild% LSS 9200 (
%eline%
echo Unsupported OS version detected [%winbuild%].
echo Ohook Activation is supported on Windows 8 and later and their server equivalent.
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
echo This script require admin privileges.
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

if %_rem%==1 goto :oh_uninstall

:oh_menu

if %_unattended%==0 (
cls
mode 76, 25
title  Ohook Activation

echo:
echo:
echo:
echo:
echo         ____________________________________________________________
echo:
echo                 [1] Install Ohook Office Activation
echo:
echo                 [2] Uninstall
echo                 ____________________________________________
echo:
echo                 [3] Download Office
echo:
echo                 [0] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "              " %_Green% "Enter a menu option in the Keyboard [1,2,3,0]"
choice /C:1230 /N
set _el=!errorlevel!
if !_el!==4  exit /b
if !_el!==3  start https://%mas%/genuine-installation-media.html &goto :oh_menu
if !_el!==2  goto :oh_uninstall
if !_el!==1  goto :oh_menu2
goto :oh_menu
)

::========================================================================================================================================

:oh_menu2

cls
mode 128, 32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=32;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

title  Ohook Activation

::  Check files

for %%# in (sppc32.dll sppc64.dll) do (
if not exist "!_work!\BIN\%%#" (
%eline%
echo '%%#' file is missing in 'BIN' folder. Aborting...
goto dk_done
)
)

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

set error=

cls
echo:
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set osarch=%%b
for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if "%%j"=="" (set fullbuild=%%i) else (set fullbuild=%%i.%%j)
echo Checking OS Info                        [%winos% ^| %fullbuild% ^| %osarch%]

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

set "_serv=sppsvc Winmgmt"
set officeact=1
call :dk_errorcheck

::  Check unsupported office versions

set o14msi=
set o14c2r=
set o16uwp=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office
%nul% reg query %_68%\14.0\Common\InstallRoot /v Path  && set o14msi=Office 2010 MSI 
%nul% reg query %_86%\14.0\Common\InstallRoot /v Path  && set o14msi=Office 2010 MSI 
%nul% reg query %_68%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 
%nul% reg query %_86%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 

if %winbuild% GEQ 10240 %psc% "Get-AppxPackage -name "Microsoft.Office.Desktop"" | find /i "Office" %nul1% && set o16uwp=Office UWP 

if not "%o14msi%%o14c2r%%o16uwp%"=="" (
echo:
set error=1
call :dk_color %Red% "Checking Unsupported Office Install     [ %o14msi%%o14c2r%%o16uwp%]"
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

if %winbuild% GEQ 10240 %psc% "Get-AppxPackage -name "Microsoft.MicrosoftOfficeHub"" | find /i "Office" %nul1% && (
echo:
echo You have only Office dashboard app installed, you need to install full Office version.
)
echo:
call :dk_color %Blue% "Download and install Office from below URL and try again."
echo:
echo https://%mas%/genuine-installation-media.html
goto dk_done
)

set multioffice=
if not "%o16c2r%%o15c2r%%o16msi%%o15msi%"=="1" set multioffice=1
if not "%o14msi%%o14c2r%%o16uwp%"=="" set multioffice=1

if defined multioffice (
call :dk_color %Gray% "Checking Multiple Office Install        [Found. Its best to install only one version]"
)

::========================================================================================================================================

::  Process Office 15.0 C2R

if not defined o15c2r goto :starto16c2r

call :oh_reset
call :oh_actids

set oVer=15
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
if not defined _oArch for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v Platform" %nul6%') do (set "_oArch=%%b")

echo "%o15c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=10) || (set _tok=9)
for /f "tokens=%_tok% delims=\" %%a in ('reg query %o15c2r_reg%\ProductReleaseIDs\Active %nul6% ^| findstr /i "Retail Volume"') do (
if not defined _oIds (set "_oIds=%%a") else (set "_oIds=!_oIds! %%a")
)

set "_oLPath=%_oRoot%\Licenses"

if [%_oArch%]==[x64] (set "_hookPath=%_oRoot%\vfs\System"    & set "_hook=sppc64.dll")
if [%_oArch%]==[x86] (set "_hookPath=%_oRoot%\vfs\SystemX86" & set "_hook=sppc32.dll")
if not [%osarch%]==[x86] (
if [%_oArch%]==[x64] set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if [%_oArch%]==[x86] set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

echo:
echo Activating Office 15.0 %_oArch% C2R...

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :starto16c2r
)

call :oh_process
call :oh_hookinstall

::========================================================================================================================================

:starto16c2r

::  Process Office 16.0 C2R

if not defined o16c2r goto :startmsi

call :oh_reset
call :oh_actids

set oVer=16
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")

echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=9) || (set _tok=8)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs" /s /f ".16" /k %nul6% ^| findstr /i "Retail Volume"') do (
if not defined _oIds (set "_oIds=%%a") else (set "_oIds=!_oIds! %%a")
)
set _oIds=%_oIds:.16=%

set "_oLPath=%_oRoot%\Licenses16"

if [%_oArch%]==[x64] (set "_hookPath=%_oRoot%\vfs\System"    & set "_hook=sppc64.dll")
if [%_oArch%]==[x86] (set "_hookPath=%_oRoot%\vfs\SystemX86" & set "_hook=sppc32.dll")
if not [%osarch%]==[x86] (
if [%_oArch%]==[x64] set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if [%_oArch%]==[x86] set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

echo:
echo Activating Office 16.0 %_oArch% C2R...

if not defined _oIds (
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
set error=1
goto :startmsi
)

call :oh_process
call :oh_hookinstall

::========================================================================================================================================

::  Find remnants of Office vNext license block and remove it because it stops non vNext licenses from appearing

set sub_next=
set kNext=HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing

reg query %kNext%\LicensingNext /v MigrationToV5Done %nul2% | find /i "0x1" %nul% && (
reg query %kNext%\LicensingNext %nul2% | findstr /i "volume retail" %nul2% | findstr /i "0x2 0x3" %nul% && (
set sub_next=1
reg delete %kNext% /f %nul%
)
)

if defined sub_next (
reg query %kNext%\LicensingNext %nul% && (
call :dk_color %Red% "Removing Office vNext Block             [Failed]"
) || (
echo Removing Office vNext Block             [Successful]
)
)

::========================================================================================================================================

::  O365 products attempt to validate the license and may show a banner "There was a problem checking this device's license status.", other products don't do that.
::  A simple registry entry can skip this check

echo "%_lic%" | find /i "Subscription" %nul% && (
echo Adding a Reg To Skip License Check      [Successful]
reg add HKCU\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /v "TimeOfLastHeartbeatFailure" /t REG_SZ /d "2033-08-18T22:18:45Z" /f %nul%
)

::========================================================================================================================================

::  mass grave[.]dev/office-license-is-not-genuine.html
::  Add registry keys for volume products so that 'non-genuine' banner won't appear 
::  Script already is using MAK instead of GVLK so it won't appear anyway, but registry keys are added incase Office installs default GVLK grace key for volume products

echo "%_oIds%" | find /i "Volume" %nul1% && (
if %winbuild% GEQ 9200 (
if not [%osarch%]==[x86] (
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f /reg:32 %nul%
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" /reg:32 %nul%
)
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f %nul%
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" %nul%
echo Adding a Reg To Prevent Banner          [Successful]
)
)

::========================================================================================================================================

:startmsi

if defined o15msi call :oh_processmsi 15 %o15msi_reg%
if defined o16msi call :oh_processmsi 16 %o16msi_reg%

::========================================================================================================================================

::  Uninstall other / grace Keys

set upk_result=0
set allapplist=

if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='0ff1ce15-a989-479d-af46-f275c6370663' and PartialProductKey is not null) get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663'' AND PartialProductKey IS NOT NULL').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined allapplist (call set "allapplist=!allapplist! %%a") else (call set "allapplist=%%a"))

for %%# in (%allapplist%) do (
echo "%_allactid%" | find /i "%%#" %nul1% || (
cscript //nologo %windir%\system32\slmgr.vbs /upk %%# %nul% && (
set upk_result=1
) || (
set error=1
set upk_result=2
)
)
)

if not %upk_result%==0 echo:
if %upk_result%==1 echo Uninstalling Other/Grace Keys           [Successful]
if %upk_result%==2 call :dk_color %Red% "Uninstalling Other/Grace Keys           [Failed]"

::========================================================================================================================================

::  Refresh Windows Insider Preview Licenses
::  It required in Insider versions otherwise office may not activate

if exist "%windir%\system32\spp\store_test\2.0\tokens.dat" (
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if !errorlevel! NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
)

::========================================================================================================================================

echo:
if not defined error (
call :dk_color %Green% "Office is permanently activated."
echo Help: https://%mas%/troubleshoot
) else (
call :dk_color %Red% "Some errors were detected."
if not defined ierror if not defined showfix if not defined serv_cor if not defined serv_cste call :dk_color %Blue% "%_fixmsg%"
echo:
call :dk_color2 %Blue% "Check this page for help" %_Yellow% " https://%mas%/troubleshoot"
)

goto :dk_done

::========================================================================================================================================

:oh_uninstall

cls
mode 99, 28
title  Uninstall Ohook Activation

set _present=
set _unerror=
call :oh_reset
call :oh_getpath

echo:
echo Uninstalling Ohook Activation...
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

reg query HKCU\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /s %nul2% | find /i "2033" %nul% && (
echo:
echo Deleting - HKCU\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency
reg delete HKCU\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /f
)

reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" %nul% && (
echo:
echo Deleting - Registry key to prevent non-genuine banner
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f
)

reg query "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" %nul% && (
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f
)

echo __________________________________________________________________________________________
echo:

if not defined _present (
echo Ohook Activation is not installed.
) else (
if defined _unerror (
call :dk_color %Red% "Failed to uninstall Ohook activation."
call :dk_color %Blue% "Close Office apps if they are running and try again."
) else (
call :dk_color %Green% "Successfully uninstalled Ohook activation."
)
)
echo __________________________________________________________________________________________

goto :dk_done

::========================================================================================================================================

:oh_reset

set _oRoot=
set _oArch=
set _oIds=
set _oLPath=
set _hookPath=
set _hook=
set _sppcPath=
set _key=
set _actid=
set _prod=
set _lic=
set _License=
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

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\OSPP.VBS" (set o16msi=1&set o16msi_reg=%_86%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\OSPP.VBS" (set o16msi=1&set o16msi_reg=%_68%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\OSPP.VBS" (set o15msi=1&set o15msi_reg=%_86%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\OSPP.VBS" (set o15msi=1&set o15msi_reg=%_68%\15.0)

exit /b

::========================================================================================================================================

:oh_installkey

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%_key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%_key%')" %nul%
if not %errorlevel%==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %_key% %nul%
set errorcode=%errorlevel%
cmd /c exit /b %errorcode%
if %errorcode% NEQ 0 set "errorcode=[0x%=ExitCode%]"

if %errorcode% EQU 0 (
call :dk_refresh
echo Installing Generic Product Key          [%_key%] [%_prod%] [%_lic%] [Successful]
) else (
call :dk_color %Red% "Installing Generic Product Key          [%_key%] [%_prod%] [%_lic%] [Failed] %errorcode%"
if not defined error (
call :dk_color %Blue% "%_fixmsg%"
set showfix=1
)
set error=1
)

exit /b

::========================================================================================================================================

:oh_installlic

if not defined _oLPath exit /b

set _License=%_prod:Retail=R_%
set _License=%_License:Volume=VL_%

for %%# in ("!_oLPath!\client-issuance-*.xrm-ms") do (
cscript //nologo %windir%\system32\slmgr.vbs /ilc "!_oLPath!\%%~nx#" %nul%
)
cscript //nologo %windir%\system32\slmgr.vbs /ilc "!_oLPath!\pkeyconfig-office.xrm-ms" %nul%

for %%# in ("!_oLPath!\%_License%*.xrm-ms") do (
cscript //nologo %windir%\system32\slmgr.vbs /ilc "!_oLPath!\%%~nx#" %nul%
)

call :oh_actids

echo "!oapplist!" | find /i "!_actid!" %nul1% && (
call :dk_color %Gray% "Installing Missing License Files        [Office %oVer%.0 %_prod%] [Successful]"
) || (
set error=1
call :dk_color %Red% "Installing Missing License Files        [Office %oVer%.0 %_prod%] [Failed]"
)

exit /b

::========================================================================================================================================

:oh_hookinstall

set ierror=
set hasherror=

del /s /q "%_hookPath%\sppcs.dll" %nul%
del /s /q "%_hookPath%\sppc.dll" %nul%

if exist "%_hookPath%\sppcs.dll" set ierror=1
if exist "%_hookPath%\sppc.dll" set ierror=1

mklink "%_hookPath%\sppcs.dll" "%_sppcPath%" %nul%
if not %errorlevel%==0 set ierror=1

pushd "!_work!\BIN\"
copy /y %_hook% "%_hookPath%\sppc.dll" %nul%
popd
if not exist "%_hookPath%\sppc.dll" set ierror=1

echo:
if not defined ierror (
echo Symlinking System's sppc.dll To         ["%_hookPath%\sppcs.dll"] [Successful]
echo Copying Custom %_hook% To            ["%_hookPath%\sppc.dll"] [Successful]
) else (
set error=1
call :dk_color %Red% "Symlinking Systems sppc.dll             [Failed]"
call :dk_color %Red% "Copying Custom %_hook%               [Failed]"
echo ["%_hookPath%\sppc.dll"]
call :dk_color %Blue% "Close Office apps if they are running and try again."
)

if not defined ierror call :oh_modify "%_hookPath%\sppc.dll"

if not defined ierror (
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
set _key=
set _actid=
set _lic=
set _prod=%%#

call :ohookdata getinfo %%#

if not [!_key!]==[] (
echo "!oapplist!" | find /i "!_actid!" %nul1% || call :oh_installlic
call :oh_installkey
) else (
set error=1
call :dk_color %Red% "Checking Product In Script              [Office %oVer%.0 %%# not found in script]"
echo Make sure you are using Latest MAS script.
)
)

exit /b

::========================================================================================================================================

:oh_msiproducts

set msitemp=%SystemRoot%\Temp\_msitemp.txt

if %oVer%==15 set _psmsikey=%o15msi_reg:HKLM\=HKLM:%
if %oVer%==16 set _psmsikey=%o16msi_reg:HKLM\=HKLM:%

if exist %msitemp% del /f /q %msitemp%
%psc% "$Key = '%_psmsikey%\Registration\{*FF1CE}'; $keydata = Get-ItemProperty -Path $Key -Name "DigitalProductID"; $binaryData = $keydata."DigitalProductID"; $stringData = [System.Text.Encoding]::Unicode.GetString($binaryData);$stringData" >>%msitemp%

if exist %msitemp% call :ohookdata getmsiprod
if exist %msitemp% del /f /q %msitemp%

exit /b

::========================================================================================================================================

:oh_processmsi

::  Process Office MSI Version

call :oh_reset
call :oh_actids

set oVer=%1
for /f "skip=2 tokens=2*" %%a in ('"reg query %2\Common\InstallRoot /v Path" %nul6%') do (set "_oRoot=%%b")
if "%_oRoot:~-1%"=="\" set "_oRoot=%_oRoot:~0,-1%"

echo "%2" | find /i "Wow6432Node" %nul1% && set _oArch=x86
if not [%osarch%]==[x86] if not defined _oArch set _oArch=x64
if [%osarch%]==[x86] set _oArch=x86

if [%_oArch%]==[x64] (set "_hookPath=%_oRoot%" & set "_hook=sppc64.dll")
if [%_oArch%]==[x86] (set "_hookPath=%_oRoot%" & set "_hook=sppc32.dll")
if not [%osarch%]==[x86] (
if [%_oArch%]==[x64] set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if [%_oArch%]==[x86] set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

call :oh_msiproducts

echo:
echo Activating Office %1.0 %_oArch% MSI...

if not defined _oIds (
set error=1
call :dk_color %Red% "Checking Installed Products             [Product IDs not found. Aborting activation...]"
exit /b
)

call :oh_process
call :oh_hookinstall

exit /b

::========================================================================================================================================

::  Refresh license status

:dk_refresh

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call RefreshLicenseStatus %nul%
if %_wmic% EQU 0 %psc% "$null=(([WMICLASS]'SoftwareLicensingService').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::  Get Windows Activation IDs

:dk_actids

set applist=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f') get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f''').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))
exit /b

::  Get Office Activation IDs

:oh_actids

set oapplist=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='0ff1ce15-a989-479d-af46-f275c6370663') get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663''').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined oapplist (call set "oapplist=!oapplist! %%a") else (call set "oapplist=%%a"))
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

:dk_errorcheck

set showfix=

::  Check corrupt services

set serv_cor=
for %%# in (%_serv%) do (
set _corrupt=
sc start %%# %nul%
if !errorlevel! EQU 1060 set _corrupt=1
sc query %%# %nul% || set _corrupt=1
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type) do if not defined _corrupt (reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v %%G %nul% || set _corrupt=1)
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
sc query %%# | find /i "RUNNING" %nul% || set checkerror=1

sc start %%# %nul%
if !errorlevel! NEQ 1056 if !errorlevel! NEQ 0 (set errorcode=!errorlevel!&set checkerror=1)
if defined checkerror if defined serv_e (set "serv_e=!serv_e!, %%#-!errorcode!") else (set "serv_e=%%#-!errorcode!")
)

if defined serv_e (
set error=1
call :dk_color %Red% "Starting Services                       [Failed] [%serv_e%]"
echo %serv_e% | findstr /i "ClipSVC-1058 sppsvc-1058" %nul% && (
call :dk_color %Blue% "Restart the system to fix disabled service error 1058."
set showfix=1
)
)

::========================================================================================================================================

::  Various error checks

if defined safeboot_option (
set error=1
set showfix=1
call :dk_color2 %Red% "Checking Boot Mode                      " %Blue% "[System is running in safe mode. Run in normal mode.]"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State" %nul2% | find /i "IMAGE_STATE_COMPLETE" %nul1% || (
set error=1
set showfix=1
call :dk_color2 %Red% "Checking Audit Mode                     " %Blue% "[IMAGE_STATE_COMPLETE status not found. Run in normal mode.]"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinPE" /v InstRoot %nul% && (
set error=1
set showfix=1
call :dk_color2 %Red% "Checking WinPE                          " %Blue% "[System is running in WinPE mode. Run in normal mode.]"
)


%psc% $ExecutionContext.SessionState.LanguageMode %nul2% | find /i "Full" %nul1% || (
set error=1
call :dk_color %Red% "Checking Powershell                     [Not Responding]"
)


DISM /English /Online /Get-CurrentEdition %nul%
set dism_error=%errorlevel%
cmd /c exit /b %dism_error%
if %dism_error% NEQ 0 set "dism_error=[0x%=ExitCode%]"
if %dism_error% NEQ 0 (
call :dk_color %Red% "Checking DISM                           [Not Responding] %dism_error%"
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
net stop sppsvc /y %nul%
net start sppsvc /y %nul%
set error=1
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Plugins\Objects\msft:rm/algorithm/hwid/4.0" /f ba02fed39662 /d %nul% || (
call :dk_color %Red% "Checking SPP Registry Key               [Incorrect ModuleId Found]"
call :dk_color %Blue% "Possibly Caused By Gaming Spoofers. Help: https://%mas%/troubleshoot"
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


if exist "%tokenstore%\" if not exist "%tokenstore%\tokens.dat" (
set error=1
call :dk_color %Red% "Checking SPP tokens.dat                 [Not Found] [%tokenstore%\]"
)


if not exist %SystemRoot%\system32\sppsvc.exe (
set error=1
set showfix=1
call :dk_color %Red% "Checking sppsvc.exe File                [Not Found]"
)


::  Below checks are performed if required services are not disabled or corrupted + if there is any error + slmgr /dlv errorlevel is not Zero + no fix was shown before

set wpaerror=
set permerror=
if not defined serv_cor if not defined serv_cste if defined error if /i not %error_code%==0 if not defined showfix (

REM  This code checks for invalid registry keys in HKLM\SYSTEM\WPA. This issue may appear even on healthy systems.

if %winbuild% GEQ 14393 (
set /a count=0
for /f %%a in ('reg query "HKLM\SYSTEM\WPA" %nul6%') do set /a count+=1
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
call :dk_color %Blue% "%_fixmsg%"
if not defined permerror call :dk_color %Blue% "If activation still fails then run Fix WPA Registry option."
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
pause %nul1%
exit /b

::========================================================================================================================================

::  1st column = Office version number
::  2nd column = Activation ID
::  3rd column = Generic key. Preference is given in this order, Retail > OEM:NONSLP > Volume:MAK > Volume:GVLK > Retail:TB:Sub
::  4th column = Last part of license description
::  5th column = Edition
::  Separator  = "_"

:ohookdata

set f=
for %%# in (
15_ab4d047b-97cf-4126-a69f-34df08e2f254_B7R%f%FY-7N%f%XPK-Q43%f%42-Y9%f%X2H-3JX%f%4X_Retail________AccessRetail
15_4374022d-56b8-48c1-9bb7-d8f2fc726343_9MF%f%9G-CN%f%32B-HV7%f%XT-9X%f%J8T-9KV%f%F4_MAK___________AccessVolume
15_1b1d9bd5-12ea-4063-964c-16e7e87d6e08_NT8%f%89-MB%f%H4X-8MD%f%4H-X8%f%R2D-WQH%f%F8_Retail________ExcelRetail
15_ac1ae7fd-b949-4e04-a330-849bc40638cf_Y3N%f%36-YC%f%HDK-XYW%f%BG-KY%f%QVV-BDT%f%J2_MAK___________ExcelVolume
15_cfaf5356-49e3-48a8-ab3c-e729ab791250_BMK%f%4W-6N%f%88B-BP9%f%QR-PH%f%FCK-MG7%f%GF_Retail________GrooveRetail
15_4825ac28-ce41-45a7-9e6e-1fed74057601_RN8%f%4D-7H%f%CWY-FTC%f%BK-JM%f%XWM-HT7%f%GJ_MAK___________GrooveVolume
15_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_2WQ%f%NF-GB%f%K4B-XVG%f%6F-BB%f%MX7-M4F%f%2Y_OEM-Perp______HomeBusinessPipcRetail
15_cd256150-a898-441f-aac0-9f8f33390e45_NVT%f%DK-QB%f%8J9-M28%f%GR-92%f%BPC-BR9%f%6Q_Retail________HomeBusinessRetail
15_98685d21-78bd-4c62-bc4f-653344a63035_R29%f%WQ-CM%f%NCM-2TH%f%CT-K4%f%H92-FX6%f%KR_Retail________HomeStudentRetail
15_44984381-406e-4a35-b1c3-e54f499556e2_RV7%f%NQ-HY%f%3WW-7CK%f%WH-QT%f%VMW-29V%f%HC_Retail________InfoPathRetail
15_9e016989-4007-42a6-8051-64eb97110cf2_C4T%f%GN-QQ%f%W6Y-FYK%f%XC-6W%f%JW7-X73%f%VG_MAK___________InfoPathVolume
15_9103f3ce-1084-447a-827e-d6097f68c895_6MD%f%N4-WF%f%3FV-4WH%f%3Q-W6%f%99V-RGC%f%MY_PrepidBypass__LyncAcademicRetail
15_ff693bf4-0276-4ddb-bb42-74ef1a0c9f4d_N42%f%BF-CB%f%Y9F-W2C%f%7R-X3%f%97X-DYF%f%QW_PrepidBypass__LyncEntryRetail
15_fada6658-bfc6-4c4e-825a-59a89822cda8_89P%f%23-2N%f%K2R-JXM%f%2M-3Q%f%8R8-BWM%f%3Y_Retail________LyncRetail
15_e1264e10-afaf-4439-a98b-256df8bb156f_3WK%f%CD-RN%f%489-4M7%f%XJ-GJ%f%2GQ-YBF%f%Q6_MAK___________LyncVolume
15_3169c8df-f659-4f95-9cc6-3115e6596e83_YYH%f%XG-KT%f%NKF-XGR%f%T7-H2%f%99D-8K6%f%MM_Retail________MondoRetail
15_f33485a0-310b-4b72-9a0e-b1d605510dbd_2YN%f%YQ-FQ%f%MVG-CB8%f%KW-6X%f%KYD-M7R%f%RJ_MAK___________MondoVolume
15_3391e125-f6e4-4b1e-899c-a25e6092d40d_4TG%f%WV-6N%f%9P6-G2H%f%8Y-2H%f%WKB-B4F%f%F4_Bypass________OneNoteFreeRetail
15_8b524bcc-67ea-4876-a509-45e46f6347e8_3KX%f%XQ-PV%f%N2C-8P7%f%YY-HC%f%V88-GVG%f%Q6_Retail________OneNoteRetail
15_b067e965-7521-455b-b9f7-c740204578a2_JDM%f%WF-NJ%f%C7B-HRC%f%HY-WF%f%T8G-BPX%f%D9_MAK___________OneNoteVolume
15_12004b48-e6c8-4ffa-ad5a-ac8d4467765a_9N4%f%RQ-CF%f%8R2-HBV%f%CB-J3%f%C9V-94P%f%4D_Retail________OutlookRetail
15_8d577c50-ae5e-47fd-a240-24986f73d503_HNG%f%29-GG%f%WRG-RFC%f%8C-JT%f%FP4-2J9%f%FH_MAK___________OutlookVolume
15_5aab8561-1686-43f7-9ff5-2c861da58d17_9CY%f%B3-NF%f%MRW-YFD%f%G6-XC%f%7TF-BY3%f%6J_OEM-Perp______PersonalPipcRetail
15_17e9df2d-ed91-4382-904b-4fed6a12caf0_2NC%f%QJ-MF%f%RMH-TXV%f%83-J7%f%V4C-RVR%f%WC_Retail________PersonalRetail
15_31743b82-bfbc-44b6-aa12-85d42e644d5b_HVM%f%N2-KP%f%HQH-DVQ%f%MK-7B%f%3CM-FGB%f%FC_Retail________PowerPointRetail
15_e40dcb44-1d5c-4085-8e8f-943f33c4f004_47D%f%KN-HP%f%JP7-RF9%f%M3-VC%f%YT2-TMQ%f%4G_MAK___________PowerPointVolume
15_064383fa-1538-491c-859b-0ecab169a0ab_N3Q%f%MM-GK%f%DT3-JQG%f%X6-7X%f%3MQ-4GB%f%G3_Retail________ProPlusRetail
15_2b88c4f2-ea8f-43cd-805e-4d41346e18a7_QKH%f%NX-M9%f%GGH-T3Q%f%MW-YP%f%K4Q-QRP%f%9V_MAK___________ProPlusVolume
15_4e26cac1-e15a-4467-9069-cb47b67fe191_CF9%f%DD-6C%f%NW2-BJW%f%JQ-CV%f%CFX-Y7T%f%XD_OEM-Perp______ProfessionalPipcRetail
15_44bc70e2-fb83-4b09-9082-e5557e0c2ede_MBQ%f%BN-CQ%f%PT6-PXR%f%MC-TY%f%JFR-3C8%f%MY_Retail________ProfessionalRetail
15_f2435de4-5fc0-4e5b-ac97-34f515ec5ee7_B4R%f%N4-H4%f%JYG-YMQ%f%MC-CY%f%GKH-46W%f%PH_Retail________ProjectProRetail
15_ed34dc89-1c27-4ecd-8b2f-63d0f4cedc32_WFC%f%T2-NB%f%FQ7-JD7%f%VV-MF%f%JX6-6F2%f%CM_MAK___________ProjectProVolume
15_5517e6a2-739b-4822-946f-7f0f1c5934b1_2B9%f%6V-X9%f%NJY-WFB%f%RC-Q8%f%MP2-7CC%f%C4_Retail________ProjectStdRetail
15_2b9e4a37-6230-4b42-bee2-e25ce86c8c7a_3CN%f%QX-T3%f%4TY-99R%f%H4-C4%f%YD2-KWY%f%GV_MAK___________ProjectStdVolume
15_c3a0814a-70a4-471f-af37-2313a6331111_TWN%f%CJ-YR%f%84W-X7P%f%PF-6D%f%PRP-D67%f%VC_Retail________PublisherRetail
15_38ea49f6-ad1d-43f1-9888-99a35d7c9409_DJP%f%HV-NC%f%JV6-GWP%f%T6-K2%f%6JX-C7G%f%X6_MAK___________PublisherVolume
15_ba3e3833-6a7e-445a-89d0-7802a9a68588_3NY%f%6J-WH%f%T3F-47B%f%DV-JH%f%F36-234%f%3W_PrepidBypass__SPDRetail
15_32255c0a-16b4-4ce2-b388-8a4267e219eb_V6V%f%WN-KC%f%2HR-YYD%f%D6-9V%f%7HQ-7T7%f%VP_Retail________StandardRetail
15_a24cca51-3d54-4c41-8a76-4031f5338cb2_9TN%f%6B-PC%f%YH4-MCV%f%DQ-KT%f%83C-TMQ%f%7T_MAK___________StandardVolume
15_15d12ad4-622d-4257-976c-5eb3282fb93d_JRJ%f%NJ-33%f%M7C-R73%f%X3-P9%f%XF7-R89%f%MY_Retail________VisioProRetail
15_3e4294dd-a765-49bc-8dbd-cf8b62a4bd3d_YN7%f%CF-XR%f%H6R-CGK%f%RY-GK%f%PV3-BG7%f%WF_MAK___________VisioProVolume
15_dae597ce-5823-4c77-9580-7268b93a4b23_G2N%f%2Q-4Q%f%D2T-DY4%f%TP-GY%f%489-3RQ%f%CG_Retail________VisioStdRetail
15_44a1f6ff-0876-4edb-9169-dbb43101ee89_RX6%f%3Y-4N%f%FK2-XTY%f%C8-C6%f%B3W-YPX%f%PJ_MAK___________VisioStdVolume
15_191509f2-6977-456f-ab30-cf0492b1e93a_NB7%f%7V-RP%f%FQ6-PMM%f%KQ-T8%f%7DV-M4D%f%84_Retail________WordRetail
15_9cedef15-be37-4ff0-a08a-13a045540641_RPH%f%PB-Y7%f%NC4-3VY%f%FM-DW%f%7VD-G8Y%f%J8_MAK___________WordVolume
15_befee371-a2f5-4648-85db-a2c55fdf324c_JXR%f%8H-NJ%f%3MK-X66%f%W8-78%f%CWD-QRQ%f%6C_Retail________O365BusinessRetail
15_537ea5b5-7d50-4876-bd38-a53a77caca32_J2W%f%28-TN%f%9C8-26P%f%WV-F7%f%J4G-72X%f%CB_Subscription1_O365HomePremRetail
15_0c4e5e7a-b436-4776-bb89-88e4b14687e2_T6J%f%JW-DY%f%N99-WK4%f%6B-9M%f%KKV-94P%f%2Y_Retail________O365ProPlusRetail
15_7a75647f-636f-4607-8e54-e1b7d1ad8930_B9V%f%9Q-F6%f%ND8-T4K%f%4K-88%f%68M-HT7%f%KT_Retail________O365SmallBusPremRetail
16_bfa358b0-98f1-4125-842e-585fa13032e6_WHK%f%4N-YQ%f%GHB-XWX%f%CC-G3%f%HYC-6JF%f%94_Retail________AccessRetail
16_9d9faf9e-d345-4b49-afce-68cb0a539c7c_RNB%f%7V-P4%f%8F4-3FY%f%Y6-2P%f%3R3-63B%f%QV_PrepidBypass__AccessRuntimeRetail
16_3b2fa33f-cd5a-43a5-bd95-f49f3f546b0b_JJ2%f%Y4-N8%f%KM3-Y8K%f%Y3-Y2%f%2FR-R3K%f%VK_MAK___________AccessVolume
16_424d52ff-7ad2-4bc7-8ac6-748d767b455d_RKJ%f%BN-VW%f%TM2-BDK%f%XX-RK%f%QFD-JTY%f%Q2_Retail________ExcelRetail
16_685062a7-6024-42e7-8c5f-6bb9e63e697f_FVG%f%NR-X8%f%2B2-6PR%f%JM-YT%f%4W7-8HV%f%36_MAK___________ExcelVolume
16_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_2WQ%f%NF-GB%f%K4B-XVG%f%6F-BB%f%MX7-M4F%f%2Y_OEM-Perp______HomeBusinessPipcRetail
16_86834d00-7896-4a38-8fae-32f20b86fa2b_HM6%f%FM-NV%f%F78-KV9%f%PM-F3%f%6B8-D9M%f%XD_Retail________HomeBusinessRetail
16_c28acdb8-d8b3-4199-baa4-024d09e97c99_PNP%f%RV-F2%f%627-Q8J%f%VC-3D%f%GR9-WTY%f%RK_Retail________HomeStudentRetail
16_e2127526-b60c-43e0-bed1-3c9dc3d5a468_YWD%f%4R-CN%f%KVT-VG8%f%VJ-93%f%33B-RC3%f%B8_Retail________HomeStudentVNextRetail
16_b21367df-9545-4f02-9f24-240691da0e58_Y3X%f%PN-P7%f%PC4-MGP%f%TJ-2B%f%Y3Q-RJR%f%FV_Retail________MondoRetail
16_2cd0ea7e-749f-4288-a05e-567c573b2a6c_FMT%f%QQ-84%f%NR8-274%f%4R-MX%f%F4P-PGY%f%R3_MAK___________MondoVolume
16_436366de-5579-4f24-96db-3893e4400030_XYN%f%TG-R9%f%6FY-369%f%HX-YF%f%PHY-F9C%f%PM_Bypass________OneNoteFreeRetail
16_83ac4dd9-1b93-40ed-aa55-ede25bb6af38_FXF%f%6F-CN%f%C26-W64%f%3C-K6%f%KB7-6XX%f%W3_Retail________OneNoteRetail
16_23b672da-a456-4860-a8f3-e062a501d7e8_9TY%f%VN-D7%f%6HK-BVM%f%WT-Y7%f%G88-9TP%f%PV_MAK___________OneNoteVolume
16_5a670809-0983-4c2d-8aad-d3c2c5b7d5d1_7N4%f%KG-P2%f%QDH-86V%f%9C-DJ%f%FVF-369%f%W9_Retail________OutlookRetail
16_50059979-ac6f-4458-9e79-710bcb41721a_7QP%f%NR-3H%f%FDG-YP6%f%T9-JQ%f%CKQ-KKX%f%XC_MAK___________OutlookVolume
16_5aab8561-1686-43f7-9ff5-2c861da58d17_9CY%f%B3-NF%f%MRW-YFD%f%G6-XC%f%7TF-BY3%f%6J_OEM-Perp______PersonalPipcRetail
16_a9f645a1-0d6a-4978-926a-abcb363b72a6_FT7%f%VF-XB%f%N92-HPD%f%JV-RH%f%MBY-6VK%f%BF_Retail________PersonalRetail
16_f32d1284-0792-49da-9ac6-deb2bc9c80b6_N7G%f%CB-WQ%f%T7K-QRH%f%WG-TT%f%PYD-7T9%f%XF_Retail________PowerPointRetail
16_9b4060c9-a7f5-4a66-b732-faf248b7240f_X3R%f%T9-ND%f%G64-VMK%f%2M-KQ%f%6XY-DPF%f%GV_MAK___________PowerPointVolume
16_de52bd50-9564-4adc-8fcb-a345c17f84f9_GM4%f%3N-F7%f%42Q-6JD%f%DK-M6%f%22J-J8G%f%DV_Retail________ProPlusRetail
16_c47456e3-265d-47b6-8ca0-c30abbd0ca36_FNV%f%K8-8D%f%VCJ-F7X%f%3J-KG%f%VQB-RC2%f%QY_MAK___________ProPlusVolume
16_4e26cac1-e15a-4467-9069-cb47b67fe191_CF9%f%DD-6C%f%NW2-BJW%f%JQ-CV%f%CFX-Y7T%f%XD_OEM-Perp______ProfessionalPipcRetail
16_d64edc00-7453-4301-8428-197343fafb16_NXF%f%TK-YD%f%9Y7-X9M%f%MJ-9B%f%WM6-J2Q%f%VH_Retail________ProfessionalRetail
16_0f42f316-00b1-48c5-ada4-2f52b5720ad0_WKK%f%GF-TT%f%N8F-QC9%f%T2-6C%f%2DC-3YH%f%63_Retail________ProjectProRetail
16_82f502b5-b0b0-4349-bd2c-c560df85b248_PKC%f%3N-8F%f%99H-28M%f%VY-J4%f%RYY-CWG%f%DH_MAK___________ProjectProVolume
16_16728639-a9ab-4994-b6d8-f81051e69833_JBN%f%PH-YF%f%2F7-Q9Y%f%29-86%f%CTG-C9Y%f%GV_MAKC2R________ProjectProXVolume
16_e9f0b3fc-962f-4944-ad06-05c10b6bcd5e_4H6%f%NF-QD%f%PG8-83W%f%97-9K%f%8XY-XBT%f%9Y_Retail________ProjectStdRetail
16_82e6b314-2a62-4e51-9220-61358dd230e6_4TG%f%WV-6N%f%9P6-G2H%f%8Y-2H%f%WKB-B4G%f%93_MAK___________ProjectStdVolume
16_431058f0-c059-44c5-b9e7-ed2dd46b6789_N3W%f%2Q-69%f%MBT-27R%f%D9-BH%f%8V3-JT2%f%C8_MAKC2R________ProjectStdXVolume
16_6e0c1d99-c72e-4968-bcb7-ab79e03e201e_WKW%f%ND-X6%f%G9G-CDM%f%TV-CP%f%GYJ-6MV%f%BF_Retail________PublisherRetail
16_fcc1757b-5d5f-486a-87cf-c4d6dedb6032_9QV%f%N2-PX%f%XRX-8V4%f%W8-Q7%f%926-TJG%f%D8_MAK___________PublisherVolume
16_9103f3ce-1084-447a-827e-d6097f68c895_6MD%f%N4-WF%f%3FV-4WH%f%3Q-W6%f%99V-RGC%f%MY_PrepidBypass__SkypeServiceBypassRetail
16_971cd368-f2e1-49c1-aedd-330909ce18b6_4N4%f%D8-3J%f%7Y3-YYW%f%7C-73%f%HD2-V8R%f%HY_PrepidBypass__SkypeforBusinessEntryRetail
16_418d2b9f-b491-4d7f-84f1-49e27cc66597_PBJ%f%79-77%f%NY4-VRG%f%FG-Y8%f%WYC-CKC%f%RC_Retail________SkypeforBusinessRetail
16_03ca3b9a-0869-4749-8988-3cbc9d9f51bb_DMT%f%CJ-KN%f%RKR-JV8%f%TQ-V2%f%CR2-VFT%f%FH_MAK___________SkypeforBusinessVolume
16_4a31c291-3a12-4c64-b8ab-cd79212be45e_2FP%f%WN-4H%f%6CM-KD8%f%QQ-8H%f%CHC-P9X%f%YW_Retail________StandardRetail
16_0ed94aac-2234-4309-ba29-74bdbb887083_WHG%f%MQ-JN%f%MGT-MDQ%f%VF-WD%f%R69-KQB%f%WC_MAK___________StandardVolume
16_2dfe2075-2d04-4e43-816a-eb60bbb77574_7TM%f%6H-KN%f%R4K-Q7H%f%BX-72%f%T2Q-H22%f%RV_Retail________VisioProRetail
16_295b2c03-4b1c-4221-b292-1411f468bd02_NRK%f%T9-C8%f%GP2-XDY%f%XQ-YW%f%72K-MG9%f%2B_MAK___________VisioProVolume
16_0594dc12-8444-4912-936a-747ca742dbdb_G98%f%Q2-B6%f%N77-CFH%f%9J-K8%f%24G-XQC%f%C4_MAKC2R________VisioProXVolume
16_c76dbcbc-d71b-4f45-b5b3-b7494cb4e23e_Q8R%f%4N-GJ%f%CPG-CF3%f%2R-JF%f%CYR-K4T%f%F3_Retail________VisioStdRetail
16_44151c2d-c398-471f-946f-7660542e3369_XNC%f%JB-YY%f%883-JRW%f%64-DP%f%XMX-JXC%f%R6_MAK___________VisioStdVolume
16_1d1c6879-39a3-47a5-9a6d-aceefa6a289d_B2H%f%TN-JP%f%H8C-J6Y%f%6V-HC%f%HKB-43M%f%GT_MAKC2R________VisioStdXVolume
16_cacaa1bf-da53-4c3b-9700-11738ef1c2a5_P8K%f%82-NQ%f%7GG-JKY%f%8T-6V%f%HVY-88G%f%GD_Retail________WordRetail
16_c3000759-551f-4f4a-bcac-a4b42cbf1de2_YHM%f%WC-YN%f%6V9-WJP%f%XD-3W%f%QKP-TMV%f%CV_MAK___________WordVolume
16_518687bd-dc55-45b9-8fa6-f918e1082e83_WRY%f%J6-G3%f%NP7-7VH%f%94-8X%f%7KP-JB7%f%HC_Retail________Access2019Retail
16_385b91d6-9c2c-4a2e-86b5-f44d44a48c5f_6FW%f%HX-NK%f%YXK-BW3%f%4Q-7X%f%C9F-Q9P%f%X7_MAK-AE________Access2019Volume
16_22e6b96c-1011-4cd5-8b35-3c8fb6366b86_FGQ%f%NJ-JW%f%JCG-7Q8%f%MG-RM%f%RGJ-9TQ%f%VF_PrepidBypass__AccessRuntime2019Retail
16_c201c2b7-02a1-41a8-b496-37c72910cd4a_KBP%f%NW-64%f%CMM-8KW%f%CB-23%f%F44-8B7%f%HM_Retail________Excel2019Retail
16_05cb4e1d-cc81-45d5-a769-f34b09b9b391_8NT%f%4X-GQ%f%MCK-62X%f%4P-TW%f%6QP-YKP%f%YF_MAK-AE________Excel2019Volume
16_7fe09eef-5eed-4733-9a60-d7019df11cac_QBN%f%2Y-9B%f%284-9KW%f%78-K4%f%8PB-R62%f%YT_Retail________HomeBusiness2019Retail
16_4539aa2c-5c31-4d47-9139-543a868e5741_XNW%f%PM-32%f%XQC-Y7Q%f%JC-QG%f%GBV-YY7%f%JK_Retail________HomeStudent2019Retail
16_20e359d5-927f-47c0-8a27-38adbdd27124_WR4%f%3D-NM%f%WQQ-HCQ%f%R2-VK%f%XDR-37B%f%7H_Retail________Outlook2019Retail
16_92a99ed8-2923-4cb7-a4c5-31da6b0b8cf3_RN3%f%QB-GT%f%6D7-YB3%f%VH-F3%f%RPB-3GQ%f%YB_MAK-AE________Outlook2019Volume
16_2747b731-0f1f-413e-a92d-386ec1277dd8_NMB%f%Y8-V3%f%CV7-BX6%f%K6-29%f%22Y-43M%f%7T_Retail________Personal2019Retail
16_7e63cc20-ba37-42a1-822d-d5f29f33a108_HN2%f%7K-JH%f%J8R-7T7%f%KK-WJ%f%YC3-FM7%f%MM_Retail________PowerPoint2019Retail
16_13c2d7bf-f10d-42eb-9e93-abf846785434_29G%f%NM-VM%f%33V-WR2%f%3K-HG%f%2DT-KTQ%f%YR_MAK-AE________PowerPoint2019Volume
16_a3072b8f-adcc-4e75-8d62-fdeb9bdfae57_BN4%f%XJ-R9%f%DYY-96W%f%48-YK%f%8DM-MY7%f%PY_Retail________ProPlus2019Retail
16_6755c7a7-4dfe-46f5-bce8-427be8e9dc62_T8Y%f%BN-4Y%f%V3X-KK2%f%4Q-QX%f%BD7-T3C%f%63_MAK-AE________ProPlus2019Volume
16_1717c1e0-47d3-4899-a6d3-1022db7415e0_9NX%f%DK-MR%f%Y98-2VJ%f%V8-GF%f%73J-TQ9%f%FK_Retail________Professional2019Retail
16_0d270ef7-5aaf-4370-a372-bc806b96adb7_JDT%f%NC-PP%f%77T-T9H%f%2W-G4%f%J2J-VH8%f%JK_Retail________ProjectPro2019Retail
16_d4ebadd6-401b-40d5-adf4-a5d4accd72d1_TBX%f%BD-FN%f%WKJ-WRH%f%BD-KB%f%PHH-XD9%f%F2_MAK-AE________ProjectPro2019Volume
16_bb7ffe5f-daf9-4b79-b107-453e1c8427b5_R3J%f%NT-8P%f%BDP-MTW%f%CK-VD%f%2V8-HMK%f%F9_Retail________ProjectStd2019Retail
16_fdaa3c03-dc27-4a8d-8cbf-c3d843a28ddc_RBR%f%FX-MQ%f%NDJ-4XF%f%HF-7Q%f%VDR-JHX%f%GC_MAK-AE________ProjectStd2019Volume
16_f053a7c7-f342-4ab8-9526-a1d6e5105823_4QC%f%36-NW%f%3YH-D2Y%f%9D-RJ%f%PC7-VVB%f%9D_Retail________Publisher2019Retail
16_40055495-be00-444e-99cc-07446729b53e_K8F%f%2D-NB%f%M32-BF2%f%6V-YC%f%KFJ-29Y%f%9W_MAK-AE________Publisher2019Volume
16_b639e55c-8f3e-47fe-9761-26c6a786ad6b_JBD%f%KF-6N%f%CD6-49K%f%3G-2T%f%V79-BKP%f%73_Retail________SkypeforBusiness2019Retail
16_15a430d4-5e3f-4e6d-8a0a-14bf3caee4c7_9MN%f%Q7-YP%f%Q3B-6WJ%f%XM-G8%f%3T3-CBB%f%DK_MAK-AE________SkypeforBusiness2019Volume
16_f88cfdec-94ce-4463-a969-037be92bc0e7_N97%f%22-BV%f%9H6-WTJ%f%TT-FP%f%B93-978%f%MK_PrepidBypass__SkypeforBusinessEntry2019Retail
16_fdfa34dd-a472-4b85-bee6-cf07bf0aaa1c_NDG%f%VM-MD%f%27H-2XH%f%VC-KD%f%DX2-YKP%f%74_Retail________Standard2019Retail
16_beb5065c-1872-409e-94e2-403bcfb6a878_NT3%f%V6-XM%f%BK7-Q66%f%MF-VM%f%KR4-FC3%f%3M_MAK-AE________Standard2019Volume
16_a6f69d68-5590-4e02-80b9-e7233dff204e_2NW%f%VW-QG%f%F4T-9CP%f%MB-WY%f%DQ9-7XP%f%79_Retail________VisioPro2019Retail
16_f41abf81-f409-4b0d-889d-92b3e3d7d005_33Y%f%F4-GN%f%CQ3-J6G%f%DM-J6%f%7P3-FM7%f%QP_MAK-AE________VisioPro2019Volume
16_4a582021-18c2-489f-9b3d-5186de48f1cd_263%f%WK-3N%f%797-7R4%f%37-28%f%BKG-3V8%f%M8_Retail________VisioStd2019Retail
16_933ed0e3-747d-48b0-9c2c-7ceb4c7e473d_BGN%f%HX-QT%f%PRJ-F9C%f%9G-R8%f%QQG-8T2%f%7F_MAK-AE________VisioStd2019Volume
16_72cee1c2-3376-4377-9f25-4024b6baadf8_JXR%f%8H-NJ%f%3MK-X66%f%W8-78%f%CWD-QRV%f%R2_Retail________Word2019Retail
16_fe5fe9d5-3b06-4015-aa35-b146f85c4709_9F3%f%6R-PN%f%VHH-3DX%f%GQ-7C%f%D2H-R9D%f%3V_MAK-AE________Word2019Volume
16_f634398e-af69-48c9-b256-477bea3078b5_P28%f%6B-N3%f%XYP-36Q%f%RQ-29%f%CMP-RVX%f%9M_Retail________Access2021Retail
16_ae17db74-16b0-430b-912f-4fe456e271db_JBH%f%3N-P9%f%7FP-FRT%f%JD-MG%f%K2C-VFW%f%G6_MAK-AE________Access2021Volume
16_fb099c19-d48b-4a2f-a160-4383011060aa_V6Q%f%FB-7N%f%7G9-PF7%f%W9-M8%f%FQM-MY8%f%G9_Retail________Excel2021Retail
16_9da1ecdb-3a62-4273-a234-bf6d43dc0778_WNY%f%R4-KM%f%R9H-KVC%f%8W-7H%f%J8B-K79%f%DQ_MAK-AE________Excel2021Volume
16_0e878942-3ba1-48ee-a215-f64d10a9f97b_QV2%f%VN-37%f%WPX-DHH%f%4K-B8%f%2H6-PDP%f%3T_Retail1_______HomeBusiness2021Retail
16_9e7d5312-95fc-4443-a551-69c20624ec6b_3QN%f%G6-8F%f%JTH-88B%f%F9-KB%f%BVM-JMB%f%78_Retail________HomeStudent2021Retail
16_279706f4-3a4b-4877-949b-f8c299cf0cc5_NB2%f%TQ-3Y%f%79C-77C%f%6M-QM%f%Y7H-7QY%f%8P_Retail________OneNote2021Retail
16_ecea2cfa-d406-4a7f-be0d-c6163250d126_4NC%f%WR-9V%f%92Y-34V%f%B2-RP%f%THR-YTG%f%R7_Retail________Outlook2021Retail
16_45bf67f9-0fc8-4335-8b09-9226cef8a576_JQ9%f%MJ-QY%f%N6B-67P%f%X9-GY%f%FVY-QJ6%f%TB_MAK-AE________Outlook2021Volume
16_8f89391e-eedb-429d-af90-9d36fbf94de6_RRR%f%YB-DN%f%749-GCP%f%W4-9H%f%6VK-HCH%f%PT_Retail________Personal2021Retail
16_c9bf5e86-f5e3-4ac6-8d52-e114a604d7bf_3KX%f%XQ-PV%f%N2C-8P7%f%YY-HC%f%V88-GVM%f%96_Retail1_______PowerPoint2021Retail
16_716f2434-41b6-4969-ab73-e61e593a3875_39G%f%2N-3B%f%D9C-C4X%f%CM-BD%f%4QG-FVY%f%DY_MAK-AE________PowerPoint2021Volume
16_c2f04adf-a5de-45c5-99a5-f5fddbda74a8_8WX%f%TP-MN%f%628-KY4%f%4G-VJ%f%WCK-C7P%f%CF_Retail________ProPlus2021Retail
16_3f180b30-9b05-4fe2-aa8d-0c1c4790f811_RNH%f%JY-DT%f%FXW-HW9%f%F8-49%f%82D-MD2%f%CW_MAK-AE1_______ProPlus2021Volume
16_96097a68-b5c5-4b19-8600-2e8d6841a0db_JRJ%f%NJ-33%f%M7C-R73%f%X3-P9%f%XF7-R9F%f%6M_MAK-AE________ProPlusSPLA2021Volume
16_711e48a6-1a79-4b00-af10-73f4ca3aaac4_DJP%f%HV-NC%f%JV6-GWP%f%T6-K2%f%6JX-C7P%f%BG_Retail________Professional2021Retail
16_3747d1d5-55a8-4bc3-b53d-19fff1913195_QKH%f%NX-M9%f%GGH-T3Q%f%MW-YP%f%K4Q-QRW%f%MV_Retail________ProjectPro2021Retail
16_17739068-86c4-4924-8633-1e529abc7efc_HVC%f%34-CV%f%NPG-RVC%f%MT-X2%f%JRF-CR7%f%RK_MAK-AE1_______ProjectPro2021Volume
16_4ea64dca-227c-436b-813f-b6624be2d54c_2B9%f%6V-X9%f%NJY-WFB%f%RC-Q8%f%MP2-7CH%f%RR_Retail________ProjectStd2021Retail
16_84313d1e-47c8-4e27-8ced-0476b7ee46c4_3CN%f%QX-T3%f%4TY-99R%f%H4-C4%f%YD2-KW6%f%WH_MAK-AE________ProjectStd2021Volume
16_b769b746-53b1-4d89-8a68-41944dafe797_CDN%f%FG-77%f%T8D-VKQ%f%JX-B7%f%KT3-KK2%f%8V_Retail1_______Publisher2021Retail
16_a0234cfe-99bd-4586-a812-4f296323c760_2KX%f%JH-3N%f%HTW-RDB%f%PX-QF%f%RXJ-MTG%f%XF_MAK-AE________Publisher2021Volume
16_c3fb48b2-1fd4-4dc8-af39-819edf194288_DVB%f%XN-HF%f%T43-CVP%f%RQ-J8%f%9TF-VMM%f%HG_Retail________SkypeforBusiness2021Retail
16_6029109c-ceb8-4ee5-b324-f8eb2981e99a_R3F%f%CY-NH%f%GC7-CBP%f%VP-8Q%f%934-YTG%f%XG_MAK-AE________SkypeforBusiness2021Volume
16_9e7e7b8e-a0e7-467b-9749-d0de82fb7297_HXN%f%XB-J4%f%JGM-TCF%f%44-2X%f%2CV-FJV%f%VH_Retail________Standard2021Retail
16_223a60d8-9002-4a55-abac-593f5b66ca45_2CJ%f%N4-C9%f%XK2-HFP%f%Q6-YH%f%498-82T%f%XH_MAK-AE________Standard2021Volume
16_b99ba8c4-e257-4b70-a31a-8bd308ce7073_BQW%f%DW-NJ%f%9YF-P7Y%f%79-H6%f%DCT-MKQ%f%9C_MAK-AE________StandardSPLA2021Volume
16_814014d3-c30b-4f63-a493-3708e0dc0ba8_T6P%f%26-NJ%f%VBR-76B%f%K8-WB%f%CDY-TX3%f%BC_Retail________VisioPro2021Retail
16_c590605a-a08a-4cc7-8dc2-f1ffb3d06949_JNK%f%BX-MH%f%9P4-K8Y%f%YV-8C%f%G2Y-VQ2%f%C8_MAK-AE________VisioPro2021Volume
16_16d43989-a5ef-47e2-9ff1-272784caee24_89N%f%YY-KB%f%93R-7X2%f%2F-93%f%QDF-DJ6%f%YM_Retail________VisioStd2021Retail
16_d55f90ee-4ba2-4d02-b216-1300ee50e2af_BW4%f%3B-4P%f%NFP-V63%f%7F-23%f%TR2-J47%f%TX_MAK-AE________VisioStd2021Volume
16_fb33d997-4aa3-494e-8b58-03e9ab0f181d_VNC%f%C4-CJ%f%QVK-BKX%f%34-77%f%Y8H-CYX%f%MR_Retail________Word2021Retail
16_0c728382-95fb-4a55-8f12-62e605f91727_BJG%f%97-NW%f%3GM-8QQ%f%Q7-FH%f%76G-686%f%XM_MAK-AE________Word2021Volume
16_6337137e-7c07-4197-8986-bece6a76fc33_2P3%f%C9-BQ%f%NJH-VCV%f%PH-YD%f%Y6M-43J%f%PQ_Subscription__O365BusinessRetail
16_2f5c71b4-5b7a-4005-bb68-f9fac26f2ea3_W62%f%NQ-26%f%7QR-RTF%f%74-PF%f%2MH-JQM%f%TH_Subscription__O365EduCloudRetail
16_537ea5b5-7d50-4876-bd38-a53a77caca32_J2W%f%28-TN%f%9C8-26P%f%WV-F7%f%J4G-72X%f%CB_Subscription1_O365HomePremRetail
16_149dbce7-a48e-44db-8364-a53386cd4580_2N3%f%82-D6%f%PKK-QTX%f%4D-2J%f%JYK-M96%f%P2_Subscription1_O365ProPlusRetail
16_bacd4614-5bef-4a5e-bafc-de4c788037a2_HN8%f%JP-87%f%TQJ-PBF%f%3P-Y6%f%6KC-W2K%f%9V_Subscription1_O365SmallBusPremRetail
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do (
if %1==getinfo if %oVer%==%%A if "%2"=="%%E" (set _key=%%C& set _actid=%%B& set _allactid=!_allactid! %%B&set _lic=%%D)

if %1==getmsiprod if %oVer%==%%A (find /i "%%E" %msitemp% %nul% && (if defined _oIds (set _oIds=!_oIds! %%E) else (set _oIds=%%E)))

)
)
exit /b

::========================================================================================================================================

::  This code is used to modify the timestamp value of sppc dll file in order to change checksums
::  It's done to lower the potential false positive detection by antivirus's. On each install, it will install a unique sppc dll file

:oh_modify

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':hexedit\:.*';& ([ScriptBlock]::Create($f[1])) '%1';" %nul2% | find /i "Error found" %nul1% && set hasherror=1
exit /b

:hexedit:
param (
    [Parameter()]
    [String]$PePath
)

# Define dynamic assembly, module, and type
$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TypeBuilder = $ModuleBuilder.DefineType(0)

# Define P/Invoke method
[void]$TypeBuilder.DefinePInvokeMethod('MapFileAndCheckSum', 'imagehlp.dll', 'Public, Static', [Reflection.CallingConventions]::Standard, [int], @([string], [int].MakeByRefType(), [int].MakeByRefType()), [Runtime.InteropServices.CallingConvention]::Winapi, [Runtime.InteropServices.CharSet]::Auto)

# Create the type
$Imagehlp = $TypeBuilder.CreateType()

# File and offset information
$PeFile = Get-ChildItem -Path $PePath
$timestampOffset = 136
$exportTimestampOffset = 3076
$checkSumOffset = 216

# Calculate timestamp
$currentTimestamp = [DateTime]::UtcNow
$unixTimestamp = [int]($currentTimestamp - (Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0)).TotalSeconds

# Open file
$Stream = [System.IO.File]::Open($PeFile.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
$Writer = New-Object System.IO.BinaryWriter($Stream)

# Change timestamps
$Writer.BaseStream.Position = $timestampOffset
$Writer.Write($unixTimestamp)

$Writer.BaseStream.Position = $exportTimestampOffset
$Writer.Write($unixTimestamp)

$Writer.Flush()

# Update hash
[int]$HeaderSum = 0
[int]$CheckSum = 0

[void]$Imagehlp::MapFileAndCheckSum($PeFile.FullName, [ref]$HeaderSum, [ref]$CheckSum)

if ($HeaderSum -ne $CheckSum) {
    $Writer.BaseStream.Position = $checkSumOffset
    $Writer.Write($CheckSum)
    $Writer.Flush()
} else {
    Write-host Error found
}

[void]$Imagehlp::MapFileAndCheckSum($PeFile.FullName, [ref]$HeaderSum, [ref]$CheckSum)

if ($HeaderSum -ne $CheckSum) {
    Write-host Error found
}

$Stream.Close()
:hexedit:

::========================================================================================================================================
:: Leave empty line below
