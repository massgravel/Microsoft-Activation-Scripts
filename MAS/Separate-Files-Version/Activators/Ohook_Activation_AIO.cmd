@set masver=2.8
@echo off



::============================================================================
::
::   Homepage: mass grave[.]dev
::      Email: mas.help@outlook.com
::
::============================================================================



::  To activate Office with Ohook activation, run the script with "/Ohook" parameter or change 0 to 1 in below line
set _act=0

::  To remove Ohook activation, run the script with /Ohook-Uninstall parameter or change 0 to 1 in below line
set _rem=0

::  To run the script in debug mode, change 0 to "/Ohook" in below line
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
title  Ohook Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args set _args=%_args:re1=%
if defined _args set _args=%_args:re2=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/Ohook"                  set _act=1
if /i "%%A"=="/Ohook-Uninstall"        set _rem=1
if /i "%%A"=="-el"                     set _elev=1
)
)

for %%A in (%_act% %_rem%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

call :dk_setvar

if %winbuild% LSS 9200 (
%eline%
echo Unsupported OS version detected [%winbuild%].
echo Ohook Activation is supported only on Windows 8/10/11 and their server equivalents.
echo:
call :dk_color %Blue% "Use Online KMS activation option instead."
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
cls

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

set officeact=1
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

::  Check already activated products list

set actiProds15=
set actiProds16=

if not "%o15c2r%%o15msi%"=="" call :oh_findactivated -like 15
if not "%o16c2r%%o16msi%"=="" call :oh_findactivated -notlike 16

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

if "%_oArch%"=="x64" (set "_hookPath=%_oRoot%\vfs\System"    & set "_hook=sppc64.dll")
if "%_oArch%"=="x86" (set "_hookPath=%_oRoot%\vfs\SystemX86" & set "_hook=sppc32.dll")
if not "%osarch%"=="x86" (
if "%_oArch%"=="x64" set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if "%_oArch%"=="x86" set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
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

if "%_oArch%"=="x64" (set "_hookPath=%_oRoot%\vfs\System"    & set "_hook=sppc64.dll")
if "%_oArch%"=="x86" (set "_hookPath=%_oRoot%\vfs\SystemX86" & set "_hook=sppc32.dll")
if not "%osarch%"=="x86" (
if "%_oArch%"=="x64" set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if "%_oArch%"=="x86" set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
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

echo: !actiProds%oVer%! | find /i "-%%#-" %nul1% && (
call :dk_color %Gray% "Checking Activation Status              [%%# is already permanently activated]"

) || (

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

if "%_oArch%"=="x64" (set "_hookPath=%_oRoot%" & set "_hook=sppc64.dll")
if "%_oArch%"=="x86" (set "_hookPath=%_oRoot%" & set "_hook=sppc32.dll")
if not "%osarch%"=="x86" (
if "%_oArch%"=="x64" set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if "%_oArch%"=="x86" set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
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

:oh_findactivated

set oVer=%2
set _FsortIds=
set actiProds=

for /f "delims=" %%a in ('%psc% "(Get-WmiObject -Query 'SELECT LicenseFamily, Name FROM %spp% WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663'' AND LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL' | Where-Object { $_.Name %1 '*Office 15*' }).LicenseFamily" %nul6%') do call set "actiProds=%%a !actiProds!"

if not defined actiProds exit /b

for %%# in (%actiProds%) do (
set _sortIds=%%#
set _sortIds=!_sortIds:OfficeSPDFreeR_=SPDRetail_!
set _sortIds=!_sortIds:XC2RVL_=XVolume_!
set _sortIds=!_sortIds:CO365R_=Retail_!
set _sortIds=!_sortIds:O365R_=Retail_!
set _sortIds=!_sortIds:E5R_=Retail_!
set _sortIds=!_sortIds:MSDNR_=Retail_!
set _sortIds=!_sortIds:DemoR_=Retail_!
set _sortIds=!_sortIds:EDUR_=Retail_!
set _sortIds=!_sortIds:R_=Retail_!
set _sortIds=!_sortIds:VL_=Volume_!
set _sortIds=!_sortIds:Office16=!
set _sortIds=!_sortIds:Office19=!
set _sortIds=!_sortIds:Office21=!
set _sortIds=!_sortIds:Office24=!
set _sortIds=!_sortIds:Office=!
for /f "tokens=1 delims=-_" %%a in ("!_sortIds!") do set "_sortIds=-%%a-"
set _FsortIds=!_sortIds! !_FsortIds!
)

call :ohookdata findactivated %2
exit /b

::  Below IDs are not checked for permanent activation
set _sortIds=!_sortIds:PreviewVL_=Volume_!
set _sortIds=!_sortIds:PreInstallR_=Retail_!

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
reg delete HKU\%%#\Software\Microsoft\Office\%%x.0\Common\Identity /f %nul%

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
reg delete HKU\!defname!\Software\Microsoft\Office\16.0\Common\Identity /f %nul%
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

if defined o16c2r if defined officeact (
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
set _allactid=
for /f "delims=" %%a in ('%psc% "(Get-WmiObject -Query 'SELECT ID, Description, LicenseFamily FROM %spp% WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663'' AND PartialProductKey IS NOT NULL' | Where-Object { $_.LicenseFamily -notmatch 'Project' -and $_.LicenseFamily -notmatch 'Visio' }).ID" %nul6%') do call set "_allactid=%%a !_allactid!"
for /f "delims=" %%a in ('%psc% "(Get-WmiObject -Query 'SELECT ID, Description, LicenseFamily FROM %spp% WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663'' AND PartialProductKey IS NOT NULL' | Where-Object { $_.Description -match 'KMSCLIENT' -and ($_.LicenseFamily -match 'Project' -or $_.LicenseFamily -match 'Visio') }).ID" %nul6%') do call set "_allactid=%%a !_allactid!"
) else (
for /f "delims=" %%a in ('%psc% "(Get-WmiObject -Query 'SELECT ID FROM %spp% WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663'' AND LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL').ID" %nul6%') do call set "_allactid=%%a !_allactid!"
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

if defined officeact if not %upk_result%==0 echo:
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
15_4374022d-56b8-48c1-9bb7-d8f2fc726343_9MF9G-CN32B-HV7XT-9XJ8T-9K%f%VF4_MAK___________AccessVolume
15_1b1d9bd5-12ea-4063-964c-16e7e87d6e08_NT889-MBH4X-8MD4H-X8R2D-WQ%f%HF8_Retail________ExcelRetail
15_ac1ae7fd-b949-4e04-a330-849bc40638cf_Y3N36-YCHDK-XYWBG-KYQVV-BD%f%TJ2_MAK___________ExcelVolume
15_cfaf5356-49e3-48a8-ab3c-e729ab791250_BMK4W-6N88B-BP9QR-PHFCK-MG%f%7GF_Retail________GrooveRetail
15_4825ac28-ce41-45a7-9e6e-1fed74057601_RN84D-7HCWY-FTCBK-JMXWM-HT%f%7GJ_MAK___________GrooveVolume
15_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_2WQNF-GBK4B-XVG6F-BBMX7-M4%f%F2Y_OEM-Perp______HomeBusinessPipcRetail
15_a2b90e7a-a797-4713-af90-f0becf52a1dd_YWD4R-CNKVT-VG8VJ-9333B-RC%f%W9F_Subscription__HomeBusinessRetail
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
15_537ea5b5-7d50-4876-bd38-a53a77caca32_J2W28-TN9C8-26PWV-F7J4G-72%f%XCB_Subscription1_O365HomePremRetail
15_e3dacc06-3bc2-4e13-8e59-8e05f3232325_H8DN8-Y2YP3-CR9JT-DHDR9-C7%f%GP3_Subscription2_O365ProPlusRetail
15_bacd4614-5bef-4a5e-bafc-de4c788037a2_HN8JP-87TQJ-PBF3P-Y66KC-W2%f%K9V_Subscription1_O365SmallBusPremRetail
:: Office 365 - 16.0 version
16_742178ed-6b28-42dd-b3d7-b7c0ea78741b_Y9NF9-M2QWD-FF6RJ-QJW36-RR%f%F2T_SubTest_______O365BusinessRetail
16_2f5c71b4-5b7a-4005-bb68-f9fac26f2ea3_W62NQ-267QR-RTF74-PF2MH-JQ%f%MTH_Subscription__O365EduCloudRetail
16_537ea5b5-7d50-4876-bd38-a53a77caca32_J2W28-TN9C8-26PWV-F7J4G-72%f%XCB_Subscription1_O365HomePremRetail
16_e3dacc06-3bc2-4e13-8e59-8e05f3232325_H8DN8-Y2YP3-CR9JT-DHDR9-C7%f%GP3_Subscription2_O365ProPlusRetail
16_bacd4614-5bef-4a5e-bafc-de4c788037a2_HN8JP-87TQJ-PBF3P-Y66KC-W2%f%K9V_Subscription1_O365SmallBusPremRetail
:: Office 2016
16_bfa358b0-98f1-4125-842e-585fa13032e6_WHK4N-YQGHB-XWXCC-G3HYC-6J%f%F94_Retail________AccessRetail
16_9d9faf9e-d345-4b49-afce-68cb0a539c7c_RNB7V-P48F4-3FYY6-2P3R3-63%f%BQV_PrepidBypass__AccessRuntimeRetail
16_3b2fa33f-cd5a-43a5-bd95-f49f3f546b0b_JJ2Y4-N8KM3-Y8KY3-Y22FR-R3%f%KVK_MAK___________AccessVolume
16_424d52ff-7ad2-4bc7-8ac6-748d767b455d_RKJBN-VWTM2-BDKXX-RKQFD-JT%f%YQ2_Retail________ExcelRetail
16_685062a7-6024-42e7-8c5f-6bb9e63e697f_FVGNR-X82B2-6PRJM-YT4W7-8H%f%V36_MAK___________ExcelVolume
16_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_2WQNF-GBK4B-XVG6F-BBMX7-M4%f%F2Y_OEM-Perp______HomeBusinessPipcRetail
16_86834d00-7896-4a38-8fae-32f20b86fa2b_HM6FM-NVF78-KV9PM-F36B8-D9%f%MXD_Retail________HomeBusinessRetail
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
16_fb099c19-d48b-4a2f-a160-4383011060aa_V6QFB-7N7G9-PF7W9-M8FQM-MY%f%8G9_Retail________Excel2021Retail
16_9da1ecdb-3a62-4273-a234-bf6d43dc0778_WNYR4-KMR9H-KVC8W-7HJ8B-K7%f%9DQ_MAK-AE________Excel2021Volume
16_38b92b63-1dff-4be7-8483-2a839441a2bc_JM99N-4MMD8-DQCGJ-VMYFY-R6%f%3YK_Subscription__HomeBusiness2021Retail
16_2f258377-738f-48dd-9397-287e43079958_N3CWD-38XVH-KRX2Y-YRP74-6R%f%BB2_Subscription__HomeStudent2021Retail
16_279706f4-3a4b-4877-949b-f8c299cf0cc5_NB2TQ-3Y79C-77C6M-QMY7H-7Q%f%Y8P_Retail________OneNote2021Retail
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

if %1==findactivated if %oVer%==%%A (
echo "!_FsortIds!" | find /i "-%%E-" %nul% && (
set actiProds%oVer%=!actiProds%oVer%! -%%E-
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

::========================================================================================================================================
:: Leave empty line below
