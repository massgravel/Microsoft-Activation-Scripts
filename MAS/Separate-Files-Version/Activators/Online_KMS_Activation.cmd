@set masver=2.5
@setlocal DisableDelayedExpansion
@echo off



::=================================================================================================
::
::  Online KMS Script is a fork of @abbodi1406's KMS_VL_ALL  forums.mydigitallife.net/posts/838808
::  
::  This fork's purpose is to avoid having any KMS binary files and activate Windows/Office using 
::  only transparent batch script with online public KMS servers.
::_____________________________________
::
::  Online KMS Activation Script is a part of 'Microsoft_Activation_Scripts' (MAS) project.
::  
::  Homepage: mass grave[.]dev
::     Email: windowsaddict@protonmail.com
::  
::=================================================================================================



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
title  Online KMS Activation %masver%

::  You are not supposed to edit anything below this.

set WMI_VBS=0
set _Debug=0
set Silent=0
set Logger=0
set AutoR2V=1
set SkipKMS38=1
set vNextOverride=1
set ActWindows=1
set ActOffice=1

set _uni=
set _args=
set _elev=
set _renetask=
set _renacttask=
set _unattended=
set _unattendedact=

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
echo "%_args%" | find /i "/KMS" >nul && set _unattended=1

for %%A in (%_args%) do (
if /i "%%A"=="-el"  (set _elev=1
) else if /i "%%A"=="/KMS-RenewalTask"  (set _renetask=1
) else if /i "%%A"=="/KMS-ActAndRenewalTask" (set _renacttask=1
) else if /i "%%A"=="/KMS-Uninstall" (set _uni=1
) else if /i "%%A"=="/KMS-Windows"   (set ActWindows=1&set ActOffice=0&set _unattendedact=1
) else if /i "%%A"=="/KMS-Office"   (set ActWindows=0&set ActOffice=1&set _unattendedact=1
) else if /i "%%A"=="/KMS-WindowsOffice"  (set ActWindows=1&set ActOffice=1&set _unattendedact=1
) else if /i "%%A"=="/KMS-KeepvNext"  (set vNextOverride=0
) else if /i "%%A"=="/KMS-Debug"   (set _Debug=1
) else if /i "%%A"=="/KMS-Logger"   (set Logger=1&set Silent=1
)
)
)

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

call :_colorprep
set "_buf={$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

set "nceline=echo. &echo ==== ERROR ==== &echo."
set "eline=echo. &call :_color %Red% "==== ERROR ====" &echo."
if %_Debug% EQU 1 set _unattended=1

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected [%winbuild%].
echo Project is supported for Windows 7/8/8.1/10/11 and their Server equivalent.
goto Done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto Done
)

::========================================================================================================================================

::  Fix special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%userprofile%\AppData\Local\Temp"
set "_Local=%LocalAppData%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
if /i not "!_work!"=="!_ttemp!" (
%nceline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo.
echo Extract the archive file and launch the script from the extracted folder.
goto Done
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%nceline%
echo This script needs admin rights.
echo To do so, right click on this script and select 'Run as administrator'.
goto Done
)

::========================================================================================================================================

::  This code disables QuickEdit for this cmd.exe session only without making permanent changes to the registry
::  It is added because clicking on the script window pauses the operation and leads to the confusion that script stopped due to an error

if defined _unattended set quedit=1
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
if not defined _unattended (
echo [1] Get Latest MAS
echo [0] Continue Anyway
echo:
call :_color %_Green% "Enter a menu option in the Keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
)
)
cls

::========================================================================================================================================

if %~z0 GEQ 300000 (set "_exitmsg=Go back") else (set "_exitmsg=Exit")

::  Check not x86 Windows

set notx86=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b
if /i not "%arch%"=="x86" set notx86=1

::========================================================================================================================================

for %%# in (wmic.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find wmic.exe in the system.
if %winbuild% GEQ 22621 echo Make sure WMIC is enabled in optional features.
goto Done
)

wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "ComputerSystem" 1>nul || (
%nceline%
echo WMI is not responding in the system.
echo:
echo In MAS, Goto Troubleshoot and run Fix WMI option.
goto Done
)

set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)

if %_WSH% EQU 0 (
reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
reg add "HKCU\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
if defined notx86 reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f /reg:32 %nul%
)

::========================================================================================================================================

if defined _uni goto _Complete_Uninstall

if defined _renetask set ActTask=&call:RenTask&timeout /t 2
cls
if defined _renacttask set ActTask=1&call:RenTask&timeout /t 2
cls
if defined _unattended if not defined _unattendedact goto Done

::========================================================================================================================================

set "_title=Online KMS Activation %masver%"
set _gui=

:_KMS_Menu

set sub_next=0
set sub_o365=0
set sub_proj=0
set sub_vsio=0
set kNext=HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext
reg query %kNext% /v MigrationToV5Done 2>nul | find /i "0x1" %nul% && call :officeSub %nul%

set _tskinstalled=
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks" /f Path /s | find /i "\Activation-Renewal" >nul && (
find /i "Ver:1.9" "%ProgramFiles%\Activation-Renewal\Activation_task.cmd" %nul% && set _tskinstalled=1
)

set _oldtsk=
if not defined _tskinstalled (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks" /f Path /s | findstr /i "\Activation-Renewal \Online_KMS_Activation_Script-Renewal" >nul && (
set _oldtsk=1
)
)

if defined _unattended (
call :Activation_Start
timeout /t 2
goto Done
)

cls
set _gui=1
title  %_title%
mode con: cols=76 lines=30

echo.
echo.
echo.
echo.
echo.       ______________________________________________________________
echo.
echo.              [1] Activate - Windows
echo.              [2] Activate - Office
echo.              [3] Activate - All
echo.
if defined _tskinstalled call :_color2 %_White% "              [4] Install Auto-Renewal      " %_Green% "[Installed]"
if defined _oldtsk       call :_color2 %_White% "              [4] Install Auto-Renewal      " %_Red% "[Old Installed]"
if not defined _tskinstalled if not defined _oldtsk echo.              [4] Install Auto-Renewal
echo.              [5] Uninstall
echo.              _______________________________________________  
echo.
if %_Debug%==0 (
echo.              [6] Enable Debug Mode         [No]
) else (
call :_color2 %_White% "              [6] Enable Debug Mode         " %_Red% "[Yes]"
)
if %vNextOverride% EQU 1 (
if %sub_next% EQU 1 (
call :_color2 %_White% "              [7] Override Office vNext     " %_Red% "[Yes]"
) else (
echo               [7] Override Office vNext     [Yes]
)
) else (
if %sub_next% EQU 1 (
call :_color2 %_White% "              [7] Override Office vNext     " %_Yellow% "[No]"
) else (
echo               [7] Override Office vNext     [No]
)
)
echo.              _______________________________________________       
echo.
echo.              [0] %_exitmsg%
echo.       ______________________________________________________________
echo.
call :_color2 %_White% "           " %_Green% "Enter a menu option in the Keyboard [1,2,3,4,5,6,7,0]"
choice /C:12345670 /N
set _el=%errorlevel%

if %_el%==8 exit /b
if %_el%==7 (if %vNextOverride% EQU 0 (set vNextOverride=1) else (set vNextOverride=0))&goto _KMS_Menu
if %_el%==6 (if %_Debug%==0 (set _Debug=1) else (set _Debug=0)) &goto _KMS_Menu
if %_el%==5 call:_Complete_Uninstall&cls&goto _KMS_Menu
if %_el%==4 set ActTask=&call:RenTask&goto _KMS_Menu
if %_el%==3 cls&setlocal&set "ActWindows=1"&set "ActOffice=1"&call :Activation_Start&endlocal&cls&goto _KMS_Menu
if %_el%==2 cls&setlocal&set "ActWindows=0"&set "ActOffice=1"&call :Activation_Start&endlocal&cls&goto _KMS_Menu
if %_el%==1 cls&setlocal&set "ActWindows=1"&set "ActOffice=0"&call :Activation_Start&endlocal&cls&goto _KMS_Menu
goto _KMS_Menu

::========================================================================================================================================

:Done

if defined _unattended exit /b

echo.
echo Press any key to exit...
pause >nul
exit /b

:=========================================================================================================================================

:Activation_Start

@setlocal DisableDelayedExpansion

set nil=
for %%# in (SppE%nil%xtComObj.exe,sppsvc.exe,osppsvc.exe) do (
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%nil%ge File Execu%nil%tion Options\%%#" /f %nul%)
)

call :Clear-KMS-Cache %nul%

set "_Null=1>nul 2>nul"
set KMS_Port=1688
if %_Debug% EQU 1 set _unattended=1
set "_run=nul"
if %Logger% EQU 1 set _run="%~dpn0_Silent.log"

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)
set "_bit=64"
set "_wow=1"
if /i "%PROCESSOR_ARCHITECTURE%"=="amd64" set "xBit=x64"&set "xOS=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "xBit=x86"&set "xOS=A64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "xBit=x86"&set "xOS=x86"&set "_wow=0"&set "_bit=32"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "xBit=x64"&set "xOS=x64"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "xBit=x86"&set "xOS=A64"
if not defined xBit set "xBit=x64"&set "xOS=x64"

set _cwmi=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "ComputerSystem" 1>nul && set _cwmi=1
)

set "_Local=%LocalAppData%"
set "_temp=%SystemRoot%\Temp"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
set _UNC=0
if "%_work:~0,2%"=="\\" (
set _UNC=1
) else (
net use %~d0 %_Null%
if not errorlevel 1 set _UNC=1
)
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%PUBLIC%\Desktop\desktop.ini" set "_dsk=%PUBLIC%\Desktop"
set "_mO21a=Detected Office 2021 C2R Retail is activated"
set "_mO19a=Detected Office 2019 C2R Retail is activated"
set "_mO16a=Detected Office 2016 C2R Retail is activated"
set "_mO15a=Detected Office 2013 C2R Retail is activated"
set "_mO21c=Detected Office 2021 C2R Retail could not be converted to Volume"
set "_mO19c=Detected Office 2019 C2R Retail could not be converted to Volume"
set "_mO16c=Detected Office 2016 C2R Retail could not be converted to Volume"
set "_mO15c=Detected Office 2013 C2R Retail could not be converted to Volume"
set "_mO14c=Detected Office 2010 C2R Retail is not supported by this script"
set "_mO14m=Detected Office 2010 MSI Retail is not supported by this script"
set "_mO15m=Detected Office 2013 MSI Retail is not supported by this script"
set "_mO16m=Detected Office 2016 MSI Retail is not supported by this script"
set "_mOuwp=Detected Office 365/2016 UWP is not supported by this script"
set DO15Ids=ProPlus,Standard,Access,Lync,Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set DO16Ids=ProPlus,Standard,Access,SkypeforBusiness,Excel,Outlook,PowerPoint,Publisher,Word
set LV16Ids=Mondo,ProPlus,ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness,OneNote,Excel,Outlook,PowerPoint,Publisher,Word
set LR16Ids=%LV16Ids%,Professional,HomeBusiness,HomeStudent,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set "ESUEditions=Enterprise,EnterpriseE,EnterpriseN,Professional,ProfessionalE,ProfessionalN,Ultimate,UltimateE,UltimateN"
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
set "ESUEditions=ServerDatacenter,ServerDatacenterCore,ServerDatacenterV,ServerDatacenterVCore,ServerStandard,ServerStandardCore,ServerStandardV,ServerStandardVCore,ServerEnterprise,ServerEnterpriseCore,ServerEnterpriseV,ServerEnterpriseVCore"
)
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set UBR=0
if %winbuild% GEQ 7601 for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR 2^>nul') do if not errorlevel 1 set /a UBR=%%b
set "_csq=cscript.exe //NoLogo //Job:WmiQuery "%~nx0?.wsf""
set "_csm=cscript.exe //NoLogo //Job:WmiMethod "%~nx0?.wsf""
set "_csp=cscript.exe //NoLogo //Job:WmiPKey "%~nx0?.wsf""
set "_csd=cscript.exe //NoLogo //Job:MPS "%~nx0?.wsf""
if %_cwmi% EQU 0 set WMI_VBS=1
if %WMI_VBS% EQU 0 (
set "_zz1=wmic path"
set "_zz2=where"
set "_zz3=get"
set "_zz4=/value"
set "_zz5=("
set "_zz6=)"
set "_zz7="wmic path"
set "_zz8=/value""
) else (
set "_zz1=%_csq%"
set "_zz2="
set "_zz3="
set "_zz4="
set "_zz5=""
set "_zz6=""
set "_zz7=%_csq%"
set "_zz8="
)

setlocal EnableDelayedExpansion
pushd "!_work!"

if not defined _unattended (
mode con cols=98 lines=31
%psc% "&%_buf%"
title  %_title%
) else (
title  Online KMS Activation %masver%
)

if defined _gui if %_Debug%==1 mode con cols=98 lines=30

if %_Debug% EQU 0 (
  set "_Nul1=1>nul"
  set "_Nul2=2>nul"
  set "_Nul6=2^>nul"
  set "_Nul3=1>nul 2>nul"
  set "_Pause=pause >nul"
  if %Silent% EQU 0 (call :Begin) else (call :Begin >!_run! 2>&1)
) else (
  set "_Nul1="
  set "_Nul2="
  set "_Nul6="
  set "_Nul3="
  set "_log=!_dsk!\%~n0"
  if %Silent% EQU 0 (
  echo.
  echo Running in Debug Mode...
  if not defined _args (echo The window will be closed when finished) else (echo please wait...)
  echo.
  echo Writing debug log to:
  echo "!_log!_Debug.log"
  )
  @echo on
  @prompt $G
  @call :Begin >"!_log!_tmp.log" 2>&1 &cmd /u /c type "!_log!_tmp.log">"!_log!_Debug.log"&del "!_log!_tmp.log"
)
@echo off
if defined _gui if %_Debug%==1 (
echo.
call :_color %_Yellow% "Press any key to go back..."
pause >nul
exit /b
)
@exit /b

:Begin

::========================================================================================================================================

set act_failed=0
set /a act_attempt=0

echo.
echo Initializing...

:: Check Internet connection. Works even if ICMP echo is disabled.

call :setserv
for %%a in (%srvlist%) do (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (
if not [%%#]==[] goto IntConnected
)
)

nslookup dns.msftncsi.com 2>nul | find "131.107.255.255" 1>nul
if [%errorlevel%]==[0] goto IntConnected

cls
if %_Debug%==1 (
echo Error: Internet is not connected.
exit /b
)

if defined _unattended (
echo.
call :_color %_Red% "Internet is not connected, continuing the process anyway."
) else (
%eline%
echo Internet is not connected.
echo:
call :_color %_Yellow% "Press any key to go back..."
pause >nul
exit /b
)

:IntConnected

call :getserv

::========================================================================================================================================

set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set SSppHook=0
for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\skus') do (
  if %winbuild% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*GVLK*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VLKMS*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VL-BYPASS*.xrm-ms" set SSppHook=1
)
set OsppHook=1
sc query osppsvc %_Nul3%
if %errorlevel% EQU 1060 set OsppHook=0

set ESU_KMS=0
if %winbuild% LSS 9200 for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\channels') do (
  if exist "%SysPath%\spp\tokens\channels\%%A\*VL-BYPASS*.xrm-ms" set ESU_KMS=1
)
if %ESU_KMS% EQU 1 (set "adoff=and LicenseDependsOn is NULL"&set "addon=and LicenseDependsOn is not NULL") else (set "adoff="&set "addon=")
set ESU_EDT=0
if %ESU_KMS% EQU 1 for %%A in (%ESUEditions%) do (
  if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%%A\*.xrm-ms" set ESU_EDT=1
)
:: if %ESU_EDT% EQU 1 set SSppHook=1
set ESU_ADD=0

if %winbuild% GEQ 9200 (
  set OSType=Win8
  set SppVer=SppExtComObj.exe
) else if %winbuild% GEQ 7600 (
  set OSType=Win7
  set SppVer=sppsvc.exe
) else (
  goto :UnsupportedVersion
)
if %OSType% EQU Win8 reg query "%IFEO%\sppsvc.exe" %_Nul3% && (
reg delete "%IFEO%\sppsvc.exe" /f %_Nul3%
call :StopService sppsvc
)

if %ActWindows% EQU 0 if %ActOffice% EQU 0 set ActWindows=1
set _AUR=1
if %winbuild% GEQ 9600 (
  reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f %_Nul3%
  if %winbuild% EQU 14393 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoAcquireGT /t REG_DWORD /d 1 /f %_Nul3%
)
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc

:ReturnHook
call :UpdateOSPPEntry osppsvc.exe

SET Win10Gov=0
SET "EditionWMI="
SET "EditionID="
IF %winbuild% LSS 14393 if %SSppHook% NEQ 0 GOTO :Main
SET "RegKey=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
SET "Pattern=Microsoft-Windows-*Edition~31bf3856ad364e35"
SET "EditionPKG=FFFFFFFF"
FOR /F "TOKENS=8 DELIMS=\" %%A IN ('REG QUERY "%RegKey%" /f "%Pattern%" /k %_Nul6% ^| FIND /I "CurrentVersion"') DO (
  REG QUERY "%RegKey%\%%A" /v "CurrentState" %_Nul2% | FIND /I "0x70" %_Nul1% && (
    FOR /F "TOKENS=3 DELIMS=-~" %%B IN ('ECHO %%A') DO SET "EditionPKG=%%B"
  )
)
IF /I "%EditionPKG:~-7%"=="Edition" (
SET "EditionID=%EditionPKG:~0,-7%"
) ELSE (
FOR /F "TOKENS=3 DELIMS=: " %%A IN ('DISM /English /Online /Get-CurrentEdition %_Nul6% ^| FIND /I "Current Edition :"') DO SET "EditionID=%%A"
)
net start sppsvc /y %_Nul3%
set "_qr=%_zz7% SoftwareLicensingProduct %_zz2% %_zz5%ApplicationID='%_wApp%' %adoff% AND PartialProductKey is not NULL%_zz6% %_zz3% LicenseFamily %_zz8%"
FOR /F "TOKENS=2 DELIMS==" %%A IN ('%_qr% %_Nul6%') DO SET "EditionWMI=%%A"
IF "%EditionWMI%"=="" (
IF %winbuild% GEQ 17063 FOR /F "SKIP=2 TOKENS=2*" %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%B"
IF %winbuild% LSS 14393 (
  FOR /F "SKIP=2 TOKENS=2*" %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%B"
  GOTO :Main
  )
)
IF NOT "%EditionWMI%"=="" SET "EditionID=%EditionWMI%"
IF /I "%EditionID%"=="IoTEnterprise" SET "EditionID=Enterprise"
IF /I "%EditionID%"=="IoTEnterpriseS" IF %winbuild% LSS 22610 (
SET "EditionID=EnterpriseS"
IF %winbuild% GEQ 19041 IF %UBR% GEQ 2788 SET "EditionID=IoTEnterpriseS"
)
IF /I "%EditionID%"=="ProfessionalSingleLanguage" SET "EditionID=Professional"
IF /I "%EditionID%"=="ProfessionalCountrySpecific" SET "EditionID=Professional"
IF /I "%EditionID%"=="EnterpriseG" SET Win10Gov=1
IF /I "%EditionID%"=="EnterpriseGN" SET Win10Gov=1

:Main
if defined EditionID (set "_winos=Windows %EditionID% edition") else (set "_winos=Detected Windows")
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName %_Nul6%') do if not errorlevel 1 set "_winos=%%b"
set "nKMS=does not support KMS activation..."
set "nEval=Evaluation Editions cannot be activated. Please install full Windows OS."
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _eval=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" set "nEval=Server Evaluation cannot be activated. Please convert to full Server OS."
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" set _eval=1&set "nEval=Server Evaluation cannot be activated. Please convert to full Server OS."
set "_C16R="
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C16R=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
)
if not defined _C16R reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C16R=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
)
set "_C15R="
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && call set "_C15R=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
if not defined _C15R reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && call set "_C15R=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag"
)
set "_C14R="
if %_wow%==0 (reg query HKLM\SOFTWARE\Microsoft\Office\14.0\CVH /f Click2run /k %_Nul3% && set "_C14R=1") else (reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\CVH /f Click2run /k %_Nul3% && set "_C14R=1")
for %%A in (14,15,16,19,21) do call :officeLoc %%A
if %_O14MSI% EQU 1 set "_C14R="

set S_OK=1
call :RunSPP
if %ActOffice% NEQ 0 call :RunOSPP
if %ActOffice% EQU 0 (echo.&echo Office activation is OFF...)

if exist "!_temp!\crv*.txt" del /f /q "!_temp!\crv*.txt"
if exist "!_temp!\*chk.txt" del /f /q "!_temp!\*chk.txt"
if exist "!_temp!\slmgr.vbs" del /f /q "!_temp!\slmgr.vbs"
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc

sc start sppsvc trigger=timer;sessionid=0 %_Nul3%

goto TheEnd

:RunSPP
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
set W1nd0ws=1
set WinPerm=0
set WinVL=0
set Off1ce=0
set RanR2V=0
set aC2R21=0
set aC2R19=0
set aC2R16=0
set aC2R15=0
if %winbuild% GEQ 9200 if %ActOffice% NEQ 0 call :sppoff
set "_qr=%_zz1% %spp% %_zz2% %_zz5%Description like '%%KMSCLIENT%%' %_zz6% %_zz3% Name %_zz4%"
%_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set WinVL=1)
if %WinVL% EQU 0 (
if %ActWindows% EQU 0 (
  echo.&echo Windows activation is OFF...
  ) else (
  if %SSppHook% EQU 0 (
    echo.&echo %_winos% %nKMS%
    if defined _eval echo %nEval%
    ) else (
    echo.&echo Failed checking KMS Activation ID^(s^) for Windows. &call :CheckWS
    exit /b
    )
  )
)
if %WinVL% EQU 0 if %Off1ce% EQU 0 exit /b
set _gvlk=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL%_zz6% %_zz3% Name %_zz4%"
if %winbuild% GEQ 10240 %_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set _gvlk=1)
set gpr=0
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL%_zz6% %_zz3% GracePeriodRemaining %_zz8%"
if %winbuild% GEQ 10240 if %SkipKMS38% NEQ 0 if %_gvlk% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set "gpr=%%A"
set "_qr=%_zz1% %spp% %_zz2% "ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL" %_zz3% LicenseFamily %_zz4%"
if %gpr% NEQ 0 if %gpr% GTR 259200 (
set W1nd0ws=0
%_qr% %_Nul2% | findstr /i EnterpriseG %_Nul1% && (call set W1nd0ws=1)
)
set "_qr=%_zz7% %sps% %_zz3% Version %_zz8%"
for /f "tokens=2 delims==" %%A in ('%_qr%') do set slsv=%%A
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32 %_Nul3%
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" /reg:32 %_Nul3%
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32 %_Null%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32 %_Nul3%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" /reg:32 %_Nul3%
)
reg delete "HKLM\%SPPk%\%_oApp%" /f %_Null%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' %_zz6% %_zz3% ID %_zz8%"
if %W1nd0ws% EQU 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkwin)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' %adoff% %_zz6% %_zz3% ID %_zz8%"
if %W1nd0ws% EQU 1 if %ActWindows% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkwin)
:: set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' %addon% %_zz6% %_zz3% ID %_zz8%"
:: if %ESU_EDT% EQU 1 if %ActWindows% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :esuchk)
if %W1nd0ws% EQU 1 if %ActWindows% EQU 0 (echo.&echo Windows activation is OFF...)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_oApp%' and Description like '%%KMSCLIENT%%' %_zz6% %_zz3% ID %_zz8%"
if %Off1ce% EQU 1 if %ActOffice% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkoff 1)
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing %_Null%
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching %_Null%
exit /b

:sppoff
set OffUWP=0
if %winbuild% GEQ 10240 reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msoxmled.exe" %_Nul3% && (
dir /b "%ProgramFiles%\WindowsApps\Microsoft.Office.Desktop*" %_Nul3% && set OffUWP=1
if not %xOS%==x86 dir /b "%ProgramW6432%\WindowsApps\Microsoft.Office.Desktop*" %_Nul3% && set OffUWP=1
)
rem nothing installed
if %loc_off21% EQU 0 if %loc_off19% EQU 0 if %loc_off16% EQU 0 if %loc_off15% EQU 0 (
if %winbuild% GEQ 9200 (
  if %OffUWP% EQU 0 (echo.&echo No Installed Office 2013-2021 Product Detected...) else (echo.&echo %_mOuwp%)
  exit /b
  )
if %winbuild% LSS 9200 (if %loc_off14% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b))
)
if %vNextOverride% EQU 1 if %AutoR2V% EQU 1 (
set sub_o365=0
set sub_proj=0
set sub_vsio=0
if %sub_next% EQU 1 reg delete HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing /f %_Nul3%
)
set Off1ce=1
set _sC2R=sppoff
set _fC2R=ReturnSPP

set vol_off14=0&set vol_off15=0&set vol_off16=0&set vol_off19=0&set vol_off21=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%Description like '%%KMSCLIENT%%' AND NOT Name like '%%MondoR_KMS_Automation%%' %_zz6% %_zz3% Name %_zz4%"
%_qr% > "!_temp!\sppchk.txt" 2>&1
find /i "Office 21" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off21=1)
find /i "Office 19" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off19=1)
find /i "Office 16" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off16=1)
find /i "Office 15" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off15=1)
if %winbuild% LSS 9200 find /i "Office 14" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off14=1)
for %%A in (14,15,16,19,21) do if !loc_off%%A! EQU 0 set vol_off%%A=0
set "_qr=%_zz1% %spp% %_zz2% "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%'" %_zz3% LicenseFamily %_zz4%"
if %vol_off16% EQU 1 find /i "Office16MondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% || (set vol_off16=0)
)
set "_qr=%_zz1% %spp% %_zz2% "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%'" %_zz3% LicenseFamily %_zz4%"
if %vol_off15% EQU 1 find /i "OfficeMondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% || (set vol_off15=0)
)

set ret_off14=0&set ret_off15=0&set ret_off16=0&set ret_off19=0&set ret_off21=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_oApp%' AND NOT Name like '%%O365%%' %_zz6% %_zz3% Name %_zz4%"
%_qr% > "!_temp!\sppchk.txt" 2>&1
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office 21" %_Nul1% && (set ret_off21=1)
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office 19" %_Nul1% && (set ret_off19=1)
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office 16" %_Nul1% && (set ret_off16=1)
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office 15" %_Nul1% && (set ret_off15=1)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_oA14%'%_zz6% %_zz3% Description %_zz4%"
if %winbuild% LSS 9200 if %vol_off14% EQU 0 %_qr% %_Nul2% | findstr /i channel %_Nul1% && (set ret_off14=1)

set run_off21=0&set prr_off21=0&set prv_off21=0
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 0 set run_off21=1
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 1 (
for %%a in (%DO16Ids%) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
for %%a in (Professional) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21ProPlus2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21Standard2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
)
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 1 if %prv_off21% LSS %prr_off21% (set vol_off21=0&set run_off21=1)

set run_off19=0&set prr_off19=0&set prv_off19=0
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 0 set run_off19=1
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 1 (
for %%a in (%DO16Ids%) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
for %%a in (Professional) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19ProPlus2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19Standard2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
)
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 1 if %prv_off19% LSS %prr_off19% (set vol_off19=0&set run_off19=1)

set run_off16=0&set prr_off16=0&set prv_off16=0
if %loc_off16% EQU 1 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if defined _C16R (
for %%a in (%DO16Ids%) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
for %%a in (Professional) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16ProPlusVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21ProPlus2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19ProPlus2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16StandardVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21Standard2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19Standard2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
)
if %loc_off16% EQU 1 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if defined _C16R if %prv_off16% LSS %prr_off16% (set vol_off16=0&set run_off16=1)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%' %_zz6% %_zz3% LicenseFamily %_zz4%"
if %loc_off16% EQU 1 if %run_off16% EQU 0 if %sub_o365% EQU 0 if defined _C16R %_qr% %_Nul2% | find /i "O365" %_Nul1% && (
find /i "Office16MondoVL" "!_temp!\sppchk.txt" %_Nul1% || set run_off16=1
)

set run_off15=0&set prr_off15=0&set prv_off15=0
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 0 if defined _C15R set run_off15=1
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 1 if defined _C15R (
for %%a in (%DO15Ids%) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
for %%a in (Professional) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "OfficeProPlusVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "OfficeStandardVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
)
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 1 if defined _C15R if %prv_off15% LSS %prr_off15% (set vol_off15=0&set run_off15=1)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%' %_zz6% %_zz3% LicenseFamily %_zz4%"
if %loc_off15% EQU 1 if %run_off15% EQU 0 if defined _C15R %_qr% %_Nul2% | find /i "O365" %_Nul1% && (
find /i "OfficeMondoVL" "!_temp!\sppchk.txt" %_Nul1% || set run_off15=1
)

set vol_offgl=1
if %vol_off21% EQU 0 if %vol_off19% EQU 0 if %vol_off16% EQU 0 if %vol_off15% EQU 0 (
if %winbuild% GEQ 9200 set vol_offgl=0
if %winbuild% LSS 9200 if %vol_off14% EQU 0 set vol_offgl=0
)
rem mixed Volume + Retail
if %run_off21% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off19% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off16% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off15% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
rem all supported Volume + message for unsupported
if %loc_off16% EQU 0 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if %OffUWP% EQU 1 (echo.&echo %_mOuwp%)
if %vol_offgl% EQU 1 (
if %ret_off16% EQU 1 if %_O16MSI% EQU 1 (echo.&echo %_mO16m%)
if %ret_off15% EQU 1 if %_O15MSI% EQU 1 (echo.&echo %_mO15m%)
if %winbuild% LSS 9200 if %loc_off14% EQU 1 if %vol_off14% EQU 0 (if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%))
exit /b
)
set Off1ce=0
rem Retail C2R
if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
:ReturnSPP
rem Retail MSI/C2R or failed C2R-R2V
if %loc_off21% EQU 1 if %vol_off21% EQU 0 (
if %aC2R21% EQU 1 (echo.&echo %_mO21a%) else (echo.&echo %_mO21c%)
)
if %loc_off19% EQU 1 if %vol_off19% EQU 0 (
if %aC2R19% EQU 1 (echo.&echo %_mO19a%) else (echo.&echo %_mO19c%)
)
if %loc_off16% EQU 1 if %vol_off16% EQU 0 (
if defined _C16R (if %aC2R16% EQU 1 (echo.&echo %_mO16a%) else (if %sub_o365% EQU 0 echo.&echo %_mO16c%)) else if %_O16MSI% EQU 1 (if %ret_off16% EQU 1 echo.&echo %_mO16m%)
)
if %loc_off15% EQU 1 if %vol_off15% EQU 0 (
if defined _C15R (if %aC2R15% EQU 1 (echo.&echo %_mO15a%) else (echo.&echo %_mO15c%)) else if %_O15MSI% EQU 1 (if %ret_off15% EQU 1 echo.&echo %_mO15m%)
)
if %winbuild% LSS 9200 if %loc_off14% EQU 1 if %vol_off14% EQU 0 (
if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%)
)
exit /b

:sppchkoff
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz4%"
%_qr% > "!_temp!\sppchk.txt"
if %winbuild% LSS 9200 find /i "Office 14" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off14% EQU 0 exit /b)
find /i "Office 15" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off15% EQU 0 exit /b)
find /i "Office 16" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off16% EQU 0 exit /b)
find /i "Office 19" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off19% EQU 0 exit /b)
find /i "Office 21" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off21% EQU 0 exit /b)
if %1 EQU 1 (set _officespp=1) else (set _officespp=0)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%PartialProductKey is not NULL%_zz6% %_zz3% ID %_zz4%"
%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz8%"
for /f "tokens=3 delims==, " %%G in ('%_qr%') do set OffVer=%%G
call :offchk%OffVer%
exit /b

:sppchkwin
set _officespp=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL%_zz6% %_zz3% Name %_zz4%"
if %winbuild% GEQ 14393 if %WinPerm% EQU 0 if %_gvlk% EQU 0 %_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set _gvlk=1)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% LicenseStatus %_zz4%"
%_qr% %_Nul2% | findstr "1" %_Nul1% && (echo.&call :activate&exit /b)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%PartialProductKey is not NULL%_zz6% %_zz3% ID %_zz4%"
%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
if %winbuild% GEQ 14393 if %_gvlk% EQU 1 exit /b
if %WinPerm% EQU 1 exit /b
if %winbuild% LSS 10240 (call :winchk&exit /b)
for %%A in (
b71515d9-89a2-4c60-88c8-656fbcca7f3a,af43f7f0-3b1e-4266-a123-1fdb53f4323b,075aca1f-05d7-42e5-a3ce-e349e7be7078
11a37f09-fb7f-4002-bd84-f3ae71d11e90,43f2ab05-7c87-4d56-b27c-44d0f9a3dabd,2cf5af84-abab-4ff0-83f8-f040fb2576eb
6ae51eeb-c268-4a21-9aae-df74c38b586d,ff808201-fec6-4fd4-ae16-abbddade5706,34260150-69ac-49a3-8a0d-4a403ab55763
4dfd543d-caa6-4f69-a95f-5ddfe2b89567,5fe40dd6-cf1f-4cf2-8729-92121ac2e997,903663f7-d2ab-49c9-8942-14aa9e0a9c72
2cc171ef-db48-4adc-af09-7c574b37f139,5b2add49-b8f4-42e0-a77c-adad4efeeeb1
) do (
if /i '%app%' EQU '%%A' exit /b
)
if not defined EditionID (call :winchk&exit /b)
if %winbuild% LSS 14393 (call :winchk&exit /b)
if /i '%app%' EQU '32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee' if /i %EditionID% NEQ EnterpriseS exit /b
if /i '%app%' EQU 'ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69' if /i %EditionID% NEQ CloudEdition exit /b
if /i '%app%' EQU 'd30136fc-cb4b-416e-a23d-87207abc44a9' if /i %EditionID% NEQ CloudEditionN exit /b
if /i '%app%' EQU '0df4f814-3f57-4b8b-9a9d-fddadcd69fac' if /i %EditionID% NEQ CloudE exit /b
if /i '%app%' EQU 'e0c42288-980c-4788-a014-c080d2e1926e' if /i %EditionID% NEQ Education exit /b
if /i '%app%' EQU '73111121-5638-40f6-bc11-f1d7b0d64300' if /i %EditionID% NEQ Enterprise exit /b
if /i '%app%' EQU '2de67392-b7a7-462a-b1ca-108dd189f588' if /i %EditionID% NEQ Professional exit /b
if /i '%app%' EQU '3f1afc82-f8ac-4f6c-8005-1d233e606eee' if /i %EditionID% NEQ ProfessionalEducation exit /b
if /i '%app%' EQU '82bbc092-bc50-4e16-8e18-b74fc486aec3' if /i %EditionID% NEQ ProfessionalWorkstation exit /b
if /i '%app%' EQU '3c102355-d027-42c6-ad23-2e7ef8a02585' if /i %EditionID% NEQ EducationN exit /b
if /i '%app%' EQU 'e272e3e2-732f-4c65-a8f0-484747d0d947' if /i %EditionID% NEQ EnterpriseN exit /b
if /i '%app%' EQU 'a80b5abf-76ad-428b-b05d-a47d2dffeebf' if /i %EditionID% NEQ ProfessionalN exit /b
if /i '%app%' EQU '5300b18c-2e33-4dc2-8291-47ffcec746dd' if /i %EditionID% NEQ ProfessionalEducationN exit /b
if /i '%app%' EQU '4b1571d3-bafb-4b40-8087-a961be2caf65' if /i %EditionID% NEQ ProfessionalWorkstationN exit /b
if /i '%app%' EQU '58e97c99-f377-4ef1-81d5-4ad5522b5fd8' if /i %EditionID% NEQ Core exit /b
if /i '%app%' EQU 'cd918a57-a41b-4c82-8dce-1a538e221a83' if /i %EditionID% NEQ CoreSingleLanguage exit /b
if /i '%app%' EQU 'ec868e65-fadf-4759-b23e-93fe37f2cc29' if /i %EditionID% NEQ ServerRdsh exit /b
if /i '%app%' EQU 'e4db50ea-bda1-4566-b047-0ca50abc6f07' if /i %EditionID% NEQ ServerRdsh exit /b
set "_qr=%_zz1% %spp% %_zz2% "Description like '%%KMSCLIENT%%'" %_zz3% ID %_zz4%"
if /i "%app%" EQU "e4db50ea-bda1-4566-b047-0ca50abc6f07" (
%_qr% | findstr /i "ec868e65-fadf-4759-b23e-93fe37f2cc29" %_Nul3% && (exit /b)
)
call :winchk
exit /b

:winchk
if not defined tok (if %winbuild% GEQ 9200 (set "tok=4") else (set "tok=7"))
set "_qr=%_zz1% %spp% %_zz2% %_zz5%LicenseStatus='1' and Description like '%%KMSCLIENT%%' %adoff% %_zz6% %_zz3% Name %_zz4%"
%_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && (exit /b)
echo.
set "_qr=%_zz1% %spp% %_zz2% %_zz5%LicenseStatus='1' and GracePeriodRemaining='0' %adoff% and PartialProductKey is not NULL%_zz6% %_zz3% Name %_zz4%"
%_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && (
set WinPerm=1
)
set WinOEM=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and LicenseStatus='1' %adoff% %_zz6% %_zz3% Name %_zz4%"
if %WinPerm% EQU 0 %_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && set WinOEM=1
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and LicenseStatus='1' %adoff% %_zz6% %_zz3% Description %_zz8%"
if %WinOEM% EQU 1 (
for /f "tokens=%tok% delims=, " %%G in ('%_qr%') do set "channel=%%G"
for %%A in (VOLUME_MAK, RETAIL, OEM_DM, OEM_SLP, OEM_COA, OEM_COA_SLP, OEM_COA_NSLP, OEM_NONSLP, OEM) do if /i "%%A"=="!channel!" set WinPerm=1
)
if %WinPerm% EQU 0 (
copy /y %SysPath%\slmgr.vbs "!_temp!\slmgr.vbs" %_Nul3%
cscript //nologo "!_temp!\slmgr.vbs" /xpr %_Nul2% | findstr /i "permanently" %_Nul3% && set WinPerm=1
)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and LicenseStatus='1' %adoff% %_zz6% %_zz3% Name %_zz8%"
if %WinPerm% EQU 1 (
for /f "tokens=2 delims==" %%x in ('%_qr%') do echo Checking: %%x
echo Product is Permanently Activated.
exit /b
)
call :insKey
exit /b

:esuchk
set _officespp=0
set ESU_ADD=1
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% LicenseStatus %_zz4%"
%_qr% %_Nul2% | findstr "1" %_Nul1% && (echo.&call :activate&exit /b)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='77db037b-95c3-48d7-a3ab-a9c6d41093e0'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "3fcc2df2-f625-428d-909a-1f76efc849b6" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='0e00c25d-8795-4fb7-9572-3803d91b6880'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "dadfcd24-6e37-47be-8f7f-4ceda614cece" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='4220f546-f522-46df-8202-4d07afd26454'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "0c29c85e-12d7-4af8-8e4d-ca1e424c480c" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='553673ed-6ddf-419c-a153-b760283472fd'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "f2b21bfc-a6b0-4413-b4bb-9f06b55f2812" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='04fa0286-fa74-401e-bbe9-fbfbb158010d'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "bfc078d0-8c7f-475c-8519-accc46773113" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='16c08c85-0c8b-4009-9b2b-f1f7319e45f9'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "23c6188f-c9d8-457e-81b6-adb6dacb8779" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='8e7bfb1e-acc1-4f56-abae-b80fce56cd4b'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "e7cce015-33d6-41c1-9831-022ba63fe1da" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%PartialProductKey is not NULL%_zz6% %_zz3% ID %_zz4%"
%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
call :insKey
exit /b

:RunOSPP
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
set Off1ce=0
set RanR2V=0
set aC2R21=0
set aC2R19=0
set aC2R16=0
set aC2R15=0
if %winbuild% LSS 9200 (set "aword=2010-2021") else (set "aword=2010")
if %OsppHook% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b)
if %winbuild% GEQ 9200 if %loc_off14% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b)
set err_offsvc=0
net start osppsvc /y %_Nul3% || (
sc start osppsvc %_Nul3%
if !errorlevel! EQU 1053 set err_offsvc=1
)
if %err_offsvc% EQU 1 (echo.&echo Error: osppsvc service is not running...&exit /b)
if %winbuild% GEQ 9200 call :oppoff
if %winbuild% LSS 9200 call :sppoff
if %Off1ce% EQU 0 exit /b
set "vPrem="&set "vProf="
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='OfficeVisioPrem-MAK'%_zz6% %_zz3% LicenseStatus %_zz8%"
if %loc_off14% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set vPrem=%%A
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='OfficeVisioPro-MAK'%_zz6% %_zz3% LicenseStatus %_zz8%"
if %loc_off14% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set vProf=%%A
set "_qr=%_zz7% %sps% %_zz3% Version %_zz8%"
for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set slsv=%%A
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
set "_qr=%_zz7% %spp% %_zz2% %_zz5%Description like '%%KMSCLIENT%%' %_zz6% %_zz3% ID %_zz8%"
for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkoff 2)
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing %_Null%
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching %_Null%
exit /b

:oppoff
set "_qr=%_zz1% %spp% %_zz3% Description %_zz4%"
%_qr% %_Nul2% | findstr /i KMSCLIENT %_Nul1% && (
set Off1ce=1
exit /b
)
set ret_off14=0
%_qr% %_Nul2% | findstr /i channel %_Nul1% && (set ret_off14=1)
if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%)
exit /b

:offchk
set ls=0
set ls2=0
set ls3=0
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='Office%~1'%_zz6% %_zz3% LicenseStatus %_zz8%"
for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls=%%A
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='Office%~3'%_zz6% %_zz3% LicenseStatus %_zz8%"
if /i not "%~3"=="" for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls2=%%A
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='Office%~5'%_zz6% %_zz3% LicenseStatus %_zz8%"
if /i not "%~5"=="" for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls3=%%A
if "%ls3%"=="1" (
echo Checking: %~6
echo Product is Permanently Activated.
exit /b
)
if "%ls2%"=="1" (
echo Checking: %~4
echo Product is Permanently Activated.
exit /b
)
if "%ls%"=="1" (
echo Checking: %~2
echo Product is Permanently Activated.
exit /b
)
call :insKey
exit /b

:offchk21
if /i '%app%' EQU 'f3fb2d68-83dd-4c8b-8f09-08e0d950ac3b' exit /b
if /i '%app%' EQU '76093b1b-7057-49d7-b970-638ebcbfd873' exit /b
if /i '%app%' EQU 'a3b44174-2451-4cd6-b25f-66638bfb9046' exit /b
if /i '%app%' EQU 'fbdb3e18-a8ef-4fb3-9183-dffd60bd0984' (
call :offchk "21ProPlus2021VL_MAK_AE1" "Office ProPlus 2021" "21ProPlus2021VL_MAK_AE2"
exit /b
)
if /i '%app%' EQU '080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3' (
call :offchk "21Standard2021VL_MAK_AE" "Office Standard 2021"
exit /b
)
if /i '%app%' EQU '76881159-155c-43e0-9db7-2d70a9a3a4ca' (
call :offchk "21ProjectPro2021VL_MAK_AE1" "Project Pro 2021" "21ProjectPro2021VL_MAK_AE2"
exit /b
)
if /i '%app%' EQU '6dd72704-f752-4b71-94c7-11cec6bfc355' (
call :offchk "21ProjectStd2021VL_MAK_AE" "Project Standard 2021"
exit /b
)
if /i '%app%' EQU 'fb61ac9a-1688-45d2-8f6b-0674dbffa33c' (
call :offchk "21VisioPro2021VL_MAK_AE" "Visio Pro 2021"
exit /b
)
if /i '%app%' EQU '72fce797-1884-48dd-a860-b2f6a5efd3ca' (
call :offchk "21VisioStd2021VL_MAK_AE" "Visio Standard 2021"
exit /b
)
call :insKey
exit /b

:offchk19
if /i '%app%' EQU '0bc88885-718c-491d-921f-6f214349e79c' exit /b
if /i '%app%' EQU 'fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9' exit /b
if /i '%app%' EQU '500f6619-ef93-4b75-bcb4-82819998a3ca' exit /b
if /i '%app%' EQU '85dd8b5f-eaa4-4af3-a628-cce9e77c9a03' (
call :offchk "19ProPlus2019VL_MAK_AE" "Office ProPlus 2019"
exit /b
)
if /i '%app%' EQU '6912a74b-a5fb-401a-bfdb-2e3ab46f4b02' (
call :offchk "19Standard2019VL_MAK_AE" "Office Standard 2019"
exit /b
)
if /i '%app%' EQU '2ca2bf3f-949e-446a-82c7-e25a15ec78c4' (
call :offchk "19ProjectPro2019VL_MAK_AE" "Project Pro 2019"
exit /b
)
if /i '%app%' EQU '1777f0e3-7392-4198-97ea-8ae4de6f6381' (
call :offchk "19ProjectStd2019VL_MAK_AE" "Project Standard 2019"
exit /b
)
if /i '%app%' EQU '5b5cf08f-b81a-431d-b080-3450d8620565' (
call :offchk "19VisioPro2019VL_MAK_AE" "Visio Pro 2019"
exit /b
)
if /i '%app%' EQU 'e06d7df3-aad0-419d-8dfb-0ac37e2bdf39' (
call :offchk "19VisioStd2019VL_MAK_AE" "Visio Standard 2019"
exit /b
)
call :insKey
exit /b

:offchk16
if /i '%app%' EQU 'd450596f-894d-49e0-966a-fd39ed4c4c64' (
call :offchk "16ProPlusVL_MAK" "Office ProPlus 2016"
exit /b
)
if /i '%app%' EQU 'dedfa23d-6ed1-45a6-85dc-63cae0546de6' (
call :offchk "16StandardVL_MAK" "Office Standard 2016"
exit /b
)
if /i '%app%' EQU '4f414197-0fc2-4c01-b68a-86cbb9ac254c' (
call :offchk "16ProjectProVL_MAK" "Project Pro 2016"
exit /b
)
if /i '%app%' EQU 'da7ddabc-3fbe-4447-9e01-6ab7440b4cd4' (
call :offchk "16ProjectStdVL_MAK" "Project Standard 2016"
exit /b
)
if /i '%app%' EQU '6bf301c1-b94a-43e9-ba31-d494598c47fb' (
call :offchk "16VisioProVL_MAK" "Visio Pro 2016"
exit /b
)
if /i '%app%' EQU 'aa2a7821-1827-4c2c-8f1d-4513a34dda97' (
call :offchk "16VisioStdVL_MAK" "Visio Standard 2016"
exit /b
)
if /i '%app%' EQU '829b8110-0e6f-4349-bca4-42803577788d' (
call :offchk "16ProjectProXC2RVL_MAKC2R" "Project Pro 2016 C2R"
exit /b
)
if /i '%app%' EQU 'cbbaca45-556a-4416-ad03-bda598eaa7c8' (
call :offchk "16ProjectStdXC2RVL_MAKC2R" "Project Standard 2016 C2R"
exit /b
)
if /i '%app%' EQU 'b234abe3-0857-4f9c-b05a-4dc314f85557' (
call :offchk "16VisioProXC2RVL_MAKC2R" "Visio Pro 2016 C2R"
exit /b
)
if /i '%app%' EQU '361fe620-64f4-41b5-ba77-84f8e079b1f7' (
call :offchk "16VisioStdXC2RVL_MAKC2R" "Visio Standard 2016 C2R"
exit /b
)
call :insKey
exit /b

:offchk15
if /i '%app%' EQU 'b322da9c-a2e2-4058-9e4e-f59a6970bd69' (
call :offchk "ProPlusVL_MAK" "Office ProPlus 2013"
exit /b
)
if /i '%app%' EQU 'b13afb38-cd79-4ae5-9f7f-eed058d750ca' (
call :offchk "StandardVL_MAK" "Office Standard 2013"
exit /b
)
if /i '%app%' EQU '4a5d124a-e620-44ba-b6ff-658961b33b9a' (
call :offchk "ProjectProVL_MAK" "Project Pro 2013"
exit /b
)
if /i '%app%' EQU '427a28d1-d17c-4abf-b717-32c780ba6f07' (
call :offchk "ProjectStdVL_MAK" "Project Standard 2013"
exit /b
)
if /i '%app%' EQU 'e13ac10e-75d0-4aff-a0cd-764982cf541c' (
call :offchk "VisioProVL_MAK" "Visio Pro 2013"
exit /b
)
if /i '%app%' EQU 'ac4efaf0-f81f-4f61-bdf7-ea32b02ab117' (
call :offchk "VisioStdVL_MAK" "Visio Standard 2013"
exit /b
)
call :insKey
exit /b

:offchk14
if /i '%app%' EQU '6f327760-8c5c-417c-9b61-836a98287e0c' (
call :offchk "ProPlus-MAK" "Office ProPlus 2010" "ProPlusAcad-MAK" "Office Professional Academic 2010"
exit /b
)
if /i '%app%' EQU '9da2a678-fb6b-4e67-ab84-60dd6a9c819a' (
call :offchk "Standard-MAK" "Office Standard 2010" "StandardAcad-MAK"  "Office Standard Academic 2010"
exit /b
)
if /i '%app%' EQU 'ea509e87-07a1-4a45-9edc-eba5a39f36af' (
call :offchk "SmallBusBasics-MAK" "Office Small Business Basics 2010"
exit /b
)
if /i '%app%' EQU 'df133ff7-bf14-4f95-afe3-7b48e7e331ef' (
call :offchk "ProjectPro-MAK" "Project Pro 2010"
exit /b
)
if /i '%app%' EQU '5dc7bf61-5ec9-4996-9ccb-df806a2d0efe' (
call :offchk "ProjectStd-MAK" "Project Standard 2010" "ProjectStd-MAK2" "Project Standard 2010"
exit /b
)
if /i '%app%' EQU '92236105-bb67-494f-94c7-7f7a607929bd' (
call :offchk "VisioPrem-MAK" "Visio Premium 2010" "VisioPro-MAK" "Visio Pro 2010"
exit /b
)
if defined vPrem exit /b
if /i '%app%' EQU 'e558389c-83c3-4b29-adfe-5e4d7f46c358' (
call :offchk "VisioPro-MAK" "Visio Pro 2010" "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
if defined vProf exit /b
if /i '%app%' EQU '9ed833ff-4f92-4f36-b370-8683a4f13275' (
call :offchk "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
call :insKey
exit /b

:officeLoc
set loc_off%1=0
set _O%1MSI=0
if %1 EQU 19 (
if defined _C16R reg query %_C16R% /v ProductReleaseIds %_Nul2% | findstr 2019 %_Nul1% && set loc_off%1=1
exit /b
)
if %1 EQU 21 (
if defined _C16R reg query %_C16R% /v ProductReleaseIds %_Nul2% | findstr 2021 %_Nul1% && set loc_off%1=1
exit /b
)

for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" (
set loc_off%1=1
set _O%1MSI=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" (
set loc_off%1=1
set _O%1MSI=1
)

if %1 EQU 16 if defined _C16R (
for /f "skip=2 tokens=2*" %%a in ('reg query %_C16R% /v ProductReleaseIds') do echo %%b> "!_temp!\c2rchk.txt"
for %%a in (%LV16Ids%,ProjectProX,ProjectStdX,VisioProX,VisioStdX) do (
  findstr /I /C:"%%aVolume" "!_temp!\c2rchk.txt" %_Nul1% && set loc_off%1=1
  )
for %%a in (%LR16Ids%) do (
  findstr /I /C:"%%aRetail" "!_temp!\c2rchk.txt" %_Nul1% && set loc_off%1=1
  )
exit /b
)

if %1 EQU 15 if defined _C15R (
set loc_off%1=1
exit /b
)

if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if not %xOS%==x86 if exist "%ProgramW6432%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if not %xOS%==x86 if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
exit /b

:officeSub
reg query %kNext% | findstr /i /r ".*retail" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x2" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*retail" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x3" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*volume" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x2" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*volume" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x3" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r "project.*" %_Nul2% | find /i "0x2" %_Nul1% && set sub_proj=1
reg query %kNext% | findstr /i /r "project.*" %_Nul2% | find /i "0x3" %_Nul1% && set sub_proj=1
reg query %kNext% | findstr /i /r "visio.*" %_Nul2% | find /i "0x2" %_Nul1% && set sub_vsio=1
reg query %kNext% | findstr /i /r "visio.*" %_Nul2% | find /i "0x3" %_Nul1% && set sub_vsio=1
if %sub_o365% EQU 1 set sub_next=1
if %sub_proj% EQU 1 set sub_next=1
if %sub_vsio% EQU 1 set sub_next=1
exit /b

:insKey
set S_OK=1
echo.
set "_key="
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz8%"
if %ESU_ADD% EQU 0 for /f "tokens=2 delims==" %%x in ('%_qr%') do echo Installing Key: %%x
if %ESU_ADD% EQU 1 for /f "tokens=2 delims==f" %%x in ('%_qr%') do echo Installing Key: %%x
set ESU_ADD=0
call :keys %app%
if "%_key%"=="" (echo No associated KMS Client key found&exit /b)
set "_qr=wmic path %sps% where Version='%slsv%' call InstallProductKey ProductKey="%_key%""
if %WMI_VBS% NEQ 0 set "_qr=%_csp% %sps% "%_key%""
%_qr% %_Nul3%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 (
cmd /c exit /b %ERRORCODE%
echo Failed: 0x!=ExitCode!
set S_OK=0
exit /b
)
set "_qr=wmic path %sps% where Version='%slsv%' call RefreshLicenseStatus"
if %WMI_VBS% NEQ 0 set "_qr=%_csm% "%sps%.Version='%slsv%'" RefreshLicenseStatus"
if %sps% EQU SoftwareLicensingService %_qr% %_Nul3%

:activate
set S_OK=1
if %sps% EQU SoftwareLicensingService (
if %_officespp% EQU 0 (reg delete "HKLM\%SPPk%\%_wApp%\%app%" /f %_Null%) else (reg delete "HKLM\%SPPk%\%_oApp%\%app%" /f %_Null%)
) else (
reg delete "HKLM\%OPPk%\%_oA14%\%app%" /f %_Null%
reg delete "HKLM\%OPPk%\%_oApp%\%app%" /f %_Null%
)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz8%"
if %W1nd0ws% EQU 0 if %_officespp% EQU 0 if %sps% EQU SoftwareLicensingService (
reg add "HKLM\%SPPk%\%_wApp%\%app%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2" %_Nul3%
reg add "HKLM\%SPPk%\%_wApp%\%app%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
for /f "tokens=2 delims==" %%x in ('%_qr%') do echo Checking: %%x
echo Product is KMS 2038 Activated.
set _keepkms38=1
exit /b
)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz8%"
if %act_attempt% LSS 1 (
if %ESU_ADD% EQU 0 for /f "tokens=2 delims==" %%x in ('%_qr%') do echo Activating: %%x
if %ESU_ADD% EQU 1 for /f "tokens=2 delims==f" %%x in ('%_qr%') do echo Activating: %%x
)

set ESU_ADD=0
set "_qr=wmic path %spp% where ID='%app%' call Activate"
if %WMI_VBS% NEQ 0 set "_qr=%_csm% "%spp%.ID='%app%'" Activate"
%_qr% %_Nul3%
call set ERRORCODE=%ERRORLEVEL%
if %act_attempt% LSS 1 if %ERRORCODE% EQU -1073418187 (
echo Product Activation Failed: 0xC004F035
if %OSType% EQU Win7 echo Windows 7 cannot be KMS-activated on this computer due to unqualified OEM BIOS.
echo See Read Me for details.
exit /b
)
if %act_attempt% LSS 1 if %ERRORCODE% EQU -1073417728 (
echo Product Activation Failed: 0xC004F200
echo Windows needs to rebuild the activation-related files.
echo See KB2736303 for details.
exit /b
)
if %act_attempt% LSS 1 if %ERRORCODE% EQU -1073422315 (
echo Product Activation Failed: 0xC004E015
echo Running slmgr.vbs /rilc to mitigate.
cscript //Nologo //B %SysPath%\slmgr.vbs /rilc
)
set gpr=0
set gpr2=0
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% GracePeriodRemaining %_zz8%"
for /f "tokens=2 delims==" %%x in ('%_qr%') do (set gpr=%%x&set /a "gpr2=(%%x+1440-1)/1440")
if %act_attempt% LSS 1 if %ERRORCODE% EQU 0 if %gpr% EQU 0 (
echo Product Activation succeeded, but Remaining Period failed to increase.
if %OSType% EQU Win7 echo This could be related to the error described in KB4487266
exit /b
)
set Act_OK=0
if %gpr% EQU 43200 if %_officespp% EQU 0 if %winbuild% GEQ 9200 set Act_OK=1
if %gpr% EQU 64800 set Act_OK=1
if %gpr% GTR 259200 if %Win10Gov% EQU 1 set Act_OK=1
if %gpr% EQU 259200 set Act_OK=1

if %ERRORCODE% EQU 0 if %Act_OK% EQU 1 (
call :_color %_Green% "Product Activation Successful"
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
set /a act_attempt=0
exit /b
)

if not !server_num! gtr %max_servers% (
if %act_attempt% LSS 3 (
set /a act_attempt+=1
call :getserv
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!"
%nul% reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!"
if %winbuild% GEQ 9200 (
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!"
if defined notx86 (
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!" /reg:32
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!" /reg:32
)
)
goto :activate
)
)

cmd /c exit /b %ERRORCODE%
if %ERRORCODE% NEQ 0 (
call :_color %_Red% "Product Activation Failed: 0x!=ExitCode!"
) else (
call :_color %_Red% "Product Activation Failed"
)
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
set S_OK=0
set act_failed=1
set /a act_attempt=0
exit /b

:StopService
sc query %1 | find /i "STOPPED" %_Nul1% || net stop %1 /y %_Nul3%
sc query %1 | find /i "STOPPED" %_Nul1% || sc stop %1 %_Nul3%
goto :eof

:UpdateOSPPEntry
if /i %1 EQU osppsvc.exe (
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!" %_Nul3%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
)
goto :eof

:CheckFR

set WMIe=0
call :CheckWS
if %WMIe% EQU 1 (
echo.
echo %_err%
echo Failed running WMI query check.
)
goto :eof

:CheckWS
set "_qrw=%_zz1% Win32_ComputerSystem %_zz3% CreationClassName %_zz4%"
set "_qrs=%_zz1% SoftwareLicensingService %_zz3% Version %_zz4%"

%_qrs% %_Nul2% | findstr /r "[0-9]*\.[0-9]*\.[0-9]*\.[0-9]*" %_Nul1% || (
  set WMIe=1
  %_qrw% %_Nul2% | find /i "ComputerSystem" %_Nul1% && (
    echo Error: SPP is not responding
    ) || (
    echo Error: WMI ^& SPP are not responding
  )
)
goto :eof

:C2RR2V
set RanR2V=1
set "_SLMGR=%SysPath%\slmgr.vbs"
if %_Debug% EQU 0 (
set "_cscript=cscript //Nologo //B"
) else (
set "_cscript=cscript //Nologo"
)
set _LTSC=0
set "_tag="&set "_ons= 2016"
sc query ClickToRunSvc %_Nul3%
set error1=%errorlevel%
sc query OfficeSvc %_Nul3%
set error2=%errorlevel%
if %error1% EQU 1060 if %error2% EQU 1060 (
echo Error: Office C2R service is not detected
goto :%_fC2R%
)
set _Office16=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
set _Office15=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
if %_Office16% EQU 0 if %_Office15% EQU 0 (
echo Error: Office C2R InstallPath is not detected
goto :%_fC2R%
)

:Reg16istry
if %_Office16% EQU 0 goto :Reg15istry
set "_InstallRoot="
set "_ProductIds="
set "_GUID="
set "_Config="
set "_PRIDs="
set "_LicensesPath="
set "_Integrator="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do (set "_InstallRoot=%%b\root")
if not "%_InstallRoot%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do (set "_OSPPVBS=%%b\Office16\OSPP.VBS")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v PackageGUID" %_Nul6%') do (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
  set "_PRIDs=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do (set "_InstallRoot=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do (set "_OSPPVBS=%%b\Office16\OSPP.VBS")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v PackageGUID" %_Nul6%') do (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
  set "_PRIDs=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\ProductReleaseIDs"
)
set "_LicensesPath=%_InstallRoot%\Licenses16"
set "_Integrator=%_InstallRoot%\integration\integrator.exe"
for /f "skip=2 tokens=2*" %%a in ('"reg query %_PRIDs% /v ActiveConfiguration" %_Nul6%') do set "_PRIDs=%_PRIDs%\%%b"
if "%_ProductIds%"=="" (
if %_Office15% EQU 0 (echo Error: Office C2R ProductIDs are not detected&goto :%_fC2R%) else (goto :Reg15istry)
)
if not exist "%_LicensesPath%\ProPlus*.xrm-ms" (
if %_Office15% EQU 0 (echo Error: Office C2R Licenses files are not detected&goto :%_fC2R%) else (goto :Reg15istry)
)
if not exist "%_Integrator%" (
if %_Office15% EQU 0 (echo Error: Office C2R Licenses Integrator is not detected&goto :%_fC2R%) else (goto :Reg15istry)
)
if exist "%_LicensesPath%\Word2019VL_KMS_Client_AE*.xrm-ms" (set "_tag=2019"&set "_ons= 2019")
if exist "%_LicensesPath%\Word2021VL_KMS_Client_AE*.xrm-ms" (set _LTSC=1)
if %winbuild% LSS 10240 if !_LTSC! EQU 1 (set "_tag=2021"&set "_ons= 2021")
if %_Office15% EQU 0 goto :CheckC2R

:Reg15istry
set "_Install15Root="
set "_Product15Ids="
set "_Con15fig="
set "_PR15IDs="
set "_OSPP15Ready="
set "_Licenses15Path="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do (set "_Install15Root=%%b\root")
if not "%_Install15Root%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do (set "_Install15Root=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration"
)
set "_OSPP15ReadT=REG_SZ"
if "%_Product15Ids%"=="" (
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
  )
reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
  )
)
set "_Licenses15Path=%_Install15Root%\Licenses"
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS"
)
if "%_Product15Ids%"=="" (
if %_Office16% EQU 0 (echo Error: Office 2013 C2R ProductIDs are not detected&goto :%_fC2R%) else (goto :CheckC2R)
)
if not exist "%_Licenses15Path%\ProPlus*.xrm-ms" (
if %_Office16% EQU 0 (echo Error: Office 2013 C2R Licenses files are not detected&goto :%_fC2R%) else (goto :CheckC2R)
)
if %winbuild% LSS 9200 if not exist "%_OSPP15VBS%" (
if %_Office16% EQU 0 (echo Error: Office 2013 C2R Licensing tool OSPP.vbs is not detected&goto :%_fC2R%) else (goto :CheckC2R)
)

:CheckC2R
set _OMSI=0
if %_Office16% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
)
if %_Office15% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
)
if %winbuild% GEQ 9200 (
set _spp=SoftwareLicensingProduct
set _sps=SoftwareLicensingService
set "_vbsi=%_SLMGR% /ilc "
set "_vbsf=%_SLMGR% /ilc "
) else (
set _spp=OfficeSoftwareProtectionProduct
set _sps=OfficeSoftwareProtectionService
set _vbsi="!_OSPP15VBS!" /inslic:
set _vbsf="!_OSPPVBS!" /inslic:
)
set "_wmi="
set "_qr=%_zz7% %_sps% %_zz3% Version %_zz8%"
for /f "tokens=2 delims==" %%# in ('%_qr%') do set _wmi=%%#
if "%_wmi%"=="" (
echo Error: %_sps% WMI version is not detected
call :CheckWS
goto :%_fC2R%
)
set _Retail=0
set "_ocq=ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey is not NULL"
if %WMI_VBS% EQU 0 wmic path %_spp% where (%_ocq%) get Description %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
set "_qr=%_csq% %_spp% "%_ocq%" Description"
if %WMI_VBS% NEQ 0 %_qr% %_Nul2% >"!_temp!\crvRetail.txt"
find /i "RETAIL channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "RETAIL(MAK) channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "TIMEBASED_SUB channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
set rancopp=0
if %_Retail% EQU 0 if %_OMSI% EQU 0 (
set rancopp=1
%_Nul3% powershell "$f=[io.file]::ReadAllText('!_batp!') -split ':cleanlicense\:.*';iex ($f[1]);"
)
set _O16O365=0
set _C16Msg=0
set _C15Msg=0
set "_qr=%_csq% %_spp% "%_ocq%" LicenseFamily"
if %_Retail% EQU 1 if %WMI_VBS% EQU 0 wmic path %_spp% where (%_ocq%) get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
if %_Retail% EQU 1 if %WMI_VBS% NEQ 0 %_qr% %_Nul2% >"!_temp!\crvRetail.txt"
set "_qr=%_csq% %_spp% "ApplicationID='%_oApp%'" LicenseFamily"
if %WMI_VBS% EQU 0 wmic path %_spp% where "ApplicationID='%_oApp%'" get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvVolume.txt" 2>&1
if %WMI_VBS% NEQ 0 %_qr% %_Nul2% >"!_temp!\crvVolume.txt" 2>&1

if %_Office16% EQU 0 goto :R15V

set _O21Ids=ProPlus2021,ProjectPro2021,VisioPro2021,Standard2021,ProjectStd2021,VisioStd2021,Access2021,SkypeforBusiness2021
set _O19Ids=ProPlus2019,ProjectPro2019,VisioPro2019,Standard2019,ProjectStd2019,VisioStd2019,Access2019,SkypeforBusiness2019
set _O16Ids=ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness
set _A21Ids=Excel2021,Outlook2021,PowerPoint2021,Publisher2021,Word2021
set _A19Ids=Excel2019,Outlook2019,PowerPoint2019,Publisher2019,Word2019
set _A16Ids=Excel,Outlook,PowerPoint,Publisher,Word
set _V21Ids=%_O21Ids%,%_A21Ids%
set _V19Ids=%_O19Ids%,%_A19Ids%
set _V16Ids=Mondo,%_O16Ids%,%_A16Ids%,OneNote
set _R16Ids=%_V16Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set _RetIds=%_V21Ids%,Professional2021,HomeBusiness2021,HomeStudent2021,%_V19Ids%,Professional2019,HomeBusiness2019,HomeStudent2019,%_R16Ids%
set _Suites=Mondo,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud,ProPlus,Standard,Professional,HomeBusiness,HomeStudent,ProPlus2019,Standard2019,Professional2019,HomeBusiness2019,HomeStudent2019,ProPlus2021,Standard2021,Professional2021,HomeBusiness2021,HomeStudent2021
set _PrjSKU=ProjectPro,ProjectStd,ProjectPro2019,ProjectStd2019,ProjectPro2021,ProjectStd2021
set _VisSKU=VisioPro,VisioStd,VisioPro2019,VisioStd2019,VisioPro2021,VisioStd2021

echo %_ProductIds%>"!_temp!\crvProductIds.txt"
for %%a in (%_RetIds%,ProPlus) do (
set _%%a=0
)
for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && set _%%a=1
)
if !_LTSC! EQU 0 for %%a in (%_V21Ids%) do (
set _%%a=0
)
if !_LTSC! EQU 1 for %%a in (%_V21Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office21%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
for %%a in (%_V19Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office19%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
for %%a in (%_V16Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aE5R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aEDUR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aXC2RVL_MAKC2R" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office19%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R19=1)
  find /i "Office19%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R19=1)
  find /i "Office19%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R19=1)
  find /i "Office19%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R19=1)
  find /i "Office21%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R21=1)
  find /i "Office21%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R21=1)
  find /i "Office21%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R21=1)
  find /i "Office21%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R21=1)
  )
)
if %_Retail% EQU 1 reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R16=1)
  find /i "Office16ProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R16=1)
  find /i "Office16ProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R16=1)
  find /i "Office16ProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R16=1)
)
set "_qr=%_zz1% %_spp% %_zz2% "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%'" %_zz3% LicenseFamily %_zz4%"
find /i "Office16MondoVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% && (
  for %%a in (O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud) do set _%%a=0
  )
)
if %sub_o365% EQU 1 (
  for %%a in (%_Suites%) do set _%%a=0
echo.
echo Microsoft Office is activated with a vNext license.
)
if %sub_proj% EQU 1 (
  for %%a in (%_PrjSKU%) do set _%%a=0
echo.
echo Microsoft Project is activated with a vNext license.
)
if %sub_vsio% EQU 1 (
  for %%a in (%_VisSKU%) do set _%%a=0
echo.
echo Microsoft Visio is activated with a vNext license.
)

for %%a in (%_RetIds%,ProPlus) do if !_%%a! EQU 1 (
set _C16Msg=1
)
if %_C16Msg% EQU 1 (
echo.
echo Converting Office C2R Retail-to-Volume:
)
if %_C16Msg% EQU 0 (if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R))

for %%# in ("!_LicensesPath!\client-issuance-*.xrm-ms") do (
%_cscript% %_vbsf%"!_LicensesPath!\%%~nx#"
)
%_cscript% %_vbsf%"!_LicensesPath!\pkeyconfig-office.xrm-ms"

if !_Mondo! EQU 1 (
call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 (
echo O365ProPlus 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365Business! EQU 1 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365Business 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365Business NCHRJ-3VPGW-X73DM-6B36K-3RQ6B
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365HomePrem! EQU 1 if !_O365SmallBusPrem! EQU 0 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365HomePrem 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365EduCloud! EQU 1 if !_O365HomePrem! EQU 0 if !_O365SmallBusPrem! EQU 0 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365EduCloud 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365EduCloud 8843N-BCXXD-Q84H8-R4Q37-T3CPT
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 set _O16O365=1
if !_Mondo! EQU 1 if !_O365ProPlus! EQU 0 (
echo Mondo 2016 Suite
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R)
)
if !_ProPlus2021! EQU 1 if !_O365ProPlus! EQU 0 (
echo ProPlus 2021 Suite
call :InsLic ProPlus2021
)
if !_ProPlus2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 (
echo ProPlus 2019 Suite -^> ProPlus%_ons% Licenses
call :InsLic ProPlus%_tag%
)
if !_ProPlus! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 (
echo ProPlus 2016 Suite -^> ProPlus%_ons% Licenses
call :InsLic ProPlus%_tag%
)
if !_Professional2021! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 (
echo Professional 2021 Suite -^> ProPlus 2021 Licenses
call :InsLic ProPlus2021
)
if !_Professional2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 (
echo Professional 2019 Suite -^> ProPlus%_ons% Licenses
call :InsLic ProPlus%_tag%
)
if !_Professional! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 (
echo Professional 2016 Suite -^> ProPlus%_ons% Licenses
call :InsLic ProPlus%_tag%
)
if !_Standard2021! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 (
echo Standard 2021 Suite
call :InsLic Standard2021
)
if !_Standard2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 (
echo Standard 2019 Suite -^> Standard%_ons% Licenses
call :InsLic Standard%_tag%
)
if !_Standard! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 (
echo Standard 2016 Suite -^> Standard%_ons% Licenses
call :InsLic Standard%_tag%
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a2021! EQU 1 (
  echo %%a 2021 SKU
  call :InsLic %%a2021
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a2019! EQU 1 (
if !_%%a2021! EQU 0 (
  echo %%a 2019 SKU -^> %%a%_ons% Licenses
  call :InsLic %%a%_tag%
  )
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! EQU 1 (
if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 SKU -^> %%a%_ons% Licenses
  call :InsLic %%a%_tag%
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a2021! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 (
  set _Standard2021=1
  echo %%a 2021 Suite -^> Standard 2021 Licenses
  call :InsLic Standard2021
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a2019! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2021! EQU 0 (
  set _Standard2019=1
  echo %%a 2019 Suite -^> Standard%_ons% Licenses
  call :InsLic Standard%_tag%
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  set _Standard=1
  echo %%a 2016 Suite -^> Standard%_ons% Licenses
  call :InsLic Standard%_tag%
  )
)
for %%a in (%_A21Ids%,OneNote) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 (
  echo %%a App
  call :InsLic %%a
  )
)
for %%a in (%_A16Ids%) do if !_%%a2019! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2021! EQU 0 (
  echo %%a 2019 App -^> %%a%_ons% Licenses
  call :InsLic %%a%_tag%
  )
)
for %%a in (%_A16Ids%) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App -^> %%a%_ons% Licenses
  call :InsLic %%a%_tag%
  )
)
for %%a in (Access) do if !_%%a2021! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 (
  echo %%a 2021 App
  call :InsLic %%a2021
  )
)
for %%a in (Access) do if !_%%a2019! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_%%a2021! EQU 0 (
  echo %%a 2019 App -^> %%a%_ons% Licenses
  call :InsLic %%a%_tag%
  )
)
for %%a in (Access) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App -^> %%a%_ons% Licenses
  call :InsLic %%a%_tag%
  )
)
for %%a in (SkypeforBusiness) do if !_%%a2021! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 (
  echo %%a 2021 App
  call :InsLic %%a2021
  )
)
for %%a in (SkypeforBusiness) do if !_%%a2019! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_%%a2021! EQU 0 (
  echo %%a 2019 App -^> %%a%_ons% Licenses
  call :InsLic %%a%_tag%
  )
)
for %%a in (SkypeforBusiness) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App -^> %%a%_ons% Licenses
  call :InsLic %%a%_tag%
  )
)
if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R)

:R15V
set _O15Ids=Standard,ProjectPro,VisioPro,ProjectStd,VisioStd,Access,Lync
set _A15Ids=Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set _R15Ids=SPD,Mondo,%_O15Ids%,%_A15Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem
set _V15Ids=Mondo,%_O15Ids%,%_A15Ids%

echo %_Product15Ids%>"!_temp!\crvProduct15s.txt"
for %%a in (%_R15Ids%,ProPlus) do (
set _%%a=0
)
for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V15Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PR15IDs%\Active\ProPlusVolume\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  )
)
if %_Retail% EQU 1 reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R15=1)
  find /i "OfficeProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R15=1)
  find /i "OfficeProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R15=1)
  find /i "OfficeProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R15=1)
)
set "_qr=%_zz1% %_spp% %_zz2% "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%'" %_zz3% LicenseFamily %_zz4%"
find /i "OfficeMondoVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% && (
  for %%a in (O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem) do set _%%a=0
  )
)

for %%a in (%_R15Ids%,ProPlus) do if !_%%a! EQU 1 (
set _C15Msg=1
)
if %_C15Msg% EQU 1 if %_C16Msg% EQU 0 (
echo.
echo Converting Office C2R Retail-to-Volume:
)
if %_C15Msg% EQU 0 goto :GVLKC2R

for %%# in ("!_Licenses15Path!\client-issuance-*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
%_cscript% %_vbsi%"!_Licenses15Path!\pkeyconfig-office.xrm-ms"

if !_Mondo! EQU 1 (
call :Ins15Lic Mondo
)
if !_O365ProPlus! EQU 1 if !_O16O365! EQU 0 (
echo O365ProPlus 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365HomePrem! EQU 1 if !_O365SmallBusPrem! EQU 0 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365HomePrem 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365Business! EQU 1 if !_O365HomePrem! EQU 0 if !_O365SmallBusPrem! EQU 0 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365Business 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365Business MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_Mondo! EQU 1 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
echo Mondo 2013 Suite
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :GVLKC2R
)
if !_SPD! EQU 1 if !_Mondo! EQU 0 if !_O365ProPlus! EQU 0 (
echo SharePoint Designer 2013 App -^> Mondo 2013 Licenses
call :Ins15Lic Mondo
goto :GVLKC2R
)
if !_ProPlus! EQU 1 if !_O365ProPlus! EQU 0 (
echo ProPlus 2013 Suite
call :Ins15Lic ProPlus
)
if !_Professional! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 (
echo Professional 2013 Suite -^> ProPlus 2013 Licenses
call :Ins15Lic ProPlus
)
if !_Standard! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 (
echo Standard 2013 Suite
call :Ins15Lic Standard
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! EQU 1 (
echo %%a 2013 SKU
call :Ins15Lic %%a
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 if !_Standard! EQU 0 (
  set _Standard=1
  echo %%a 2013 Suite -^> Standard 2013 Licenses
  call :Ins15Lic Standard
  )
)
for %%a in (%_A15Ids%) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 if !_Standard! EQU 0 (
  echo %%a 2013 App
  call :Ins15Lic %%a
  )
)
for %%a in (Access) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 (
  echo %%a 2013 App
  call :Ins15Lic %%a
  )
)
for %%a in (Lync) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 (
  echo SkypeforBusiness 2015 App
  call :Ins15Lic %%a
  )
)
goto :GVLKC2R

:InsLic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
set "_kpey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=PidKey=%2"
set "_kpey=%2"
)
reg delete %_Config% /f /v %_ID%.OSPPReady %_Nul3%
"!_Integrator!" /I /License PRIDName=%_ID%.16 %_pkey% PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!" %_Nul1%

set fallback=0
set "_qr=wmic path %_spp% where ApplicationID='%_oApp%' get LicenseFamily"
if %WMI_VBS% NEQ 0 set "_qr=%_csq% %_spp% "ApplicationID='%_oApp%'" LicenseFamily"
%_qr% %_Nul2% | find /i "%_patt%" %_Nul1% || (set fallback=1)
if %fallback% equ 0 goto :IntOK

set "_lsfs="
for %%# in ("!_LicensesPath!\%_patt%*.xrm-ms") do (
set "_lsfs=!_lsfs! %%~nx#"
)
if defined _kpey (
  for %%# in ("!_LicensesPath!\%1DemoR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1E5R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1EDUR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1MSDNR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1O365R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1CO365R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
)
for %%# in (!_lsfs!) do (
%_cscript% %_vbsf%"!_LicensesPath!\%%#"
)
set "_qr=wmic path %_sps% where Version='%_wmi%' call InstallProductKey ProductKey="%_kpey%""
if %WMI_VBS% NEQ 0 set "_qr=%_csp% %_sps% "%_kpey%""
if defined _kpey %_qr% %_Nul3%

:IntOK
reg add %_Config% /f /v %_ID%.OSPPReady /t REG_SZ /d 1 %_Nul1%
reg query %_Config% /v ProductReleaseIds | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config% /v ProductReleaseIds') do reg add %_Config% /v ProductReleaseIds /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:Ins15Lic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=%2"
)
reg delete %_OSPP15Ready% /f /v %_ID%.OSPPReady %_Nul3%
for %%# in ("!_Licenses15Path!\%_patt%*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
set "_qr=wmic path %_sps% where Version='%_wmi%' call InstallProductKey ProductKey="%_pkey%""
if %WMI_VBS% NEQ 0 set "_qr=%_csp% %_sps% "%_pkey%""
if defined _pkey %_qr% %_Nul3%
reg add %_OSPP15Ready% /f /v %_ID%.OSPPReady /t %_OSPP15ReadT% /d 1 %_Nul1%
reg query %_Con15fig% %_Nul2% | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Con15fig% %_Nul6%') do reg add %_Con15fig% /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:GVLKC2R
set _CtRMsg=0
if %_C16Msg% EQU 1 set _CtRMsg=1
if %_C15Msg% EQU 1 set _CtRMsg=1
if %_Office16% EQU 1 (
for %%a in (%_RetIds%,ProPlus) do set "_%%a="
for %%A in (19,21) do call :officeLoc %%A
)
if %_Office15% EQU 1 (
for %%a in (%_R15Ids%,ProPlus) do set "_%%a="
)
set "_qr=wmic path %_sps% where version='%_wmi%' call RefreshLicenseStatus"
if %WMI_VBS% NEQ 0 set "_qr=%_csm% "%_sps%.Version='%_wmi%'" RefreshLicenseStatus"
if %winbuild% GEQ 9200 %_qr% %_Nul3%
if exist "%SysPath%\spp\store_test\2.0\tokens.dat" if %rancopp% EQU 1 if %_CtRMsg% EQU 1 (
%_cscript% %_SLMGR% /rilc
if !ERRORLEVEL! NEQ 0 %_cscript% %_SLMGR% /rilc
)
goto :%_sC2R%

:keys
if "%~1"=="" exit /b
set yh=-
goto :%1 %_Nul2%

:: Windows 11 [Ni]
:59eb965c-9150-42b7-a0ec-22151b9897c5
set "_key=KBN8V%yh%HFGQ4%yh%MGXVD%yh%347P6%yh%PDQGT" &:: IoT Enterprise LTSC
exit /b

:: Windows 11 [Co]
:ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69
set "_key=37D7F%yh%N49CB%yh%WQR8W%yh%TBJ73%yh%FM8RX" &:: SE {Cloud}
exit /b

:d30136fc-cb4b-416e-a23d-87207abc44a9
set "_key=6XN7V%yh%PCBDC%yh%BDBRH%yh%8DQY7%yh%G6R44" &:: SE N {Cloud N}
exit /b

:: Windows 10 [RS5]
:32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee
set "_key=M7XTQ%yh%FN8P6%yh%TTKYV%yh%9D4CC%yh%J462D" &:: Enterprise LTSC 2019
exit /b

:7103a333-b8c8-49cc-93ce-d37c09687f92
set "_key=92NFX%yh%8DJQP%yh%P6BBQ%yh%THF9C%yh%7CG2H" &:: Enterprise LTSC 2019 N
exit /b

:ec868e65-fadf-4759-b23e-93fe37f2cc29
set "_key=CPWHC%yh%NT2C7%yh%VYW78%yh%DHDB2%yh%PG3GK" &:: Enterprise for Virtual Desktops
exit /b

:0df4f814-3f57-4b8b-9a9d-fddadcd69fac
set "_key=NBTWJ%yh%3DR69%yh%3C4V8%yh%C26MC%yh%GQ9M6" &:: Lean
exit /b

:: Windows 10 [RS3]
:82bbc092-bc50-4e16-8e18-b74fc486aec3
set "_key=NRG8B%yh%VKK3Q%yh%CXVCJ%yh%9G2XF%yh%6Q84J" &:: Pro Workstation
exit /b

:4b1571d3-bafb-4b40-8087-a961be2caf65
set "_key=9FNHH%yh%K3HBT%yh%3W4TD%yh%6383H%yh%6XYWF" &:: Pro Workstation N
exit /b

:e4db50ea-bda1-4566-b047-0ca50abc6f07
set "_key=7NBT4%yh%WGBQX%yh%MP4H7%yh%QXFF8%yh%YP3KX" &:: Enterprise Remote Server
exit /b

:: Windows 10 [RS2]
:e0b2d383-d112-413f-8a80-97f373a5820c
set "_key=YYVX9%yh%NTFWV%yh%6MDM3%yh%9PT4T%yh%4M68B" &:: Enterprise G
exit /b

:e38454fb-41a4-4f59-a5dc-25080e354730
set "_key=44RPN%yh%FTY23%yh%9VTTB%yh%MP9BX%yh%T84FV" &:: Enterprise G N
exit /b

:: Windows 10 [RS1]
:2d5a5a60-3040-48bf-beb0-fcd770c20ce0
set "_key=DCPHK%yh%NFMTC%yh%H88MJ%yh%PFHPY%yh%QJ4BJ" &:: Enterprise 2016 LTSB
exit /b

:9f776d83-7156-45b2-8a5c-359b9c9f22a3
set "_key=QFFDN%yh%GRT3P%yh%VKWWX%yh%X7T3R%yh%8B639" &:: Enterprise 2016 LTSB N
exit /b

:3f1afc82-f8ac-4f6c-8005-1d233e606eee
set "_key=6TP4R%yh%GNPTD%yh%KYYHQ%yh%7B7DP%yh%J447Y" &:: Pro Education
exit /b

:5300b18c-2e33-4dc2-8291-47ffcec746dd
set "_key=YVWGF%yh%BXNMC%yh%HTQYQ%yh%CPQ99%yh%66QFC" &:: Pro Education N
exit /b

:: Windows 10 [TH]
:58e97c99-f377-4ef1-81d5-4ad5522b5fd8
set "_key=TX9XD%yh%98N7V%yh%6WMQ6%yh%BX7FG%yh%H8Q99" &:: Home
exit /b

:7b9e1751-a8da-4f75-9560-5fadfe3d8e38
set "_key=3KHY7%yh%WNT83%yh%DGQKR%yh%F7HPR%yh%844BM" &:: Home N
exit /b

:cd918a57-a41b-4c82-8dce-1a538e221a83
set "_key=7HNRX%yh%D7KGG%yh%3K4RQ%yh%4WPJ4%yh%YTDFH" &:: Home Single Language
exit /b

:a9107544-f4a0-4053-a96a-1479abdef912
set "_key=PVMJN%yh%6DFY6%yh%9CCP6%yh%7BKTT%yh%D3WVR" &:: Home China
exit /b

:2de67392-b7a7-462a-b1ca-108dd189f588
set "_key=W269N%yh%WFGWX%yh%YVC9B%yh%4J6C9%yh%T83GX" &:: Pro
exit /b

:a80b5abf-76ad-428b-b05d-a47d2dffeebf
set "_key=MH37W%yh%N47XK%yh%V7XM9%yh%C7227%yh%GCQG9" &:: Pro N
exit /b

:e0c42288-980c-4788-a014-c080d2e1926e
set "_key=NW6C2%yh%QMPVW%yh%D7KKK%yh%3GKT6%yh%VCFB2" &:: Education
exit /b

:3c102355-d027-42c6-ad23-2e7ef8a02585
set "_key=2WH4N%yh%8QGBV%yh%H22JP%yh%CT43Q%yh%MDWWJ" &:: Education N
exit /b

:73111121-5638-40f6-bc11-f1d7b0d64300
set "_key=NPPR9%yh%FWDCX%yh%D2C8J%yh%H872K%yh%2YT43" &:: Enterprise
exit /b

:e272e3e2-732f-4c65-a8f0-484747d0d947
set "_key=DPH2V%yh%TTNVB%yh%4X9Q3%yh%TJR4H%yh%KHJW4" &:: Enterprise N
exit /b

:7b51a46c-0c04-4e8f-9af4-8496cca90d5e
set "_key=WNMTR%yh%4C88C%yh%JK8YV%yh%HQ7T2%yh%76DF9" &:: Enterprise 2015 LTSB
exit /b

:87b838b7-41b6-4590-8318-5797951d8529
set "_key=2F77B%yh%TNFGY%yh%69QQF%yh%B8YKP%yh%D69TJ" &:: Enterprise 2015 LTSB N
exit /b

:: Windows Server 2022 [Fe]
:9bd77860-9b31-4b7b-96ad-2564017315bf
set "_key=VDYBN%yh%27WPP%yh%V4HQT%yh%9VMD4%yh%VMK7H" &:: Standard
exit /b

:ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03
set "_key=WX4NM%yh%KYWYW%yh%QJJR4%yh%XV3QB%yh%6VM33" &:: Datacenter
exit /b

:8c8f0ad3-9a43-4e05-b840-93b8d1475cbc
set "_key=6N379%yh%GGTMK%yh%23C6M%yh%XVVTC%yh%CKFRQ" &:: Azure Core
exit /b

:f5e9429c-f50b-4b98-b15c-ef92eb5cff39
set "_key=67KN8%yh%4FYJW%yh%2487Q%yh%MQ2J7%yh%4C4RG" &:: Standard ACor
exit /b

:39e69c41-42b4-4a0a-abad-8e3c10a797cc
set "_key=QFND9%yh%D3Y9C%yh%J3KKY%yh%6RPVP%yh%2DPYV" &:: Datacenter ACor
exit /b

:: Windows Server 2019 [RS5]
:de32eafd-aaee-4662-9444-c1befb41bde2
set "_key=N69G4%yh%B89J2%yh%4G8F4%yh%WWYCC%yh%J464C" &:: Standard
exit /b

:34e1ae55-27f8-4950-8877-7a03be5fb181
set "_key=WMDGN%yh%G9PQG%yh%XVVXX%yh%R3X43%yh%63DFG" &:: Datacenter
exit /b

:a99cc1f0-7719-4306-9645-294102fbff95
set "_key=FDNH6%yh%VW9RW%yh%BXPJ7%yh%4XTYG%yh%239TB" &:: Azure Core
exit /b

:73e3957c-fc0c-400d-9184-5f7b6f2eb409
set "_key=N2KJX%yh%J94YW%yh%TQVFB%yh%DG9YT%yh%724CC" &:: Standard ACor
exit /b

:90c362e5-0da1-4bfd-b53b-b87d309ade43
set "_key=6NMRW%yh%2C8FM%yh%D24W7%yh%TQWMY%yh%CWH2D" &:: Datacenter ACor
exit /b

:034d3cbb-5d4b-4245-b3f8-f84571314078
set "_key=WVDHN%yh%86M7X%yh%466P6%yh%VHXV7%yh%YY726" &:: Essentials
exit /b

:8de8eb62-bbe0-40ac-ac17-f75595071ea3
set "_key=GRFBW%yh%QNDC4%yh%6QBHG%yh%CCK3B%yh%2PR88" &:: ServerARM64
exit /b

:19b5e0fb-4431-46bc-bac1-2f1873e4ae73
set "_key=NTBV8%yh%9K7Q8%yh%V27C6%yh%M2BTV%yh%KHMXV" &:: Azure Datacenter - ServerTurbine
exit /b

:: Windows Server 2016 [RS4]
:43d9af6e-5e86-4be8-a797-d072a046896c
set "_key=K9FYF%yh%G6NCK%yh%73M32%yh%XMVPY%yh%F9DRR" &:: ServerARM64
exit /b

:: Windows Server 2016 [RS3]
:61c5ef22-f14f-4553-a824-c4b31e84b100
set "_key=PTXN8%yh%JFHJM%yh%4WC78%yh%MPCBR%yh%9W4KR" &:: Standard ACor
exit /b

:e49c08e7-da82-42f8-bde2-b570fbcae76c
set "_key=2HXDN%yh%KRXHB%yh%GPYC7%yh%YCKFJ%yh%7FVDG" &:: Datacenter ACor
exit /b

:: Windows Server 2016 [RS1]
:8c1c5410-9f39-4805-8c9d-63a07706358f
set "_key=WC2BQ%yh%8NRM3%yh%FDDYY%yh%2BFGV%yh%KHKQY" &:: Standard
exit /b

:21c56779-b449-4d20-adfc-eece0e1ad74b
set "_key=CB7KF%yh%BWN84%yh%R7R2Y%yh%793K2%yh%8XDDG" &:: Datacenter
exit /b

:3dbf341b-5f6c-4fa7-b936-699dce9e263f
set "_key=VP34G%yh%4NPPG%yh%79JTQ%yh%864T4%yh%R3MQX" &:: Azure Core
exit /b

:2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283
set "_key=JCKRF%yh%N37P4%yh%C2D82%yh%9YXRT%yh%4M63B" &:: Essentials
exit /b

:7b4433f4-b1e7-4788-895a-c45378d38253
set "_key=QN4C6%yh%GBJD2%yh%FB422%yh%GHWJK%yh%GJG2R" &:: Cloud Storage
exit /b

:: Windows 8.1
:fe1c3238-432a-43a1-8e25-97e7d1ef10f3
set "_key=M9Q9P%yh%WNJJT%yh%6PXPY%yh%DWX8H%yh%6XWKK" &:: Core
exit /b

:78558a64-dc19-43fe-a0d0-8075b2a370a3
set "_key=7B9N3%yh%D94CG%yh%YTVHR%yh%QBPX3%yh%RJP64" &:: Core N
exit /b

:c72c6a1d-f252-4e7e-bdd1-3fca342acb35
set "_key=BB6NG%yh%PQ82V%yh%VRDPW%yh%8XVD2%yh%V8P66" &:: Core Single Language
exit /b

:db78b74f-ef1c-4892-abfe-1e66b8231df6
set "_key=NCTT7%yh%2RGK8%yh%WMHRF%yh%RY7YQ%yh%JTXG3" &:: Core China
exit /b

:ffee456a-cd87-4390-8e07-16146c672fd0
set "_key=XYTND%yh%K6QKT%yh%K2MRH%yh%66RTM%yh%43JKP" &:: Core ARM
exit /b

:c06b6981-d7fd-4a35-b7b4-054742b7af67
set "_key=GCRJD%yh%8NW9H%yh%F2CDX%yh%CCM8D%yh%9D6T9" &:: Pro
exit /b

:7476d79f-8e48-49b4-ab63-4d0b813a16e4
set "_key=HMCNV%yh%VVBFX%yh%7HMBH%yh%CTY9B%yh%B4FXY" &:: Pro N
exit /b

:096ce63d-4fac-48a9-82a9-61ae9e800e5f
set "_key=789NJ%yh%TQK6T%yh%6XTH8%yh%J39CJ%yh%J8D3P" &:: Pro with Media Center
exit /b

:81671aaf-79d1-4eb1-b004-8cbbe173afea
set "_key=MHF9N%yh%XY6XB%yh%WVXMC%yh%BTDCT%yh%MKKG7" &:: Enterprise
exit /b

:113e705c-fa49-48a4-beea-7dd879b46b14
set "_key=TT4HM%yh%HN7YT%yh%62K67%yh%RGRQJ%yh%JFFXW" &:: Enterprise N
exit /b

:0ab82d54-47f4-4acb-818c-cc5bf0ecb649
set "_key=NMMPB%yh%38DD4%yh%R2823%yh%62W8D%yh%VXKJB" &:: Embedded Industry Pro
exit /b

:cd4e2d9f-5059-4a50-a92d-05d5bb1267c7
set "_key=FNFKF%yh%PWTVT%yh%9RC8H%yh%32HB2%yh%JB34X" &:: Embedded Industry Enterprise
exit /b

:f7e88590-dfc7-4c78-bccb-6f3865b99d1a
set "_key=VHXM3%yh%NR6FT%yh%RY6RT%yh%CK882%yh%KW2CJ" &:: Embedded Industry Automotive
exit /b

:e9942b32-2e55-4197-b0bd-5ff58cba8860
set "_key=3PY8R%yh%QHNP9%yh%W7XQD%yh%G6DPH%yh%3J2C9" &:: with Bing
exit /b

:c6ddecd6-2354-4c19-909b-306a3058484e
set "_key=Q6HTR%yh%N24GM%yh%PMJFP%yh%69CD8%yh%2GXKR" &:: with Bing N
exit /b

:b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c
set "_key=KF37N%yh%VDV38%yh%GRRTV%yh%XH8X6%yh%6F3BB" &:: with Bing Single Language
exit /b

:ba998212-460a-44db-bfb5-71bf09d1c68b
set "_key=R962J%yh%37N87%yh%9VVK2%yh%WJ74P%yh%XTMHR" &:: with Bing China
exit /b

:e58d87b5-8126-4580-80fb-861b22f79296
set "_key=MX3RK%yh%9HNGX%yh%K3QKC%yh%6PJ3F%yh%W8D7B" &:: Pro for Students
exit /b

:cab491c7-a918-4f60-b502-dab75e334f40
set "_key=TNFGH%yh%2R6PB%yh%8XM3K%yh%QYHX2%yh%J4296" &:: Pro for Students N
exit /b

:: Windows Server 2012 R2
:b3ca044e-a358-4d68-9883-aaa2941aca99
set "_key=D2N9P%yh%3P6X9%yh%2R39C%yh%7RTCD%yh%MDVJX" &:: Standard
exit /b

:00091344-1ea4-4f37-b789-01750ba6988c
set "_key=W3GGN%yh%FT8W3%yh%Y4M27%yh%J84CP%yh%Q3VJ9" &:: Datacenter
exit /b

:21db6ba4-9a7b-4a14-9e29-64a60c59301d
set "_key=KNC87%yh%3J2TX%yh%XB4WP%yh%VCPJV%yh%M4FWM" &:: Essentials
exit /b

:b743a2be-68d4-4dd3-af32-92425b7bb623
set "_key=3NPTF%yh%33KPT%yh%GGBPR%yh%YX76B%yh%39KDD" &:: Cloud Storage
exit /b

:: Windows 8
:c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60
set "_key=BN3D2%yh%R7TKB%yh%3YPBD%yh%8DRP2%yh%27GG4" &:: Core
exit /b

:197390a0-65f6-4a95-bdc4-55d58a3b0253
set "_key=8N2M2%yh%HWPGY%yh%7PGT9%yh%HGDD8%yh%GVGGY" &:: Core N
exit /b

:8860fcd4-a77b-4a20-9045-a150ff11d609
set "_key=2WN2H%yh%YGCQR%yh%KFX6K%yh%CD6TF%yh%84YXQ" &:: Core Single Language
exit /b

:9d5584a2-2d85-419a-982c-a00888bb9ddf
set "_key=4K36P%yh%JN4VD%yh%GDC6V%yh%KDT89%yh%DYFKP" &:: Core China
exit /b

:af35d7b7-5035-4b63-8972-f0b747b9f4dc
set "_key=DXHJF%yh%N9KQX%yh%MFPVR%yh%GHGQK%yh%Y7RKV" &:: Core ARM
exit /b

:a98bcd6d-5343-4603-8afe-5908e4611112
set "_key=NG4HW%yh%VH26C%yh%733KW%yh%K6F98%yh%J8CK4" &:: Pro
exit /b

:ebf245c1-29a8-4daf-9cb1-38dfc608a8c8
set "_key=XCVCF%yh%2NXM9%yh%723PB%yh%MHCB7%yh%2RYQQ" &:: Pro N
exit /b

:a00018a3-f20f-4632-bf7c-8daa5351c914
set "_key=GNBB8%yh%YVD74%yh%QJHX6%yh%27H4K%yh%8QHDG" &:: Pro with Media Center
exit /b

:458e1bec-837a-45f6-b9d5-925ed5d299de
set "_key=32JNW%yh%9KQ84%yh%P47T8%yh%D8GGY%yh%CWCK7" &:: Enterprise
exit /b

:e14997e7-800a-4cf7-ad10-de4b45b578db
set "_key=JMNMF%yh%RHW7P%yh%DMY6X%yh%RF3DR%yh%X2BQT" &:: Enterprise N
exit /b

:10018baf-ce21-4060-80bd-47fe74ed4dab
set "_key=RYXVT%yh%BNQG7%yh%VD29F%yh%DBMRY%yh%HT73M" &:: Embedded Industry Pro
exit /b

:18db1848-12e0-4167-b9d7-da7fcda507db
set "_key=NKB3R%yh%R2F8T%yh%3XCDP%yh%7Q2KW%yh%XWYQ2" &:: Embedded Industry Enterprise
exit /b

:: Windows Server 2012
:f0f5ec41-0d55-4732-af02-440a44a3cf0f
set "_key=XC9B7%yh%NBPP2%yh%83J2H%yh%RHMBY%yh%92BT4" &:: Standard
exit /b

:d3643d60-0c42-412d-a7d6-52e6635327f6
set "_key=48HP8%yh%DN98B%yh%MYWDG%yh%T2DCC%yh%8W83P" &:: Datacenter
exit /b

:7d5486c7-e120-4771-b7f1-7b56c6d3170c
set "_key=HM7DN%yh%YVMH3%yh%46JC3%yh%XYTG7%yh%CYQJJ" &:: MultiPoint Standard
exit /b

:95fd1c83-7df5-494a-be8b-1300e1c9d1cd
set "_key=XNH6W%yh%2V9GX%yh%RGJ4K%yh%Y8X6F%yh%QGJ2G" &:: MultiPoint Premium
exit /b

:: Windows 7
:b92e9980-b9d5-4821-9c94-140f632f6312
set "_key=FJ82H%yh%XT6CR%yh%J8D7P%yh%XQJJ2%yh%GPDD4" &:: Professional
exit /b

:54a09a0d-d57b-4c10-8b69-a842d6590ad5
set "_key=MRPKT%yh%YTG23%yh%K7D7T%yh%X2JMM%yh%QY7MG" &:: Professional N
exit /b

:5a041529-fef8-4d07-b06f-b59b573b32d2
set "_key=W82YF%yh%2Q76Y%yh%63HXB%yh%FGJG9%yh%GF7QX" &:: Professional E
exit /b

:ae2ee509-1b34-41c0-acb7-6d4650168915
set "_key=33PXH%yh%7Y6KF%yh%2VJC9%yh%XBBR8%yh%HVTHH" &:: Enterprise
exit /b

:1cb6d605-11b3-4e14-bb30-da91c8e3983a
set "_key=YDRBP%yh%3D83W%yh%TY26F%yh%D46B2%yh%XCKRJ" &:: Enterprise N
exit /b

:46bbed08-9c7b-48fc-a614-95250573f4ea
set "_key=C29WB%yh%22CC8%yh%VJ326%yh%GHFJW%yh%H9DH4" &:: Enterprise E
exit /b

:db537896-376f-48ae-a492-53d0547773d0
set "_key=YBYF6%yh%BHCR3%yh%JPKRB%yh%CDW7B%yh%F9BK4" &:: Embedded POSReady 7
exit /b

:e1a8296a-db37-44d1-8cce-7bc961d59c54
set "_key=XGY72%yh%BRBBT%yh%FF8MH%yh%2GG8H%yh%W7KCW" &:: Embedded Standard
exit /b

:aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395
set "_key=73KQT%yh%CD9G6%yh%K7TQG%yh%66MRP%yh%CQ22C" &:: Embedded ThinPC
exit /b

:: Windows Server 2008 R2
:a78b8bd9-8017-4df5-b86a-09f756affa7c
set "_key=6TPJF%yh%RBVHG%yh%WBW2R%yh%86QPH%yh%6RTM4" &:: Web
exit /b

:cda18cf3-c196-46ad-b289-60c072869994
set "_key=TT8MH%yh%CG224%yh%D3D7Q%yh%498W2%yh%9QCTX" &:: HPC
exit /b

:68531fb9-5511-4989-97be-d11a0f55633f
set "_key=YC6KT%yh%GKW9T%yh%YTKYR%yh%T4X34%yh%R7VHC" &:: Standard
exit /b

:7482e61b-c589-4b7f-8ecc-46d455ac3b87
set "_key=74YFP%yh%3QFB3%yh%KQT8W%yh%PMXWJ%yh%7M648" &:: Datacenter
exit /b

:620e2b3d-09e7-42fd-802a-17a13652fe7a
set "_key=489J6%yh%VHDMP%yh%X63PK%yh%3K798%yh%CPX3Y" &:: Enterprise
exit /b

:8a26851c-1c7e-48d3-a687-fbca9b9ac16b
set "_key=GT63C%yh%RJFQ3%yh%4GMB6%yh%BRFB9%yh%CB83V" &:: Itanium
exit /b

:f772515c-0e87-48d5-a676-e6962c3e1195
set "_key=736RG%yh%XDKJK%yh%V34PF%yh%BHK87%yh%J6X3K" &:: MultiPoint Server - ServerEmbeddedSolution
exit /b

:: Office 2021
:fbdb3e18-a8ef-4fb3-9183-dffd60bd0984
set "_key=FXYTK%yh%NJJ8C%yh%GB6DW%yh%3DYQT%yh%6F7TH" &:: Professional Plus
exit /b

:080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3
set "_key=KDX7X%yh%BNVR8%yh%TXXGX%yh%4Q7Y8%yh%78VT3" &:: Standard
exit /b

:76881159-155c-43e0-9db7-2d70a9a3a4ca
set "_key=FTNWT%yh%C6WBT%yh%8HMGF%yh%K9PRX%yh%QV9H8" &:: Project Professional
exit /b

:6dd72704-f752-4b71-94c7-11cec6bfc355
set "_key=J2JDC%yh%NJCYY%yh%9RGQ4%yh%YXWMH%yh%T3D4T" &:: Project Standard
exit /b

:fb61ac9a-1688-45d2-8f6b-0674dbffa33c
set "_key=KNH8D%yh%FGHT4%yh%T8RK3%yh%CTDYJ%yh%K2HT4" &:: Visio Professional
exit /b

:72fce797-1884-48dd-a860-b2f6a5efd3ca
set "_key=MJVNY%yh%BYWPY%yh%CWV6J%yh%2RKRT%yh%4M8QG" &:: Visio Standard
exit /b

:1fe429d8-3fa7-4a39-b6f0-03dded42fe14
set "_key=WM8YG%yh%YNGDD%yh%4JHDC%yh%PG3F4%yh%FC4T4" &:: Access
exit /b

:ea71effc-69f1-4925-9991-2f5e319bbc24
set "_key=NWG3X%yh%87C9K%yh%TC7YY%yh%BC2G7%yh%G6RVC" &:: Excel
exit /b

:a5799e4c-f83c-4c6e-9516-dfe9b696150b
set "_key=C9FM6%yh%3N72F%yh%HFJXB%yh%TM3V9%yh%T86R9" &:: Outlook
exit /b

:6e166cc3-495d-438a-89e7-d7c9e6fd4dea
set "_key=TY7XF%yh%NFRBR%yh%KJ44C%yh%G83KF%yh%GX27K" &:: PowerPoint
exit /b

:aa66521f-2370-4ad8-a2bb-c095e3e4338f
set "_key=2MW9D%yh%N4BXM%yh%9VBPG%yh%Q7W6M%yh%KFBGQ" &:: Publisher
exit /b

:1f32a9af-1274-48bd-ba1e-1ab7508a23e8
set "_key=HWCXN%yh%K3WBT%yh%WJBKY%yh%R8BD9%yh%XK29P" &:: Skype for Business
exit /b

:abe28aea-625a-43b1-8e30-225eb8fbd9e5
set "_key=TN8H9%yh%M34D3%yh%Y64V9%yh%TR72V%yh%X79KV" &:: Word
exit /b

:: Office 2019
:85dd8b5f-eaa4-4af3-a628-cce9e77c9a03
set "_key=NMMKJ%yh%6RK4F%yh%KMJVX%yh%8D9MJ%yh%6MWKP" &:: Professional Plus
exit /b

:6912a74b-a5fb-401a-bfdb-2e3ab46f4b02
set "_key=6NWWJ%yh%YQWMR%yh%QKGCB%yh%6TMB3%yh%9D9HK" &:: Standard
exit /b

:2ca2bf3f-949e-446a-82c7-e25a15ec78c4
set "_key=B4NPR%yh%3FKK7%yh%T2MBV%yh%FRQ4W%yh%PKD2B" &:: Project Professional
exit /b

:1777f0e3-7392-4198-97ea-8ae4de6f6381
set "_key=C4F7P%yh%NCP8C%yh%6CQPT%yh%MQHV9%yh%JXD2M" &:: Project Standard
exit /b

:5b5cf08f-b81a-431d-b080-3450d8620565
set "_key=9BGNQ%yh%K37YR%yh%RQHF2%yh%38RQ3%yh%7VCBB" &:: Visio Professional
exit /b

:e06d7df3-aad0-419d-8dfb-0ac37e2bdf39
set "_key=7TQNQ%yh%K3YQQ%yh%3PFH7%yh%CCPPM%yh%X4VQ2" &:: Visio Standard
exit /b

:9e9bceeb-e736-4f26-88de-763f87dcc485
set "_key=9N9PT%yh%27V4Y%yh%VJ2PD%yh%YXFMF%yh%YTFQT" &:: Access
exit /b

:237854e9-79fc-4497-a0c1-a70969691c6b
set "_key=TMJWT%yh%YYNMB%yh%3BKTF%yh%644FC%yh%RVXBD" &:: Excel
exit /b

:c8f8a301-19f5-4132-96ce-2de9d4adbd33
set "_key=7HD7K%yh%N4PVK%yh%BHBCQ%yh%YWQRW%yh%XW4VK" &:: Outlook
exit /b

:3131fd61-5e4f-4308-8d6d-62be1987c92c
set "_key=RRNCX%yh%C64HY%yh%W2MM7%yh%MCH9G%yh%TJHMQ" &:: PowerPoint
exit /b

:9d3e4cca-e172-46f1-a2f4-1d2107051444
set "_key=G2KWX%yh%3NW6P%yh%PY93R%yh%JXK2T%yh%C9Y9V" &:: Publisher
exit /b

:734c6c6e-b0ba-4298-a891-671772b2bd1b
set "_key=NCJ33%yh%JHBBY%yh%HTK98%yh%MYCV8%yh%HMKHJ" &:: Skype for Business
exit /b

:059834fe-a8ea-4bff-b67b-4d006b5447d3
set "_key=PBX3G%yh%NWMT6%yh%Q7XBW%yh%PYJGG%yh%WXD33" &:: Word
exit /b

:0bc88885-718c-491d-921f-6f214349e79c
set "_key=VQ9DP%yh%NVHPH%yh%T9HJC%yh%J9PDT%yh%KTQRG" &:: Pro Plus 2019 Preview
exit /b

:fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9
set "_key=XM2V9%yh%DN9HH%yh%QB449%yh%XDGKC%yh%W2RMW" &:: Project Pro 2019 Preview
exit /b

:500f6619-ef93-4b75-bcb4-82819998a3ca
set "_key=N2CG9%yh%YD3YK%yh%936X4%yh%3WR82%yh%Q3X4H" &:: Visio Pro 2019 Preview
exit /b

:f3fb2d68-83dd-4c8b-8f09-08e0d950ac3b
set "_key=HFPBN%yh%RYGG8%yh%HQWCW%yh%26CH6%yh%PDPVF" &:: Pro Plus 2021 Preview
exit /b

:76093b1b-7057-49d7-b970-638ebcbfd873
set "_key=WDNBY%yh%PCYFY%yh%9WP6G%yh%BXVXM%yh%92HDV" &:: Project Pro 2021 Preview
exit /b

:a3b44174-2451-4cd6-b25f-66638bfb9046
set "_key=2XYX7%yh%NXXBK%yh%9CK7W%yh%K2TKW%yh%JFJ7G" &:: Visio Pro 2021 Preview
exit /b

:: Office 2016
:829b8110-0e6f-4349-bca4-42803577788d
set "_key=WGT24%yh%HCNMF%yh%FQ7XH%yh%6M8K7%yh%DRTW9" &:: Project Professional C2R-P
exit /b

:cbbaca45-556a-4416-ad03-bda598eaa7c8
set "_key=D8NRQ%yh%JTYM3%yh%7J2DX%yh%646CT%yh%6836M" &:: Project Standard C2R-P
exit /b

:b234abe3-0857-4f9c-b05a-4dc314f85557
set "_key=69WXN%yh%MBYV6%yh%22PQG%yh%3WGHK%yh%RM6XC" &:: Visio Professional C2R-P
exit /b

:361fe620-64f4-41b5-ba77-84f8e079b1f7
set "_key=NY48V%yh%PPYYH%yh%3F4PX%yh%XJRKJ%yh%W4423" &:: Visio Standard C2R-P
exit /b

:e914ea6e-a5fa-4439-a394-a9bb3293ca09
set "_key=DMTCJ%yh%KNRKX%yh%26982%yh%JYCKT%yh%P7KB6" &:: MondoR
exit /b

:9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
set "_key=HFTND%yh%W9MK4%yh%8B7MJ%yh%B6C4G%yh%XQBR2" &:: Mondo
exit /b

:d450596f-894d-49e0-966a-fd39ed4c4c64
set "_key=XQNVK%yh%8JYDB%yh%WJ9W3%yh%YJ8YR%yh%WFG99" &:: Professional Plus
exit /b

:dedfa23d-6ed1-45a6-85dc-63cae0546de6
set "_key=JNRGM%yh%WHDWX%yh%FJJG3%yh%K47QV%yh%DRTFM" &:: Standard
exit /b

:4f414197-0fc2-4c01-b68a-86cbb9ac254c
set "_key=YG9NW%yh%3K39V%yh%2T3HJ%yh%93F3Q%yh%G83KT" &:: Project Professional
exit /b

:da7ddabc-3fbe-4447-9e01-6ab7440b4cd4
set "_key=GNFHQ%yh%F6YQM%yh%KQDGJ%yh%327XX%yh%KQBVC" &:: Project Standard
exit /b

:6bf301c1-b94a-43e9-ba31-d494598c47fb
set "_key=PD3PC%yh%RHNGV%yh%FXJ29%yh%8JK7D%yh%RJRJK" &:: Visio Professional
exit /b

:aa2a7821-1827-4c2c-8f1d-4513a34dda97
set "_key=7WHWN%yh%4T7MP%yh%G96JF%yh%G33KR%yh%W8GF4" &:: Visio Standard
exit /b

:67c0fc0c-deba-401b-bf8b-9c8ad8395804
set "_key=GNH9Y%yh%D2J4T%yh%FJHGG%yh%QRVH7%yh%QPFDW" &:: Access
exit /b

:c3e65d36-141f-4d2f-a303-a842ee756a29
set "_key=9C2PK%yh%NWTVB%yh%JMPW8%yh%BFT28%yh%7FTBF" &:: Excel
exit /b

:d8cace59-33d2-4ac7-9b1b-9b72339c51c8
set "_key=DR92N%yh%9HTF2%yh%97XKM%yh%XW2WJ%yh%XW3J6" &:: OneNote
exit /b

:ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
set "_key=R69KK%yh%NTPKF%yh%7M3Q4%yh%QYBHW%yh%6MT9B" &:: Outlook
exit /b

:d70b1bba-b893-4544-96e2-b7a318091c33
set "_key=J7MQP%yh%HNJ4Y%yh%WJ7YM%yh%PFYGF%yh%BY6C6" &:: Powerpoint
exit /b

:041a06cb-c5b8-4772-809f-416d03d16654
set "_key=F47MM%yh%N3XJP%yh%TQXJ9%yh%BP99D%yh%8K837" &:: Publisher
exit /b

:83e04ee1-fa8d-436d-8994-d31a862cab77
set "_key=869NQ%yh%FJ69K%yh%466HW%yh%QYCP2%yh%DDBV6" &:: Skype for Business
exit /b

:bb11badf-d8aa-470e-9311-20eaf80fe5cc
set "_key=WXY84%yh%JN2Q9%yh%RBCCQ%yh%3Q3J3%yh%3PFJ6" &:: Word
exit /b

:: Office 2013
:dc981c6b-fc8e-420f-aa43-f8f33e5c0923
set "_key=42QTK%yh%RN8M7%yh%J3C4G%yh%BBGYM%yh%88CYV" &:: Mondo
exit /b

:b322da9c-a2e2-4058-9e4e-f59a6970bd69
set "_key=YC7DK%yh%G2NP3%yh%2QQC3%yh%J6H88%yh%GVGXT" &:: Professional Plus
exit /b

:b13afb38-cd79-4ae5-9f7f-eed058d750ca
set "_key=KBKQT%yh%2NMXY%yh%JJWGP%yh%M62JB%yh%92CD4" &:: Standard
exit /b

:4a5d124a-e620-44ba-b6ff-658961b33b9a
set "_key=FN8TT%yh%7WMH6%yh%2D4X9%yh%M337T%yh%2342K" &:: Project Professional
exit /b

:427a28d1-d17c-4abf-b717-32c780ba6f07
set "_key=6NTH3%yh%CW976%yh%3G3Y2%yh%JK3TX%yh%8QHTT" &:: Project Standard
exit /b

:e13ac10e-75d0-4aff-a0cd-764982cf541c
set "_key=C2FG9%yh%N6J68%yh%H8BTJ%yh%BW3QX%yh%RM3B3" &:: Visio Professional
exit /b

:ac4efaf0-f81f-4f61-bdf7-ea32b02ab117
set "_key=J484Y%yh%4NKBF%yh%W2HMG%yh%DBMJC%yh%PGWR7" &:: Visio Standard
exit /b

:6ee7622c-18d8-4005-9fb7-92db644a279b
set "_key=NG2JY%yh%H4JBT%yh%HQXYP%yh%78QH9%yh%4JM2D" &:: Access
exit /b

:f7461d52-7c2b-43b2-8744-ea958e0bd09a
set "_key=VGPNG%yh%Y7HQW%yh%9RHP7%yh%TKPV3%yh%BG7GB" &:: Excel
exit /b

:fb4875ec-0c6b-450f-b82b-ab57d8d1677f
set "_key=H7R7V%yh%WPNXQ%yh%WCYYC%yh%76BGV%yh%VT7GH" &:: Groove
exit /b

:a30b8040-d68a-423f-b0b5-9ce292ea5a8f
set "_key=DKT8B%yh%N7VXH%yh%D963P%yh%Q4PHY%yh%F8894" &:: InfoPath
exit /b

:1b9f11e3-c85c-4e1b-bb29-879ad2c909e3
set "_key=2MG3G%yh%3BNTT%yh%3MFW9%yh%KDQW3%yh%TCK7R" &:: Lync
exit /b

:efe1f3e6-aea2-4144-a208-32aa872b6545
set "_key=TGN6P%yh%8MMBC%yh%37P2F%yh%XHXXK%yh%P34VW" &:: OneNote
exit /b

:771c3afa-50c5-443f-b151-ff2546d863a0
set "_key=QPN8Q%yh%BJBTJ%yh%334K3%yh%93TGY%yh%2PMBT" &:: Outlook
exit /b

:8c762649-97d1-4953-ad27-b7e2c25b972e
set "_key=4NT99%yh%8RJFH%yh%Q2VDH%yh%KYG2C%yh%4RD4F" &:: Powerpoint
exit /b

:00c79ff1-6850-443d-bf61-71cde0de305f
set "_key=PN2WF%yh%29XG2%yh%T9HJ7%yh%JQPJR%yh%FCXK4" &:: Publisher
exit /b

:d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3
set "_key=6Q7VD%yh%NX8JD%yh%WJ2VH%yh%88V73%yh%4GBJ7" &:: Word
exit /b

:: Office 2010
:09ed9640-f020-400a-acd8-d7d867dfd9c2
set "_key=YBJTT%yh%JG6MD%yh%V9Q7P%yh%DBKXJ%yh%38W9R" &:: Mondo
exit /b

:ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c
set "_key=7TC2V%yh%WXF6P%yh%TD7RT%yh%BQRXR%yh%B8K32" &:: Mondo2
exit /b

:6f327760-8c5c-417c-9b61-836a98287e0c
set "_key=VYBBJ%yh%TRJPB%yh%QFQRF%yh%QFT4D%yh%H3GVB" &:: Professional Plus
exit /b

:9da2a678-fb6b-4e67-ab84-60dd6a9c819a
set "_key=V7QKV%yh%4XVVR%yh%XYV4D%yh%F7DFM%yh%8R6BM" &:: Standard
exit /b

:df133ff7-bf14-4f95-afe3-7b48e7e331ef
set "_key=YGX6F%yh%PGV49%yh%PGW3J%yh%9BTGG%yh%VHKC6" &:: Project Professional
exit /b

:5dc7bf61-5ec9-4996-9ccb-df806a2d0efe
set "_key=4HP3K%yh%88W3F%yh%W2K3D%yh%6677X%yh%F9PGB" &:: Project Standard
exit /b

:92236105-bb67-494f-94c7-7f7a607929bd
set "_key=D9DWC%yh%HPYVV%yh%JGF4P%yh%BTWQB%yh%WX8BJ" &:: Visio Premium
exit /b

:e558389c-83c3-4b29-adfe-5e4d7f46c358
set "_key=7MCW8%yh%VRQVK%yh%G677T%yh%PDJCM%yh%Q8TCP" &:: Visio Professional
exit /b

:9ed833ff-4f92-4f36-b370-8683a4f13275
set "_key=767HD%yh%QGMWX%yh%8QTDB%yh%9G3R2%yh%KHFGJ" &:: Visio Standard
exit /b

:8ce7e872-188c-4b98-9d90-f8f90b7aad02
set "_key=V7Y44%yh%9T38C%yh%R2VJK%yh%666HK%yh%T7DDX" &:: Access
exit /b

:cee5d470-6e3b-4fcc-8c2b-d17428568a9f
set "_key=H62QG%yh%HXVKF%yh%PP4HP%yh%66KMR%yh%CW9BM" &:: Excel
exit /b

:8947d0b8-c33b-43e1-8c56-9b674c052832
set "_key=QYYW6%yh%QP4CB%yh%MBV6G%yh%HYMCJ%yh%4T3J4" &:: Groove - SharePoint Workspace
exit /b

:ca6b6639-4ad6-40ae-a575-14dee07f6430
set "_key=K96W8%yh%67RPQ%yh%62T9Y%yh%J8FQJ%yh%BT37T" &:: InfoPath
exit /b

:ab586f5c-5256-4632-962f-fefd8b49e6f4
set "_key=Q4Y4M%yh%RHWJM%yh%PY37F%yh%MTKWH%yh%D3XHX" &:: OneNote
exit /b

:ecb7c192-73ab-4ded-acf4-2399b095d0cc
set "_key=7YDC2%yh%CWM8M%yh%RRTJC%yh%8MDVC%yh%X3DWQ" &:: Outlook
exit /b

:45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a
set "_key=RC8FX%yh%88JRY%yh%3PF7C%yh%X8P67%yh%P4VTT" &:: Powerpoint
exit /b

:b50c4f75-599b-43e8-8dcd-1081a7967241
set "_key=BFK7F%yh%9MYHM%yh%V68C7%yh%DRQ66%yh%83YTP" &:: Publisher
exit /b

:2d0882e7-a4e7-423b-8ccc-70d91e0158b1
set "_key=HVHB3%yh%C6FV7%yh%KQX9W%yh%YQG79%yh%CRY7T" &:: Word
exit /b

:ea509e87-07a1-4a45-9edc-eba5a39f36af
set "_key=D6QFG%yh%VBYP2%yh%XQHM7%yh%J97RH%yh%VVRCK" &:: Small Business Basics
exit /b

:TheEnd

if %act_failed% EQU 1 (
echo ____________________________________________________________________
echo.
call :_errorinfo
)

if not defined _tskinstalled if not defined _oldtsk (
echo.
if %winbuild% GEQ 9200 (
call :leavenonexistentkms %nul%
echo Keeping the non-existent IP address 10.0.0.10 as KMS Server.
) else (
call :Clear-KMS-Cache
)
)

if not [%Act_OK%]==[1] (
echo.
echo In case of any issues, check https://mass%-%grave.dev/troubleshoot
)

if defined _unattended exit /b

echo ____________________________________________________________________
echo.
call :_color %_Yellow% "Press any key to go back..."
pause >nul
exit /b

::========================================================================================================================================

:_errorinfo

call :CheckFR

set _intcon=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _intcon (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not [%%#]==[] set _intcon=1)
)

if not defined _intcon (
call :_color %_Red% "Internet is not connected."
exit /b
)

if [%ERRORCODE%]==[-1073418124] (
echo Checking Port 1688 connection, it may take a while...
echo.

set /a count=0
set _portcon=
for %%a in (%srvlist%) do if not defined _portcon if !count! LEQ 7 (
set /a count+=1
%psc% "$t = New-Object Net.Sockets.TcpClient;try{$t.Connect("""%%a""", 1688)}catch{};$t.Connected" | findstr /i true 1>nul && set _portcon=1
)

if not defined _portcon (
call :_color %Red% "Port 1688 is blocked in your Internet connection."
echo.
echo Reason:   Probably restricted Internet [Office/College] is connected,
echo           or Firewall is blocking the connection.
echo.
echo Solution: Either use another Internet connection or use offline KMS
echo           https://github.com/abbodi1406/KMS_VL_ALL_AIO
) else (
echo Port 1688 connection test is passed.
echo.
echo Make sure system files are not blocked by your firewall.
echo If the issue persists, try offline KMS
echo https://github.com/abbodi1406/KMS_VL_ALL_AIO
)
echo.
)

echo KMS server is not an issue in this case.
exit /b

::========================================================================================================================================

:setserv

::  Multi KMS servers integration and servers randomization

set srvlist=
set -=

set "srvlist=kms.zhu%-%xiaole.org kms-default.cangs%-%hui.net kms.six%-%yin.com kms.moe%-%club.org kms.cgt%-%soft.com"
set "srvlist=%srvlist% kms.id%-%ina.cn kms.moe%-%yuuko.com xinch%-%eng213618.cn kms.wl%-%rxy.cn kms.ca%-%tqu.com"
set "srvlist=%srvlist% kms.0%-%t.net.cn kms.its%-%jzx.com kms.wx%-%lost.com kms.moe%-%yuuko.top kms.gh%-%pym.com"

set n=1
for %%a in (%srvlist%) do (set %%a=&set server!n!=%%a&set /a n+=1)
set max_servers=15
set /a server_num=0
exit /b

:getserv

if %server_num% equ %max_servers% set /a server_num+=1&set KMS_IP=222.184.9.98&exit /b
set /a rand=%Random%%%(15+1-1)+1
if defined !server%rand%! goto :getserv
set KMS_IP=!server%rand%!
set !server%rand%!=1

::  Get IPv4 address of KMS server to use for the activation, works even if ICMP echo is disabled.
::  Microsoft and Antivirus's may flag the issue if public KMS server host name is directly used for the activation.

set /a server_num+=1
(for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%a"
if [%KMS_IP%]==[!KMS_IP!] for /f "delims=[] tokens=2" %%# in ('pathping -4 -h 1 -n -p 1 -q 1 -w 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%#"
if not [%KMS_IP%]==[!KMS_IP!] exit /b
goto :getserv
)

:==========================================================================================================================================

:Clear-KMS-Cache

set OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform
set SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform

set _wApp=55c92734-d682-4d71-983e-d6ec3f16059f
set _oApp=0ff1ce15-a989-479d-af46-f275c6370663
set _oA14=59a52881-a989-479d-af46-f275c6370663

%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort
%nul% reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
%nul% reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
%nul% reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if defined notx86 (
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName /reg:32
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort /reg:32
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
)
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f
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

:: check KMS38 lock

%nul% reg query "HKLM\%SPPk%\%_wApp%" && (
set error_=9
echo Failed to completely clear KMS Cache.
reg query "HKLM\%SPPk%\%_wApp%" /s 2>nul | findstr /i "127.0.0.2" >nul && echo KMS38 activation is locked.
) || (
echo Cleared KMS Cache successfully.
)
exit /b

:=========================================================================================================================================

:leavenonexistentkms

reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10"
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
if not defined _keepkms38 reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" /reg:32
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" /reg:32
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
)
reg delete "HKLM\%SPPk%\%_oApp%" /f
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10"
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)
if %winbuild% GEQ 9600 (
reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10"
reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%OPPk%\%_oA14%" /f
reg delete "HKLM\%OPPk%\%_oApp%" /f
goto :eof

:=========================================================================================================================================

:_Complete_Uninstall

cls
mode con: cols=91 lines=30
title Online KMS Complete Uninstall %masver%

set "key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks"

set "_C16R="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath /reg:32" 2^>nul') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
if %winbuild% GEQ 9200 if defined _C16R (
echo.
echo ## Notice ##
echo.
echo To make sure Office programs do not show a non-genuine banner,
echo please run the activation option once, and don't uninstall afterward.
echo __________________________________________________________________________________________
)

set error_=
echo.
call :Clear-KMS-Cache
call :clearstuff

if defined error_ (
if [%error_%]==[1] (
echo __________________________________________________________________________________________
%eline%
echo Try Again / Restart the System
echo __________________________________________________________________________________________
)
) else (
echo __________________________________________________________________________________________
echo.
call :_color %Green% "Online KMS Complete Uninstall was done successfully."
echo __________________________________________________________________________________________
)

if defined _unattended timeout /t 2 & exit /b

echo.
call :_color %_Yellow% "Press any key to go back..."
pause >nul
exit /b

:clearstuff

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul && (
echo Deleting [Task] Activation-Renewal
schtasks /delete /tn Activation-Renewal /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul && (
echo Deleting [Task] Activation-Run_Once
schtasks /delete /tn Activation-Run_Once /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Renewal" >nul && (
echo Deleting [Task] Online_KMS_Activation_Script-Renewal
schtasks /delete /tn Online_KMS_Activation_Script-Renewal /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Run_Once" >nul && (
echo Deleting [Task] Online_KMS_Activation_Script-Run_Once
schtasks /delete /tn Online_KMS_Activation_Script-Run_Once /f %nul%
)

If exist "%windir%\Online_KMS_Activation_Script\" (
echo Deleting [Folder] %windir%\Online_KMS_Activation_Script\
rmdir /s /q "%windir%\Online_KMS_Activation_Script\" %nul%
)

if exist "%ProgramData%\Online_KMS_Activation.cmd" (
echo Deleting [File] %ProgramData%\Online_KMS_Activation.cmd
del /f /q "%ProgramData%\Online_KMS_Activation.cmd" %nul%
)

If exist "%ProgramData%\Online_KMS_Activation\" (
echo Deleting [Folder] %ProgramData%\Online_KMS_Activation\
rmdir /s /q "%ProgramData%\Online_KMS_Activation\" %nul%
)

If exist "%ProgramData%\Activation-Renewal\" (
echo Deleting [Folder] %ProgramData%\Activation-Renewal\
rmdir /s /q "%ProgramData%\Activation-Renewal\" %nul%
)

If exist "%ProgramFiles%\Activation-Renewal\" (
echo Deleting [Folder] %ProgramFiles%\Activation-Renewal\
rmdir /s /q "%ProgramFiles%\Activation-Renewal\" %nul%
)

reg query "HKCR\DesktopBackground\shell\Activate Windows - Office" %nul% && (
echo Deleting [Registry] HKCR\DesktopBackground\shell\Activate Windows - Office
Reg delete "HKCR\DesktopBackground\shell\Activate Windows - Office" /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul && (set error_=1)
reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul && (set error_=1)
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script" >nul && (set error_=1)
If exist "%windir%\Online_KMS_Activation_Script\" (set error_=1)
reg query "HKCR\DesktopBackground\shell\Activate Windows - Office" %nul% && (set error_=1)
if exist "%ProgramData%\Online_KMS_Activation.cmd" (set error_=1)
if exist "%ProgramData%\Online_KMS_Activation\" (set error_=1)
if exist "%ProgramData%\Activation-Renewal\" (set error_=1)
if exist "%ProgramFiles%\Activation-Renewal\" (set error_=1)
exit /b

:=========================================================================================================================================

:RenTask

cls
mode con cols=91 lines=30
title  Install Activation Auto-Renewal %masver%

set error_=
set "_dest=%ProgramFiles%\Activation-Renewal"
set "key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks"

call :clearstuff %nul%

if defined error_ (
%eline%
echo Failed to completely clear KMS related folders/tasks.
echo Run the Uninstall option and then try again.
goto :RenDone
)

if not exist "%_dest%\" md "%_dest%\" %nul%
set "_temp=%SystemRoot%\Temp\_taskwork_%Random%"

set nil=
if exist "%_temp%\.*" rmdir /s /q "%_temp%\" %nul%
md "%_temp%\" %nul%
call :RenExport renewal "%_temp%\Renewal.xml" Unicode
if defined ActTask (call :RenExport run_once "%_temp%\Run_Once.xml" Unicode)
s%nil%cht%nil%asks /cre%nil%ate /tn "Activation-Renewal" /ru "SYS%nil%TEM" /xml "%_temp%\Renewal.xml" %nul%
if defined ActTask (s%nil%cht%nil%asks /cre%nil%ate /tn "Activation-Run_Once" /ru "SYS%nil%TEM" /xml "%_temp%\Run_Once.xml" %nul%)
if exist "%_temp%\.*" rmdir /s /q "%_temp%\" %nul%

call :createInfo.txt
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":_extracttask\:.*`r`n\"; [io.file]::WriteAllText('%_dest%\Activation_task.cmd', '@REM Dummy ' + '%random%' + [Environment]::NewLine + $f[1].Trim(), [System.Text.Encoding]::ASCII);"
title  Install Activation Auto-Renewal %masver%

::========================================================================================================================================

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul || (set error_=1)
if defined ActTask reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul || (set error_=1)

If not exist "%_dest%\Activation_task.cmd" (set error_=1)
If not exist "%_dest%\Info.txt" (set error_=1)

if defined error_ (

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul && (
schtasks /delete /tn Activation-Renewal /f %nul%
)
reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul && (
schtasks /delete /tn Activation-Run_Once /f %nul%
)

If exist "%_dest%\" (
rmdir /s /q "%_dest%\" %nul%
)

%eline%
echo Run the Uninstall option and then try again.
goto :RenDone
)

echo __________________________________________________________________________________________
echo.
echo Files created:
echo %_dest%\Activation_task.cmd
echo %_dest%\Info.txt
echo.
(if defined ActTask (echo Scheduled Tasks created:) else (echo Scheduled Task created:))
echo \Activation-Renewal [Weekly]
if defined ActTask (echo \Activation-Run_Once)
echo __________________________________________________________________________________________
echo.
echo Info:
echo Activation will be renewed every week if the Internet connection is found.
echo It'll only renew installed KMS licenses. It won't convert any license to KMS.
echo __________________________________________________________________________________________
echo.
if defined ActTask (
call :_color %Green% "Renewal and Activation Tasks were successfully created."
) else (
call :_color %Green% "Renewal Task was successfully created."
)
echo.
call :_color %Gray% "Make sure you have run the Activation option at least once."
echo __________________________________________________________________________________________
)

::========================================================================================================================================

:RenDone

if defined _unattended exit /b

echo.
call :_color %_Yellow% "Press any key to go back..."
pause >nul
exit /b

::========================================================================================================================================

:createInfo.txt

(
echo   The use of this script is to renew your Windows/Office KMS license using online KMS.
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
echo   Online KMS Activation Script is a part of 'Microsoft_Activation_Scripts' [MAS] project.
echo:   
echo   Homepage: mass grave[.]dev
echo      Email: windowsaddict@protonmail.com
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

::  Extract the text from batch script without character and file encoding issue

:RenExport

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('%~2',$f[1].Trim(),[System.Text.Encoding]::%~3);"
exit /b

::========================================================================================================================================

:_extracttask:
@echo off

::   Renew K-M-S activation with Online servers via scheduled task

::============================================================================
::
::   This script is a part of 'Microsoft_Activation_Scripts' (MAS) project.
::
::   Homepage: mass grave[.]dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================


if not "%~1"=="Task" (
echo.
echo ====== Error ======
echo.
echo This file is supposed to be run only by the scheduled task.
echo.
echo Press any key to exit
pause >nul
exit /b
)

::  Set Path variable, it helps if it is misconfigured in the system

set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "PATH=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%PATH%"
)

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
wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul && set _wmic=1
)

setlocal EnableDelayedExpansion
if exist "%ProgramFiles%\Activation-Renewal\" call :_taskstart>>"%ProgramFiles%\Activation-Renewal\Logs.txt" & exit

::========================================================================================================================================

:_taskstart

echo.
echo %date%, %time%

set /a loop=1
set /a max_loop=4

call :_tasksetserv

:_intrepeat

::  Check Internet connection. Works even if ICMP echo is disabled.

for %%a in (%srvlist%) do (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (
if not [%%#]==[] goto _taskIntConnected
)
)

nslookup dns.msftncsi.com 2>nul | find "131.107.255.255" 1>nul
if [%errorlevel%]==[0] goto _taskIntConnected

if %loop%==%max_loop% (
set _tserror=1
goto _taskend
)

echo.
echo Error: Internet is not connected
echo Waiting 30 seconds

timeout /t 30 >nul
set /a loop=%loop%+1
goto _intrepeat

:_taskIntConnected

::========================================================================================================================================

::  Check not x86 Windows

set notx86=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b
if /i not "%arch%"=="x86" set notx86=1

::========================================================================================================================================

set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"

set "slp=SoftwareLicensingProduct"
set "ospp=OfficeSoftwareProtectionProduct"

set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"

::========================================================================================================================================

::  Clean existing KMS cache from the registry / Set port value to 1688

%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
%nul% reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if defined notx86 (
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
)
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)
if %winbuild% GEQ 9600 (
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
%nul% reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
%nul% reg delete "HKLM\%OPPk%\%_oA14%" /f
%nul% reg delete "HKLM\%OPPk%\%_oApp%" /f

::========================================================================================================================================

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
echo.
echo Error: Not Respoding- !e_wmispp!
echo.
)

::========================================================================================================================================

::  Check installed volume products activation ID's

call :_taskgetids sppwid %slp% windows
call :_taskgetids sppoid %slp% office
call :_taskgetids osppid %ospp% office

::========================================================================================================================================

echo.
echo Renewing KMS activation for all installed Volume products

if not defined sppwid if not defined sppoid if not defined osppid (
echo.
echo No installed Volume Windows / Office product found
echo.
echo Renewing KMS server
call :_taskgetserv
call :_taskregserv
goto :_skipact
)

::========================================================================================================================================

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

:: Set specific KMS host to Local Host so that global KMS IP can not replace KMS38 activation but can be used with Office and other Windows Editions.

if %_kms38% EQU 1 (
%nul% reg add "HKLM\%SPPk%\%_wApp%\%sppwid%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2"
%nul% reg add "HKLM\%SPPk%\%_wApp%\%sppwid%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)

::========================================================================================================================================

echo.
if defined sppwid (
set _path=%slp%
set _actid=%sppwid%
call :_actprod
call :_act act_win
call :_actinfo act_win
) else (
echo Checking: Volume version of Windows is not installed
)

if defined sppoid (
set _path=%slp%
for %%# in (%sppoid%) do (
echo.
set _actid=%%#
call :_actprod
call :_act
call :_actinfo
)
)

if defined osppid (
set _path=%ospp%
for %%# in (%osppid%) do (
echo.
set _actid=%%#
call :_actprod
call :_act
call :_actinfo
)
)

if not defined sppoid if not defined osppid (
echo.
echo Checking: Volume version of Office is not installed
)

:_skipact

::========================================================================================================================================

if defined run_once (
echo.
echo Deleting Scheduled Task Activation-Run_Once
schtasks /delete /tn Activation-Run_Once /f %nul%
)

::========================================================================================================================================

:_taskend

echo.
echo Exiting
echo ______________________________________________________________________

if defined _tserror (exit /b 123456789) else (exit /b 0)

::========================================================================================================================================

:_act

set errorcode=12345
set /a act_attempt=0

:_act2

if %act_attempt% GTR 4 exit /b

if not [%act_ok%]==[1] (
call :_taskgetserv
call :_taskregserv
)

if not !server_num! GTR %max_servers% (

if [%1]==[act_win] if %_kms38% EQU 1 (
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
if [%1]==[act_win] if !errorcode! EQU -1073418187 if %winbuild% LSS 9200 (
set act_ok=1
exit /b
)

set act_ok=0
set /a act_attempt+=1
goto _act2
)
exit /b

:_actprod

if %_wmic% EQU 1 for /f "tokens=2 delims==" %%x in ('"wmic path !_path! where ID='!_actid!' get Name /VALUE" 2^>nul') do call echo Activating: %%x
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%x in ('%psc% "(([WMISEARCHER]'SELECT Name FROM !_path! WHERE ID=''!_actid!''').Get()).Name | %% {echo ('Name='+$_)}" 2^>nul') do call echo Activating: %%x
exit /b

::========================================================================================================================================

:_actinfo

if [%1]==[act_win] if %_kms38% EQU 1 (
echo Windows is activated with KMS38
exit /b
)

if %errorcode% EQU 12345 (
echo Product Activation Failed
echo Unable to test KMS servers due to restricted or no Internet
set _tserror=1
exit /b
)

if %errorcode% EQU -1073418187 (
echo Product Activation Failed: 0xC004F035
if [%1]==[act_win] if %winbuild% LSS 9200 echo Windows 7 cannot be KMS-activated on this computer due to unqualified OEM BIOS
exit /b
)

if %errorcode% EQU -1073417728 (
echo Product Activation Failed: 0xC004F200
echo Windows needs to rebuild the activation-related files.
set _tserror=1
exit /b
)

set gpr=0
set gpr2=0
call :_taskgetgrace
set /a "gpr2=(%gpr%+1440-1)/1440"

if %errorcode% EQU 0 if %gpr% EQU 0 (
echo Product Activation succeeded, but Remaining Period failed to increase.
if [%1]==[act_win] if %winbuild% LSS 9200 echo This could be related to the error described in KB4487266
set _tserror=1
exit /b
)

set _actpass=1
if %gpr% EQU 43200  if [%1]==[act_win] if %winbuild% GEQ 9200 set _actpass=0
if %gpr% EQU 64800  set _actpass=0
if %gpr% GTR 259200 if [%1]==[act_win] call :_taskchkEnterpriseG _actpass
if %gpr% EQU 259200 set _actpass=0

if %errorcode% EQU 0 if %_actpass% EQU 0 (
echo Product Activation Successful
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
exit /b
)

cmd /c exit /b %errorcode%
if %errorcode% NEQ 0 (
echo Product Activation Failed: 0x!=ExitCode!
) else (
echo Product Activation Failed
)
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
set _tserror=1
exit /b

::========================================================================================================================================

:_taskgetids

set %1=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %2 where (Name like '%%%3%%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get ID /VALUE" 2^>nul')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %2 WHERE Name like ''%%%3%%'' and Description like ''%%KMSCLIENT%%'' and PartialProductKey is not NULL').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
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

:_taskregserv

%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"
%nul% reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"

if %winbuild% GEQ 9200 (
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"
if defined notx86 (
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32
)
)
exit /b

::========================================================================================================================================

:_tasksetserv

::  Multi KMS servers integration and servers randomization

set srvlist=
set -=

set "srvlist=kms.zhu%-%xiaole.org kms-default.cangs%-%hui.net kms.six%-%yin.com kms.moe%-%club.org kms.cgt%-%soft.com"
set "srvlist=%srvlist% kms.id%-%ina.cn kms.moe%-%yuuko.com xinch%-%eng213618.cn kms.wl%-%rxy.cn kms.ca%-%tqu.com"
set "srvlist=%srvlist% kms.0%-%t.net.cn kms.its%-%jzx.com kms.wx%-%lost.com kms.moe%-%yuuko.top kms.gh%-%pym.com"

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

::  Get IPv4 address of KMS server to use for the activation, works even if ICMP echo is disabled.
::  Microsoft and Antivirus's may flag the issue if public KMS server host name is directly used for the activation.

set /a server_num+=1
(for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%a"
if [%KMS_IP%]==[!KMS_IP!] for /f "delims=[] tokens=2" %%# in ('pathping -4 -h 1 -n -p 1 -q 1 -w 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%#"
if not [%KMS_IP%]==[!KMS_IP!] exit /b
goto :_taskgetserv
)

:: Ver:1.9
::========================================================================================================================================
:_extracttask:

:======================================================================================================================================================

:_color

if %_NCS% EQU 1 (
if defined _unattended (echo %~2) else (echo %esc%[%~1%~2%esc%[0m)
) else (
if defined _unattended (echo %~2) else (call :batcol %~1 "%~2")
)
exit /b

:_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
call :batcol %~1 "%~2" %~3 "%~4"
)
exit /b

::=======================================

:: Colored text with pure batch method
:: Thanks to @dbenham and @jeb
:: stackoverflow.com/a/10407642

:batcol

pushd %_coltemp%
if not exist "'" (<nul >"'" set /p "=.")
setlocal
set "s=%~2"
set "t=%~4"
call :_batcol %1 s %3 t
del /f /q "'"
del /f /q "`.txt"
popd
exit /b

:_batcol

setlocal EnableDelayedExpansion
set "s=!%~2!"
set "t=!%~4!"
for /f delims^=^ eol^= %%i in ("!s!") do (
  if "!" equ "" setlocal DisableDelayedExpansion
    >`.txt (echo %%i\..\')
    findstr /a:%~1 /f:`.txt "."
    <nul set /p "=%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%"
)
if "%~4"=="" echo(&exit /b
setlocal EnableDelayedExpansion
for /f delims^=^ eol^= %%i in ("!t!") do (
  if "!" equ "" setlocal DisableDelayedExpansion
    >`.txt (echo %%i\..\')
    findstr /a:%~3 /f:`.txt "."
    <nul set /p "=%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%"
)
echo(
exit /b

::=======================================

:_colorprep

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"

set     "Red="41;97m""
set    "Gray="100;97m""
set   "Black="30m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "Yellow="43;97m""
set "Magenta="45;97m""

set    "_Red="40;91m""
set  "_Green="40;92m""
set   "_Blue="40;94m""
set  "_White="40;37m""
set "_Yellow="40;93m""

exit /b
)

for /f %%A in ('"prompt $H&for %%B in (1) do rem"') do set "_BS=%%A %%A"
set "_coltemp=%SystemRoot%\Temp"

set     "Red="CF""
set    "Gray="8F""
set   "Black="00""
set   "Green="2F""
set    "Blue="1F""
set  "Yellow="6F""
set "Magenta="5F""

set    "_Red="0C""
set  "_Green="0A""
set   "_Blue="09""
set  "_White="07""
set "_Yellow="0E""

exit /b

::========================================================================================================================================

::  https://gist.github.com/ave9858/9fff6af726ba3ddc646285d1bbf37e71
::  This code is used to clean Office licenses

:cleanlicense:
function UninstallLicenses($DllPath) {
    $AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
    $TypeBuilder = $ModuleBuilder.DefineType(0)
    
    [void]$TypeBuilder.DefinePInvokeMethod('SLOpen', $DllPath, 'Public, Static', 1, [int], @([IntPtr].MakeByRefType()), 1, 3)
    [void]$TypeBuilder.DefinePInvokeMethod('SLGetSLIDList', $DllPath, 'Public, Static', 1, [int],
        @([IntPtr], [int], [Guid].MakeByRefType(), [int], [int].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
    [void]$TypeBuilder.DefinePInvokeMethod('SLUninstallLicense', $DllPath, 'Public, Static', 1, [int], @([IntPtr], [IntPtr]), 1, 3)

    $SPPC = $TypeBuilder.CreateType()
    $Handle = 0
    [void]$SPPC::SLOpen([ref]$Handle)
    $pnReturnIds = 0
    $ppReturnIds = 0

    if (!$SPPC::SLGetSLIDList($Handle, 0, [ref][Guid]"0ff1ce15-a989-479d-af46-f275c6370663", 6, [ref]$pnReturnIds, [ref]$ppReturnIds)) {
        foreach ($i in 0..($pnReturnIds - 1)) {
            [void]$SPPC::SLUninstallLicense($Handle, [Int64]$ppReturnIds + [Int64]16 * $i)
        }    
    }
}

$OSPP = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -ErrorAction SilentlyContinue).Path
if ($OSPP) {
    Write-Output "Found Office Software Protection installed, cleaning"
    UninstallLicenses($OSPP + "osppc.dll")
}
UninstallLicenses("sppc.dll")
:cleanlicense:

::========================================================================================================================================
:: Leave empty line below
