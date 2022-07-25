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
echo Press any key to exit...
pause >nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title  Change Windows Edition

set _elev=
if /i "%~1"=="-el" set _elev=1

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
set "line=echo ___________________________________________________________________________________________"

::========================================================================================================================================

if %winbuild% LSS 10240 (
%eline%
echo Unsupported OS version detected.
echo Project is supported for Windows 10/11/Server Build 10240 and later.
goto ced_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto ced_done
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
goto ced_done
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul% reg query HKU\S-1-5-19 || (
if not defined _elev %nul% %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto ced_done
)

::========================================================================================================================================

cls
mode 98, 30

call :dk_initial

if not defined applist (
cls
%eline%
echo Not Respoding: !e_wmispp!
goto ced_done
)

::========================================================================================================================================

::  Check Windows Edition

set osedition=
for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition 2^>nul ^| find /i "Current Edition :"') do set "osedition=%%a"

cls
if "%osedition%"=="" (
%eline%
DISM /English /Online /Get-CurrentEdition %nul%
cmd /c exit /b !errorlevel!
echo DISM command failed [Error Code - 0x!=ExitCode!]
echo OS Edition was not detected properly. Aborting...
goto ced_done
)

::  Check product name

call :dk_product

::  Check SKU value

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
goto ced_done
)

::  Check PowerShell

%psc% $ExecutionContext.SessionState.LanguageMode 2>nul | find /i "Full" 1>nul || (
%eline%
echo PowerShell is not responding properly. Aborting...
goto ced_done
)

::  Check slmgr /dlv

cscript //nologo %windir%\system32\slmgr.vbs /dlv %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 set "error_code=[0x%=ExitCode%]"
if %error_code% NEQ 0 (
%eline%
echo slmgr /dlv is not responding %error_code%
goto ced_done
)

::========================================================================================================================================

::  Get Target editions list

set _target=
set _ntarget=
for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (if defined _target (set "_target=!_target! %%a") else (set "_target=%%a"))

::========================================================================================================================================

::  Block the change to/from CountrySpecific and CloudEdition editions

for %%# in (99 139 202 203) do if %osSKU%==%%# (
%eline%
echo [%winos% ^| SKU:%osSKU% ^| %winbuild%]
echo It's not recommended to change this installed edition to any other.
echo Aborting...
goto ced_done
)

if defined _target (
for %%# in (%_target%) do (
echo %%# | findstr /i "CountrySpecific CloudEdition" %nul% || (if defined _ntarget (set "_ntarget=!_ntarget! %%#") else (set "_ntarget=%%#"))
)
)

if not defined _ntarget (
%line%
echo:
call :dk_color %Gray% "Target Edition not found."
echo Current Edition [%osedition% ^| %winbuild%] can not be changed to any other Edition.
%line%
goto ced_done
)

::========================================================================================================================================

:cedmenu2

cls
mode 98, 30
set inpt=
set counter=0
set verified=0
set targetedition=

%line%
echo:
call :dk_color %Gray% "You can change the Current Edition [%osedition%] to one of the following."
%line%
echo:

for %%A in (%_ntarget%) do (
set /a counter+=1
echo [!counter!]  %%A
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  Exit
echo:
call :dk_color %_Green% "Enter option number in keyboard, and press "Enter":"
set /p inpt=
if "%inpt%"=="" goto cedmenu2
if "%inpt%"=="0" exit /b
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto cedmenu2

::========================================================================================================================================

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
goto :ced_change_server
)

cls
set key=
set _chan=
set _changepk=0
set "keyflow=Retail Volume:MAK Volume:GVLK OEM:NONSLP OEM:DM"

::  Check if changepk.exe or slmgr.vbs is required for edition upgrade

if not exist "%SystemRoot%\System32\spp\tokens\skus\%targetedition%\" (
set _changepk=1
)

if /i "%osedition:~0,4%"=="Core" (
if /i not "%targetedition:~0,4%"=="Core" (
set _changepk=1
)
)

if %winbuild% LEQ 19044 call :changeeditiondata

if not defined key call :ced_targetSKU %targetedition%
if not defined key if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Unable to get product key from pkeyhelper.dll
echo Make sure you are using updated version of the script
goto ced_done
)

::========================================================================================================================================

%line%

::  Changing from Core to Non-Core & Changing editions in Windows build older than 17134 requires "changepk /productkey" method and restart
::  In other cases, editions can be changed instantly with "slmgr /ipk"

:ced_loop

cls
if %_changepk%==1 (
echo "%_chan%" | find /i "OEM" >NUL && (
%eline%
echo [%osedition%] can not be changed to [%targetedition%] Edition due to lack of non OEM keys.
echo Non-OEM keys are required to change from Core to Non-Core Editions.
goto ced_done
)
for %%a in (dns.msftncsi.com,www.microsoft.com,one.one.one.one,resolver1.opendns.com) do (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (
if not [%%#]==[] (
%eline%
echo Disconnect the Internet and then press any key...
pause >nul
goto ced_loop
)
)
)
)

echo:
echo Changing the Current Edition [%osedition%] to [%targetedition%]
echo:

if %_changepk%==1 (
call :dk_color %Green% "You can safely ignore if error appears in the upgrade Window."
call :dk_color %Red% "But in that case you must manually reboot the system."
echo:
call :dk_color %Magenta% "Important - Save your work before continue, system will auto reboot."
echo:
choice /C:21 /N /M "[1] Continue [2] Exit : "
if !errorlevel!==1 exit /b
)

::========================================================================================================================================

if %_changepk%==0 (
echo Installing %_chan% Key [%key%]
echo:
if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%key%')" %nul%
if not !errorlevel!==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%

set error_code=!errorlevel!
cmd /c exit /b !error_code!
if !error_code! NEQ 0 set "error_code=[0x!=ExitCode!]"

if !error_code! EQU 0 (
call :dk_refresh
call :dk_color %Green% "[Successful]"
echo:
call :dk_color %Gray% "Reboot is required to properly change the Edition."
) else (
call :dk_color %Red% "[Unsuccessful] [Error Code: 0x!=ExitCode!]"
)
)

if %_changepk%==1 (
echo:
echo Applying the command with %_chan% Key
echo start changepk.exe /ProductKey %key%
start changepk.exe /ProductKey %key%
)
%line%

goto ced_done

::========================================================================================================================================

:ced_change_server

cls
mode con cols=105 lines=32
%nul% %psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

set key=
set pkeychannel=
set "keyflow=Volume:GVLK Retail Volume:MAK OEM:NONSLP OEM:DM"
call :changeeditionserverdata

if not defined key call :ced_targetSKU %targetedition%
if not defined key if defined targetSKU call :ced_windowskey
if defined key if not defined pkeychannel call :dk_pkeychannel %key%

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Unable to get product key from pkeyhelper.dll
echo Make sure you are using updated version of the script
goto ced_done
)

::========================================================================================================================================

cls
echo:
echo Changing the Current Edition [%osedition%] to [%targetedition%]
echo:
echo Applying the command with %pkeychannel% Key
echo DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula
DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula

call :dk_color %Magenta% "Make sure to restart the system."

::========================================================================================================================================

:ced_done

echo:
call :dk_color %_Yellow% "Press any key to exit..."
pause >nul
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

:dk_initial

echo:
echo Initializing...

::  Check and enable WinMgmt, sppsvc services if required

for %%# in (WinMgmt sppsvc) do (
for /f "skip=2 tokens=2*" %%a in ('reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v Start 2^>nul') do if /i %%b NEQ 0x2 (
echo:
echo Enabling %%# service...
if /i %%#==sppsvc  sc config %%# start= delayed-auto %nul% || echo Failed
if /i %%#==WinMgmt sc config %%# start= auto %nul% || echo Failed
)
sc start %%# %nul%
if !errorlevel! NEQ 1056 if !errorlevel! NEQ 0 (
echo:
echo Starting %%# service...
sc start %%#
echo:
call :dk_color %Red% "Failed to start [%%#] service, rest of the process may take a long time..."
)
)

::  Check WMI and SPP Errors

call :dk_ckeckwmic

set e_wmi=
set e_wmispp=
call :dk_actids

if not defined applist (
net stop sppsvc /y %nul%
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if !errorlevel! NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
call :dk_refresh

if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" 2>nul | find /i "computersystem" 1>nul
if !errorlevel! NEQ 0 set e_wmi=1

if defined e_wmi (set e_wmispp=WMI, SPP) else (set e_wmispp=SPP)
call :dk_actids
)
exit /b

::========================================================================================================================================

::  Get Product Key from pkeyhelper.dll for future new editions
::  It works on Windows 10 1803 (17134) and later builds.

:dk_pkey

set pkey=
set d1=[DllImport(\"pkeyhelper.dll\",CharSet=CharSet.Unicode)]public static extern int SkuGetProductKeyForEdition(int e, string c, out string k, out string p);
set d2=$AP=Add-Type -Member '%d1%' -Name D1 -PassThru; $k=''; $null=$AP::SkuGetProductKeyForEdition(%1, %2, [ref]$k, [ref]$null); $k
for /f %%a in ('%psc% "%d2%"') do if not errorlevel 1 (set pkey=%%a)
exit /b

::  Get channel name for the key which was extracted from pkeyhelper.dll

:dk_pkeychannel

set k=%1
set pkeychannel=
set p=%SystemRoot%\System32\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms
set m=[System.Runtime.InteropServices.Marshal]
set d1=[DllImport(\"PidGenX.dll\",CharSet=CharSet.Unicode)]public static extern int PidGenX(string k,string p,string m,int u,IntPtr i,IntPtr d,IntPtr f);
set d2=$AP=Add-Type -Member '%d1%' -Name D1 -PassThru; $k='%k%'; $p='%p%'; $r=[byte[]]::new(0x04F8); $r[0]=0xF8; $r[1]=0x04; $f=%m%::AllocHGlobal(1272); %m%::Copy($r,0,$f,1272);
set d3=%d2% [void]$AP::PidGenX($k,$p,\"00000\",0,0,0,$f); %m%::Copy($f,$r,0,1272); %m%::FreeHGlobal($f); [System.Text.Encoding]::Unicode.GetString($r, 1016, 128).Replace('0','')
for /f %%a in ('%psc% "%d3%"') do if not errorlevel 1 (set pkeychannel=%%a)
exit /b

:ced_windowskey

for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b
for %%# in (%keyflow%) do (
call :dk_pkey %targetSKU% '%%#'
if defined pkey call :dk_pkeychannel !pkey!
if /i [!pkeychannel!]==[%%#] (
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
set d1=[DllImport(\"pkeyhelper.dll\",CharSet=CharSet.Unicode)]public static extern int GetEditionIdFromName(string e, out int s);
set d2=$AP=Add-Type -Member '%d1%' -Name D1 -PassThru; $s=0; $null=$AP::GetEditionIdFromName('%k%', [ref]$s); $s
for /f %%a in ('%psc% "%d2%"') do if not errorlevel 1 (set targetSKU=%%a)
if "%targetSKU%"=="0" set targetSKU=
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

::  1st column = Generic Retail/OEM/MAK/GVLK Key
::  2nd column = Key Type
::  3rd column = WMI Edition ID
::  4th column = Version name incase same Edition ID is used in different OS versions with different key
::  Separator  = _

::  Key preference is in the following order. Retail > Volume:MAK > Volume:GVLK > OEM:NONSLP > OEM:DM
::  OEM keys are in last because they can't be used in edition change if "changepk /productkey" method is needed instead of "slmgr /ipk"
::  OEM keys are listed here because we don't have other keys for that edition

:changeeditiondata

for %%# in (
44NYX-TKR9D-CCM2D-V6B8F-HQWWR_Volume:MAK_Enterprise
D6RD9-D4N8T-RT9QX-YW6YT-FCWWJ_____Retail_Starter
3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT_Volume:MAK_EnterpriseN
3NFXW-2T27M-2BDW6-4GHRV-68XRX_____Retail_StarterN
VK7JG-NPHTM-C97JM-9MPGT-3V66T_____Retail_Professional
2B87N-8KFHP-DKV6R-Y2C8J-PKCKT_____Retail_ProfessionalN
4CPRK-NM3K3-X6XXQ-RXX86-WXCHW_____Retail_CoreN
N2434-X9D7W-8PF6X-8DV9T-8TYMD_____Retail_CoreCountrySpecific
BT79Q-G7N6G-PGBYW-4YWX6-6F4BT_____Retail_CoreSingleLanguage
YTMG3-N6DKC-DKB77-7M9GH-8HVX7_____Retail_Core
XKCNC-J26Q9-KFHD2-FKTHY-KD72Y_OEM:NONSLP_PPIPro
YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY_____Retail_Education
84NGF-MHBT6-FXBX8-QWJK7-DRR8H_____Retail_EducationN
KCNVH-YKWX8-GJJB9-H9FDT-6F7W2_Volume:MAK_EnterpriseS_2021
VBX36-N7DDY-M9H62-83BMJ-CPR42_Volume:MAK_EnterpriseS_2019
PN3KR-JXM7T-46HM4-MCQGK-7XPJQ_Volume:MAK_EnterpriseS_2016
DVWKN-3GCMV-Q2XF4-DDPGM-VQWWY_Volume:MAK_EnterpriseS_2015
RQFNW-9TPM3-JQ73T-QV4VQ-DV9PT_Volume:MAK_EnterpriseSN_2021
M33WV-NHY3C-R7FPM-BQGPT-239PG_Volume:MAK_EnterpriseSN_2019
2DBW3-N2PJG-MVHW3-G7TDK-9HKR4_Volume:MAK_EnterpriseSN_2016
NTX6B-BRYC2-K6786-F6MVQ-M7V2X_Volume:MAK_EnterpriseSN_2015
G3KNM-CHG6T-R36X3-9QDG6-8M8K9_____Retail_ProfessionalSingleLanguage
HNGCC-Y38KG-QVK8D-WMWRK-X86VK_____Retail_ProfessionalCountrySpecific
DXG7C-N36C4-C4HTG-X4T3X-2YV77_____Retail_ProfessionalWorkstation
WYPNQ-8C467-V2W6J-TX4WX-WT2RQ_____Retail_ProfessionalWorkstationN
8PTT6-RNW4C-6V7J2-C2D3X-MHBPB_____Retail_ProfessionalEducation
GJTYN-HDMQY-FRR76-HVGC7-QPF8P_____Retail_ProfessionalEducationN
C4NTJ-CX6Q2-VXDMR-XVKGM-F9DJC_Volume:MAK_EnterpriseG
46PN6-R9BK9-CVHKB-HWQ9V-MBJY8_Volume:MAK_EnterpriseGN
NJCF7-PW8QT-3324D-688JX-2YV66_____Retail_ServerRdsh
V3WVW-N2PV2-CGWC3-34QGF-VMJ2C_____Retail_Cloud
NH9J3-68WK7-6FB93-4K3DF-DJ4F6_____Retail_CloudN
2HN6V-HGTM8-6C97C-RK67V-JQPFD_____Retail_CloudE
XQQYW-NFFMW-XJPBH-K8732-CKFFD_____OEM:DM_IoTEnterprise
QPM6N-7J2WJ-P88HH-P3YRH-YY74H_OEM:NONSLP_IoTEnterpriseS
K9VKN-3BGWV-Y624W-MCRMQ-BHDCD_____Retail_CloudEditionN
KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W_____Retail_CloudEdition
) do (
for /f "tokens=1-4 delims=_" %%A in ("%%#") do if /i %targetedition%==%%C (

if not defined key (
set 4th=%%D
if not defined 4th (
set "key=%%A" & set "_chan=%%B"
) else (
echo "%winos%" | find "%%D" 1>nul && (set "key=%%A" & set "_chan=%%B")
)
)
)
)
exit /b

::========================================================================================================================================

:changeeditionserverdata

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" (set Cor=Cor) else (set Cor=)
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildBranch 2^>nul') do set "branch=%%b"

::  Only RS3 and older version keys (GVLK/Generic Retail) are stored here, later ones are extracted from the system itself

for %%# in (
WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY_RS1_ServerStandard%Cor%
CB7KF-BWN84-R7R2Y-793K2-8XDDG_RS1_ServerDatacenter%Cor%
JCKRF-N37P4-C2D82-9YXRT-4M63B_RS1_ServerSolution
QN4C6-GBJD2-FB422-GHWJK-GJG2R_RS1_ServerCloudStorage
VP34G-4NPPG-79JTQ-864T4-R3MQX_RS1_ServerAzureCor
9JQNQ-V8HQ6-PKB8H-GGHRY-R62H6_RS1_ServerAzureNano
VN8D3-PR82H-DB6BJ-J9P4M-92F6J_RS1_ServerStorageStandard
48TQX-NVK3R-D8QR3-GTHHM-8FHXC_RS1_ServerStorageWorkgroup
2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG_RS3_ServerDatacenterACor
PTXN8-JFHJM-4WC78-MPCBR-9W4KR_RS3_ServerStandardACor
) do (
for /f "tokens=1-3 delims=_" %%A in ("%%#") do if /i %targetedition%==%%C (
echo "%branch%" | find /i "%%B" 1>nul && (set "key=%%A")
)
)
exit /b

::========================================================================================================================================