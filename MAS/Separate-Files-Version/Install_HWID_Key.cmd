@setlocal DisableDelayedExpansion
@echo off



::============================================================================
::
::   This script is a part of 'Microsoft_Activation_Scripts' (MAS) project.
::
::   Homepage: massgrave.dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================



:: For unattended mode, run the script with "/Insert-HWID-Key" parameter



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
title  Install Windows HWID Key

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="-el" set _elev=1
if /i "%%A"=="/Insert-HWID-Key"  set _unattended=1
)
)

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
set   "Green="42;97m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set   "Green="DarkGreen" "white""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
set "line=echo ___________________________________________________________________________________________"
if %~z0 GEQ 200000 (set "_exitmsg=Go back") else (set "_exitmsg=Exit")

::========================================================================================================================================

if %winbuild% LSS 10240 (
%eline%
echo Unsupported OS version detected.
echo This option is supported only for Windows 10/11.
goto ins_done
)

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo HWID Activation is not supported for Windows Server.
goto ins_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto ins_done
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
goto ins_done
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

>nul fltmc || (
if not defined _elev %nul% %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script require admin privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto ins_done
)

::========================================================================================================================================

cls
mode 98, 30
echo:
echo Initializing...
call :dk_product
call :dk_ckeckwmic
call :dk_actids

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
goto ins_done
)

::========================================================================================================================================

::  Detect key

set key=
set channel=
set actidnotfound=

for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildBranch 2^>nul') do set "branch=%%b"

if defined applist call :hwidkey key attempt1
if not defined key call :hwidkey key attempt2

if not defined key (
%eline%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
echo Unable to find this product in the HWID supported product list.
echo Make sure you are using updated version of the script.
echo https://massgrave.dev
goto ins_done
)

::========================================================================================================================================

if %_unattended%==1 goto insertkey

cls
%line%
echo:
echo Install [%winos% ^| SKU:%osSKU% ^| %winbuild%] %channel% Key
echo [%key%]
%line%
echo:
if not "%regSKU%"=="%wmiSKU%" (
echo Note: Difference Found In SKU Value- WMI:%wmiSKU% Reg:%regSKU%
echo:
)
call :dk_color %_Green% "Press [1] to Continue or [0] to %_exitmsg%"
choice /C:01 /N
if %errorlevel%==1 exit /b

::========================================================================================================================================

:insertkey

cls
%line%

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%key%')" %nul%
if not %errorlevel%==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%

set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 set "error_code=[0x%=ExitCode%]"

echo:
echo [%winos% ^| SKU:%osSKU% ^| %winbuild%]
echo Installing %channel% [%key%]
echo:

if %error_code% EQU 0 (
call :dk_refresh
call :dk_color %Green% "[Successful]"
) else (
call :dk_color %Red% "[Unsuccessful] %error_code%"
if defined actidnotfound call :dk_color %Red% "Activation ID not found for this key."
echo Check this page for help https://massgrave.dev/troubleshoot
)
%line%

::========================================================================================================================================

:ins_done

echo:
if %_unattended%==1 timeout /t 2 & exit /b
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
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

:dk_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
%psc% write-host -back '%1' -fore '%2' '%3'
)
exit /b

::========================================================================================================================================

::  1st column = Activation ID
::  2nd column = Generic Retail/OEM/MAK Key
::  3rd column = SKU ID
::  4th column = Key Type
::  5th column = WMI Edition ID
::  6th column = Version name incase same Edition ID is used in different OS versions with different key
::  Separator  = _


:hwidkey

set f=
for %%# in (
8b351c9c-f398-4515-9900-09df49427262_XGV%f%PP-NM%f%H47-7TTH%f%J-W3F%f%W7-8HV%f%2C___4_OEM:NONSLP_Enterprise
c83cef07-6b72-4bbc-a28f-a00386872839_3V6%f%Q6-NQ%f%XCX-V8YX%f%R-9QC%f%YV-QPF%f%CT__27_Volume:MAK_EnterpriseN
4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7%f%JG-NP%f%HTM-C97J%f%M-9MP%f%GT-3V6%f%6T__48_____Retail_Professional
9fbaf5d6-4d83-4422-870d-fdda6e5858aa_2B8%f%7N-8K%f%FHP-DKV6%f%R-Y2C%f%8J-PKC%f%KT__49_____Retail_ProfessionalN
f742e4ff-909d-4fe9-aacb-3231d24a0c58_4CP%f%RK-NM%f%3K3-X6XX%f%Q-RXX%f%86-WXC%f%HW__98_____Retail_CoreN
1d1bac85-7365-4fea-949a-96978ec91ae0_N24%f%34-X9%f%D7W-8PF6%f%X-8DV%f%9T-8TY%f%MD__99_____Retail_CoreCountrySpecific
3ae2cc14-ab2d-41f4-972f-5e20142771dc_BT7%f%9Q-G7%f%N6G-PGBY%f%W-4YW%f%X6-6F4%f%BT_100_____Retail_CoreSingleLanguage
2b1f36bb-c1cd-4306-bf5c-a0367c2d97d8_YTM%f%G3-N6%f%DKC-DKB7%f%7-7M9%f%GH-8HV%f%X7_101_____Retail_Core
2a6137f3-75c0-4f26-8e3e-d83d802865a4_XKC%f%NC-J2%f%6Q9-KFHD%f%2-FKT%f%HY-KD7%f%2Y_119_OEM:NONSLP_PPIPro
e558417a-5123-4f6f-91e7-385c1c7ca9d4_YNM%f%GQ-8R%f%YV3-4PGQ%f%3-C8X%f%TP-7CF%f%BY_121_____Retail_Education
c5198a66-e435-4432-89cf-ec777c9d0352_84N%f%GF-MH%f%BT6-FXBX%f%8-QWJ%f%K7-DRR%f%8H_122_____Retail_EducationN
cce9d2de-98ee-4ce2-8113-222620c64a27_KCN%f%VH-YK%f%WX8-GJJB%f%9-H9F%f%DT-6F7%f%W2_125_Volume:MAK_EnterpriseS_VB
d06934ee-5448-4fd1-964a-cd077618aa06_43T%f%BQ-NH%f%92J-XKTM%f%7-KT3%f%KK-P39%f%PB_125_OEM:NONSLP_EnterpriseS_RS5
706e0cfd-23f4-43bb-a9af-1a492b9f1302_NK9%f%6Y-D9%f%CD8-W44C%f%Q-R8Y%f%TK-DYJ%f%WX_125_OEM:NONSLP_EnterpriseS_RS1
faa57748-75c8-40a2-b851-71ce92aa8b45_FWN%f%7H-PF%f%93Q-4GGP%f%8-M8R%f%F3-MDW%f%WW_125_OEM:NONSLP_EnterpriseS_TH
2c060131-0e43-4e01-adc1-cf5ad1100da8_RQF%f%NW-9T%f%PM3-JQ73%f%T-QV4%f%VQ-DV9%f%PT_126_Volume:MAK_EnterpriseSN_VB
e8f74caa-03fb-4839-8bcc-2e442b317e53_M33%f%WV-NH%f%Y3C-R7FP%f%M-BQG%f%PT-239%f%PG_126_Volume:MAK_EnterpriseSN_RS5
3d1022d8-969f-4222-b54b-327f5a5af4c9_2DB%f%W3-N2%f%PJG-MVHW%f%3-G7T%f%DK-9HK%f%R4_126_Volume:MAK_EnterpriseSN_RS1
60c243e1-f90b-4a1b-ba89-387294948fb6_NTX%f%6B-BR%f%YC2-K678%f%6-F6M%f%VQ-M7V%f%2X_126_Volume:MAK_EnterpriseSN_TH
eb6d346f-1c60-4643-b960-40ec31596c45_DXG%f%7C-N3%f%6C4-C4HT%f%G-X4T%f%3X-2YV%f%77_161_____Retail_ProfessionalWorkstation
89e87510-ba92-45f6-8329-3afa905e3e83_WYP%f%NQ-8C%f%467-V2W6%f%J-TX4%f%WX-WT2%f%RQ_162_____Retail_ProfessionalWorkstationN
62f0c100-9c53-4e02-b886-a3528ddfe7f6_8PT%f%T6-RN%f%W4C-6V7J%f%2-C2D%f%3X-MHB%f%PB_164_____Retail_ProfessionalEducation
13a38698-4a49-4b9e-8e83-98fe51110953_GJT%f%YN-HD%f%MQY-FRR7%f%6-HVG%f%C7-QPF%f%8P_165_____Retail_ProfessionalEducationN
df96023b-dcd9-4be2-afa0-c6c871159ebe_NJC%f%F7-PW%f%8QT-3324%f%D-688%f%JX-2YV%f%66_175_____Retail_ServerRdsh
d4ef7282-3d2c-4cf0-9976-8854e64a8d1e_V3W%f%VW-N2%f%PV2-CGWC%f%3-34Q%f%GF-VMJ%f%2C_178_____Retail_Cloud
af5c9381-9240-417d-8d35-eb40cd03e484_NH9%f%J3-68%f%WK7-6FB9%f%3-4K3%f%DF-DJ4%f%F6_179_____Retail_CloudN
8ab9bdd1-1f67-4997-82d9-8878520837d9_XQQ%f%YW-NF%f%FMW-XJPB%f%H-K87%f%32-CKF%f%FD_188_____OEM:DM_IoTEnterprise
ed655016-a9e8-4434-95d9-4345352c2552_QPM%f%6N-7J%f%2WJ-P88H%f%H-P3Y%f%RH-YY7%f%4H_191_OEM:NONSLP_IoTEnterpriseS_VB
d4bdc678-0a4b-4a32-a5b3-aaa24c3b0f24_K9V%f%KN-3B%f%GWV-Y624%f%W-MCR%f%MQ-BHD%f%CD_202_____Retail_CloudEditionN
92fb8726-92a8-4ffc-94ce-f82e07444653_KY7%f%PN-VR%f%6RX-83W6%f%Y-6DD%f%YQ-T6R%f%4W_203_____Retail_CloudEdition
d4f9b41f-205c-405e-8e08-3d16e88e02be_J7N%f%JW-V6%f%KBM-CC8R%f%W-Y29%f%Y4-HQ2%f%MJ_205_OEM:NONSLP_IoTEnterpriseSK
) do (
for /f "tokens=1-6 delims=_" %%A in ("%%#") do (

if %1==key if %osSKU%==%%C (

REM Detect key attempt 1

if "%2"=="attempt1" if not defined key (
echo "!applist!" | find /i "%%A" 1>nul && (
set key=%%B
set channel=%%D
)
)

REM Detect key attempt 2

if "%2"=="attempt2" if not defined key (
set actidnotfound=1
set 6th=%%F
if not defined 6th (
set key=%%B
set channel=%%D
) else (
echo "%branch%" | find /i "%%F" 1>nul && (
set key=%%B
set channel=%%D
)
)
)
)

)
)
exit /b

::========================================================================================================================================