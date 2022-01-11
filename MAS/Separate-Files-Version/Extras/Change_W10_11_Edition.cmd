@setlocal DisableDelayedExpansion
@echo off



::============================================================================
::
::   This script is a part of 'Microsoft Activation Scripts' (MAS) project.
::
::   Homepage: windowsaddict.ml
::      Email: windowsaddict@protonmail.com
::
::============================================================================




::========================================================================================================================================

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

if exist %SystemRoot%\Sysnative\cmd.exe (
set "_cmdf=%~f0"
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %*"
exit /b
)

:: Re-launch the script with ARM32 process if it was initiated by x64 process on ARM64 Windows

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 (
set "_cmdf=%~f0"
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %*"
exit /b
)

::  Set Path variable, it helps if it is misconfigured in the system

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"

::========================================================================================================================================

cls
color 07
title  Change Windows 10-11 Edition

set _elev=
if /i "%~1"=="-el" set _elev=1

set winbuild=1
set "nul=>nul 2>&1"
set "_psc=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
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
set slp=SoftwareLicensingProduct
set sls=SoftwareLicensingService
set wApp=55c92734-d682-4d71-983e-d6ec3f16059f
set "line=echo ___________________________________________________________________________________________"

::========================================================================================================================================

if %winbuild% LSS 10240 (
%eline%
echo Unsupported OS version detected.
echo Project is supported for Windows 10/11.
goto ced_done
)

if %winbuild% GEQ 22483 if not exist %_psc% (
%nceline%
echo Powershell is not installed in the system.
goto ced_done
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set "_PSarg="""%~f0""" -el %_args%"

set "_ttemp=%temp%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" 1>nul && (
%eline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto ced_done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul% reg query HKU\S-1-5-19 || (
if not defined _elev %nul% %_psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto ced_done
)

::========================================================================================================================================

mode 98, 30
echo:
echo Initializing...

::  Check WMI and sppsvc Errors

set applist=
net start sppsvc /y %nul%
if %winbuild% LSS 22483 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%wApp%') get ID /VALUE" 2^>nul')"
if %winbuild% GEQ 22483 set "chkapp=for /f "tokens=2 delims==" %%a in ('%_psc% "(([WMISEARCHER]'SELECT ID FROM %slp% WHERE ApplicationID=''%wApp%''').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))

if not defined applist (
%eline%
echo Failed running WMI query check, verify that these services are working correctly
echo Windows Management Instrumentation [WinMgmt], Software Protection [sppsvc]
echo:
echo Script will try to enable these services.
echo:
call :dk_color %_Yellow% "Press any key to continue..."
pause >nul
for /f "skip=2 tokens=2*" %%a in ('reg query HKLM\SYSTEM\CurrentControlSet\Services\WinMgmt /v Start 2^>nul') do if /i %%b equ 0x4 (sc config WinMgmt start= auto %nul%)
net start WinMgmt /y %nul%
net stop sppsvc /y %nul%
net start sppsvc /y %nul%
cls
)

::========================================================================================================================================

::  Check Windows Server version

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo Windows Server version is not supported.
goto ced_done
)

::  Refresh license status, it helps to get correct product name in Windows 17134 and later builds

call :dk_refresh

::  Check Windows Edition

set osedition=
for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition 2^>nul ^| find /i "Current Edition :"') do set "osedition=%%a"

cls
if "%osedition%"=="" (
%eline%
echo OS Edition was not detected properly. Aborting...
goto ced_done
)

::  Check product name

set winos=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "winos=%%b"

::========================================================================================================================================

set _target=
for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (if defined _target (set "_target=!_target! %%a") else (set "_target=%%a"))

if not defined _target (
%line%
echo:
call :dk_color %Gray% "Target Edition not found."
echo Current Edition [%osedition% ^| %winbuild%] can not be changed to any other Edition.
%line%
goto ced_done
)

::========================================================================================================================================

%line%
echo:
call :dk_color %Gray% "You can change the Current Edition [%osedition%] to one of the following."
%line%
echo:
for %%# in (%_target%) do echo %%#
%line%
echo:
call :dk_color %_Green% "Press [1] to Continue or [2] to Exit"
choice /C:21 /N

if %errorlevel%==1 exit /b
cls

::========================================================================================================================================

%line%
echo:
call :dk_color %Gray% "Current Edition - [%osedition%]"
echo:
for %%# in (%_target%) do (
choice /C:NY /N /M "Do you want to change to the [%%#] edition? [Y,N] : "
if [!errorlevel!]==[2] (
set targetedition=%%#
goto ced_change
)
)

%line%
goto ced_done

::========================================================================================================================================

:ced_change

cls
set key=
set _changepk=1

call :changeeditiondata

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Unable to find this product key in the supported product list.
if %winbuild% GTR 19044 echo Make sure you are using updated version of the script
goto ced_done
)

::========================================================================================================================================

%line%

::  Changing from Core to Non-Core & Changing editions in Windows build older than 17134 requires "changepk /productkey" method and restart
::  In other cases, editions can be changed instantly with "slmgr /ipk"

if %_changepk%==1 (
echo %_chan% | find /i "OEM" >NUL && (
%eline%
echo [%osedition%] can not be changed to [%targetedition%] Edition due to lack of non OEM keys.
echo Non-OEM keys are required to change from Core to Non-Core Editions.
goto ced_done
)
for %%a in (dns.msftncsi.com,www.microsoft.com,one.one.one.one,resolver1.opendns.com) do (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (
if not [%%#]==[] (
%eline%
echo Disconnect the Internet and then try again.
goto ced_done
)
)
)

echo:
echo The system will automatically reboot to complete the process.
echo:
call :dk_color %Gray% "If the upgrade Window shows an error then you can safely ignore it."
call :dk_color %Gray% "However you will need to manually reboot the system in that case."
echo:
call :dk_color %Magenta% "Important - Save your work before continue."
echo:
choice /C:21 /N /M "[1] Continue [2] Exit : "
if !errorlevel!==1 exit /b
)

::========================================================================================================================================

echo:
echo Changing the Current Edition [%osedition%] to [%targetedition%]
echo:

if %_changepk%==0 (
echo Installing %_chan% Key [%key%]
echo:
if %winbuild% LSS 22483 wmic path %sls% where __CLASS='%sls%' call InstallProductKey ProductKey="%key%" %nul%
if %winbuild% GEQ 22483 %_psc% "(([WMISEARCHER]'SELECT Version FROM %sls%').Get()).InstallProductKey('%key%')" %nul%
if not !errorlevel!==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%

if !errorlevel!==0 (
call :dk_refresh
call :dk_color %Green% "[Successful]"
echo:
call :dk_color %Gray% "Reboot is required to properly change the Edition."
) else (
%eline%
echo [Unsuccessful]
)
)

if %_changepk%==1 (
echo Applying the command with %_chan% Key
echo start changepk /productkey %key%
start changepk /productkey %key%
)
%line%

::========================================================================================================================================

:ced_done

echo:
call :dk_color %_Yellow% "Press any key to exit..."
pause >nul
exit /b

::========================================================================================================================================

::  Refresh license status

:dk_refresh

if %winbuild% LSS 22483 wmic path %sls% where __CLASS='%sls%' call RefreshLicenseStatus %nul%
if %winbuild% GEQ 22483 %_psc% "$null=(([WMICLASS]'%sls%').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::========================================================================================================================================

:dk_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
if not exist %_psc% (echo %~3) else (%_psc% write-host -back '%1' -fore '%2' '%3')
)
exit /b

::========================================================================================================================================

::  1st column = Activation ID
::  2nd column = Generic Retail/OEM/MAK/GVLK Key
::  3rd column = SKU ID
::  4th column = 1 = activation is not working (at the time of writing this), 0 = activation is working
::  5th column = Key Type
::  6th column = WMI Edition ID
::  7th column = Version name incase same Edition ID is used in different OS versions with different key
::  Separator  = _

::  Key preference is in the following order. Retail > Volume:MAK > Volume:GVLK > OEM:NONSLP > OEM:DM
::  OEM keys are in last because they can't be used in edition change if "changepk /productkey" method is needed instead of "slmgr /ipk"
::  OEM keys are listed here because we don't have other keys for that edition

:changeeditiondata

for %%# in (
2ffd8952-423e-4903-b993-72a1aa44cf82_44NYX-TKR9D-CCM2D-V6B8F-HQWWR___4_0_Volume:MAK_Enterprise
23505d51-32d6-41f0-8ca7-e78ad0f16e71_D6RD9-D4N8T-RT9QX-YW6YT-FCWWJ__11_1_____Retail_Starter
c83cef07-6b72-4bbc-a28f-a00386872839_3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT__27_0_Volume:MAK_EnterpriseN
211b80cc-7f64-482c-89e9-4ba21ff827ad_3NFXW-2T27M-2BDW6-4GHRV-68XRX__47_1_____Retail_StarterN
4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7JG-NPHTM-C97JM-9MPGT-3V66T__48_0_____Retail_Professional
9fbaf5d6-4d83-4422-870d-fdda6e5858aa_2B87N-8KFHP-DKV6R-Y2C8J-PKCKT__49_0_____Retail_ProfessionalN
f742e4ff-909d-4fe9-aacb-3231d24a0c58_4CPRK-NM3K3-X6XXQ-RXX86-WXCHW__98_0_____Retail_CoreN
1d1bac85-7365-4fea-949a-96978ec91ae0_N2434-X9D7W-8PF6X-8DV9T-8TYMD__99_0_____Retail_CoreCountrySpecific
3ae2cc14-ab2d-41f4-972f-5e20142771dc_BT79Q-G7N6G-PGBYW-4YWX6-6F4BT_100_0_____Retail_CoreSingleLanguage
2b1f36bb-c1cd-4306-bf5c-a0367c2d97d8_YTMG3-N6DKC-DKB77-7M9GH-8HVX7_101_0_____Retail_Core
2a6137f3-75c0-4f26-8e3e-d83d802865a4_XKCNC-J26Q9-KFHD2-FKTHY-KD72Y_119_0_OEM:NONSLP_PPIPro
e558417a-5123-4f6f-91e7-385c1c7ca9d4_YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY_121_0_____Retail_Education
c5198a66-e435-4432-89cf-ec777c9d0352_84NGF-MHBT6-FXBX8-QWJK7-DRR8H_122_0_____Retail_EducationN
cce9d2de-98ee-4ce2-8113-222620c64a27_KCNVH-YKWX8-GJJB9-H9FDT-6F7W2_125_1_Volume:MAK_EnterpriseS_2021
97348f2f-bebc-4653-a4bd-18a895d316d9_VBX36-N7DDY-M9H62-83BMJ-CPR42_125_0_Volume:MAK_EnterpriseS_2019
2782d615-3249-495b-8260-15a4c2295448_PN3KR-JXM7T-46HM4-MCQGK-7XPJQ_125_0_Volume:MAK_EnterpriseS_2016
6366a32b-72e4-4212-bf11-c22b0e98a435_DVWKN-3GCMV-Q2XF4-DDPGM-VQWWY_125_0_Volume:MAK_EnterpriseS_2015
2c060131-0e43-4e01-adc1-cf5ad1100da8_RQFNW-9TPM3-JQ73T-QV4VQ-DV9PT_126_1_Volume:MAK_EnterpriseSN_2021
e8f74caa-03fb-4839-8bcc-2e442b317e53_M33WV-NHY3C-R7FPM-BQGPT-239PG_126_1_Volume:MAK_EnterpriseSN_2019
3d1022d8-969f-4222-b54b-327f5a5af4c9_2DBW3-N2PJG-MVHW3-G7TDK-9HKR4_126_0_Volume:MAK_EnterpriseSN_2016
60c243e1-f90b-4a1b-ba89-387294948fb6_NTX6B-BRYC2-K6786-F6MVQ-M7V2X_126_0_Volume:MAK_EnterpriseSN_2015
a48938aa-62fa-4966-9d44-9f04da3f72f2_G3KNM-CHG6T-R36X3-9QDG6-8M8K9_138_1_____Retail_ProfessionalSingleLanguage
f7af7d09-40e4-419c-a49b-eae366689ebd_HNGCC-Y38KG-QVK8D-WMWRK-X86VK_139_1_____Retail_ProfessionalCountrySpecific
eb6d346f-1c60-4643-b960-40ec31596c45_DXG7C-N36C4-C4HTG-X4T3X-2YV77_161_0_____Retail_ProfessionalWorkstation
89e87510-ba92-45f6-8329-3afa905e3e83_WYPNQ-8C467-V2W6J-TX4WX-WT2RQ_162_0_____Retail_ProfessionalWorkstationN
62f0c100-9c53-4e02-b886-a3528ddfe7f6_8PTT6-RNW4C-6V7J2-C2D3X-MHBPB_164_0_____Retail_ProfessionalEducation
13a38698-4a49-4b9e-8e83-98fe51110953_GJTYN-HDMQY-FRR76-HVGC7-QPF8P_165_0_____Retail_ProfessionalEducationN
5f87a508-7e1c-4fab-9d45-2356c6002081_C4NTJ-CX6Q2-VXDMR-XVKGM-F9DJC_171_1_Volume:MAK_EnterpriseG
1681ae34-3080-4bfa-a1b5-6d792342e692_46PN6-R9BK9-CVHKB-HWQ9V-MBJY8_172_1_Volume:MAK_EnterpriseGN
df96023b-dcd9-4be2-afa0-c6c871159ebe_NJCF7-PW8QT-3324D-688JX-2YV66_175_0_____Retail_ServerRdsh
d4ef7282-3d2c-4cf0-9976-8854e64a8d1e_V3WVW-N2PV2-CGWC3-34QGF-VMJ2C_178_0_____Retail_Cloud
af5c9381-9240-417d-8d35-eb40cd03e484_NH9J3-68WK7-6FB93-4K3DF-DJ4F6_179_0_____Retail_CloudN
c7051f63-3a76-4992-bce5-731ec0b1e825_2HN6V-HGTM8-6C97C-RK67V-JQPFD_183_1_____Retail_CloudE
8ab9bdd1-1f67-4997-82d9-8878520837d9_XQQYW-NFFMW-XJPBH-K8732-CKFFD_188_0_____OEM:DM_IoTEnterprise
ed655016-a9e8-4434-95d9-4345352c2552_QPM6N-7J2WJ-P88HH-P3YRH-YY74H_191_0_OEM:NONSLP_IoTEnterpriseS
d4bdc678-0a4b-4a32-a5b3-aaa24c3b0f24_K9VKN-3BGWV-Y624W-MCRMQ-BHDCD_202_0_____Retail_CloudEditionN
92fb8726-92a8-4ffc-94ce-f82e07444653_KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W_203_0_____Retail_CloudEdition
) do (
for /f "tokens=1-7 delims=_" %%A in ("%%#") do if /i %targetedition%==%%F (

if not defined key (
set 7th=%%G
if not defined 7th (
set "key=%%B" & set "_chan=%%E"
echo "!applist!" | find /i "%%A" 1>nul && set _changepk=0
) else (
echo "%winos%" | find "%%G" 1>nul && (set "key=%%B" & set "_chan=%%E" & echo "!applist!" | find /i "%%A" 1>nul && set _changepk=0)
)
)
)
)
exit /b

::========================================================================================================================================