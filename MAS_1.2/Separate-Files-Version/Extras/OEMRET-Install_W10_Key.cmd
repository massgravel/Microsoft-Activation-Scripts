@setlocal DisableDelayedExpansion
@echo off

:: For unattended mode, run the script with /u parameter.





:: =======================================================================================================
::
::   This script is a part of 'Microsoft Activation Scripts' project.
::
::   Homepages-
::   NsaneForums: (Login Required) https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/
::   GitLab: https://gitlab.com/massgrave/microsoft-activation-scripts
::
::   Maintained by @WindowsAddict
::
:: =======================================================================================================











::========================================================================================================================================

cls
title Install Windows 10 Retail/OEM Key
set Unattended=
set _args=
set _elev=
set "_arg1=%~1"
if not defined _arg1 goto :NoProgArgs
set "_args=%~1"
set "_arg2=%~2"
if defined _arg2 set "_args=%~1 %~2"
for %%A in (%_args%) do (
if /i "%%A"=="-el" set _elev=1
if /i "%%A"=="/u" set Unattended=1)
:NoProgArgs
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set "_psc=powershell -nop -ep bypass -c"
set "nul=1>nul 2>nul"
set "ELine=echo: &echo ==== ERROR ==== &echo:"

::========================================================================================================================================

if %winbuild% LSS 10240 (
%ELine%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 10.
goto Ins_Done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop
::  Thanks to @hearywarlot [ https://forums.mydigitallife.net/threads/.74332/ ] for the VBS method.
::  Thanks to @abbodi1406 for the powershell method and solving special characters issue in file path name.

%nul% reg query HKU\S-1-5-19 && (
  goto :Passed
  ) || (
  if defined _elev goto :E_Admin
)

set "_batf=%~f0"
set "_vbsf=%temp%\admin.vbs"
set _PSarg="""%~f0""" -el
if defined _args set _PSarg="""%~f0""" -el """%_args%"""

setlocal EnableDelayedExpansion

(
echo Set strArg=WScript.Arguments.Named
echo Set strRdlproc = CreateObject^("WScript.Shell"^).Exec^("rundll32 kernel32,Sleep"^)
echo With GetObject^("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" ^& strRdlproc.ProcessId ^& "'"^)
echo With GetObject^("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" ^& .ParentProcessId ^& "'"^)
echo If InStr ^(.CommandLine, WScript.ScriptName^) ^<^> 0 Then
echo strLine = Mid^(.CommandLine, InStr^(.CommandLine , "/File:"^) + Len^(strArg^("File"^)^) + 8^)
echo End If
echo End With
echo .Terminate
echo End With
echo CreateObject^("Shell.Application"^).ShellExecute "cmd.exe", "/c " ^& chr^(34^) ^& chr^(34^) ^& strArg^("File"^) ^& chr^(34^) ^& strLine ^& chr^(34^), "", "runas", 1
)>"!_vbsf!"

(%nul% cscript //NoLogo "!_vbsf!" /File:"!_batf!" -el "!_args!") && (
del /f /q "!_vbsf!"
exit /b
) || (
del /f /q "!_vbsf!"
%nul% %_psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && (
exit /b
) || (
goto :E_Admin
)
)
exit /b

:E_Admin
%ELine%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto Ins_Done

:Passed

mode con: cols=98 lines=30
setlocal EnableDelayedExpansion

::========================================================================================================================================

::  Check Windows OS name

set winos=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do if not errorlevel 1 set "winos=%%b"
if not defined winos for /f "tokens=2* delims== " %%a in ('"wmic os get caption /value" 2^>nul') do if not errorlevel 1 set "winos=%%b"

::  Check SKU value

set SKU=
for /f "tokens=2 delims==" %%a IN ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" 2^>nul') do if not errorlevel 1 (set osSKU=%%a)
if not defined SKU for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do if not errorlevel 1 (set osSKU=%%a)

if "%osSKU%"=="" (
%ELine%
echo SKU value was not detected properly. Aborting...
goto Ins_Done
)

::  Check Windows Edition with SKU value for better accuracy

set osedition=
call :CheckEdition %nul%

if "%osedition%"=="" (
%ELine%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%] HWID Activation is Not Supported.
goto Ins_Done
)

set key=
call :%osedition% %nul%

if "%key%"=="" (
%ELine%
echo [%winos% ^| %winbuild%] HWID Activation is Not Supported.
goto Ins_Done
)

::========================================================================================================================================

if defined Unattended goto ContinueKeyInsert

cls
echo ___________________________________________________________________________________________
echo:
echo  Install [%winos% ^| %winbuild%] Retail/OEM Key 
echo  [%key%]
echo ___________________________________________________________________________________________
echo:
choice /C:12 /N /M "[1] Continue [2] Exit : "

if errorlevel 2 exit /b
if errorlevel 1 goto ContinueKeyInsert

:ContinueKeyInsert

cls
echo ___________________________________________________________________________________________

::  Thanks to @abbodi1406 for the WMI methods

set slp=SoftwareLicensingProduct
set sls=SoftwareLicensingService
set wApp=55c92734-d682-4d71-983e-d6ec3f16059f

wmic path %sls% where __CLASS='%sls%' call InstallProductKey ProductKey="%key%" %nul% && (
for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get ProductKeyChannel /value 2^>nul') do set "_channel=%%#"
wmic path %sls% where __CLASS='%sls%' call RefreshLicenseStatus %nul%
echo:
echo [%winos% ^| %winbuild%]
call echo Installing %%_channel%% Key [%key%] 
echo [Successful]
) || (
%ELine%
echo Installing Retail/OEM Key [%key%]
echo [Unsuccessful]
)
echo ___________________________________________________________________________________________

::========================================================================================================================================

:Ins_Done
echo:
if defined Unattended (
echo Exiting in 3 seconds...
if %winbuild% LSS 7600 (ping -n 3 127.0.0.1 > nul) else (timeout /t 3)
exit /b
)
echo Press any key to exit...
pause >nul
exit /b

::========================================================================================================================================

::  Check Windows Edition with SKU value for better accuracy

:CheckEdition

for %%# in (
4:Enterprise
27:EnterpriseN
48:Professional
49:ProfessionalN
98:CoreN
99:CoreCountrySpecific
100:CoreSingleLanguage
101:Core
121:Education
122:EducationN
125:EnterpriseS
126:EnterpriseSN
161:ProfessionalWorkstation
162:ProfessionalWorkstationN
164:ProfessionalEducation
165:ProfessionalEducationN
175:ServerRdsh
188:IoTEnterprise
) do for /f "tokens=1,2 delims=:" %%A in ("%%#") do (
if %osSKU%==%%A set "osedition=%%B"
)
exit /b

::========================================================================================================================================

:: Retail/OEM Key List

:Core
set "key=YTMG3-N6DKC-DKB77-7M9GH-8HVX7"
exit /b

:CoreCountrySpecific
set "key=N2434-X9D7W-8PF6X-8DV9T-8TYMD"
exit /b

:CoreN
set "key=4CPRK-NM3K3-X6XXQ-RXX86-WXCHW"
exit /b

:CoreSingleLanguage
set "key=BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"
exit /b

:Education
set "key=YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"
exit /b

:EducationN
set "key=84NGF-MHBT6-FXBX8-QWJK7-DRR8H"
exit /b

:Enterprise
set "key=XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
exit /b

:EnterpriseN
set "key=3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT"
exit /b

:EnterpriseS
if "%winbuild%" EQU "10240" set "key=FWN7H-PF93Q-4GGP8-M8RF3-MDWWW"
if "%winbuild%" EQU "14393" set "key=NK96Y-D9CD8-W44CQ-R8YTK-DYJWX"
exit /b

:EnterpriseSN
if "%winbuild%" EQU "10240" set "key=8V8WN-3GXBH-2TCMG-XHRX3-9766K"
if "%winbuild%" EQU "14393" set "key=2DBW3-N2PJG-MVHW3-G7TDK-9HKR4"
exit /b

:Professional
set "key=VK7JG-NPHTM-C97JM-9MPGT-3V66T"
exit /b

:ProfessionalEducation
set "key=8PTT6-RNW4C-6V7J2-C2D3X-MHBPB"
exit /b

:ProfessionalEducationN
set "key=GJTYN-HDMQY-FRR76-HVGC7-QPF8P"
exit /b

:ProfessionalN
set "key=2B87N-8KFHP-DKV6R-Y2C8J-PKCKT"
exit /b

:ProfessionalWorkstation
set "key=DXG7C-N36C4-C4HTG-X4T3X-2YV77"
exit /b

:ProfessionalWorkstationN
set "key=WYPNQ-8C467-V2W6J-TX4WX-WT2RQ"
exit /b

:ServerRdsh
set "key=NJCF7-PW8QT-3324D-688JX-2YV66"
exit /b

:IoTEnterprise
set "key=XQQYW-NFFMW-XJPBH-K8732-CKFFD"
exit /b

::========================================================================================================================================