@setlocal DisableDelayedExpansion
@echo off





:: =======================================================================================================
::
::   This script is a part of 'Microsoft Activation Scripts' project.
::
::   Homepages-
::   NsaneForums: (Login Required) https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/
::   GitHub: https://github.com/massgravel/Microsoft-Activation-Scripts
::   GitLab: https://gitlab.com/massgrave/microsoft-activation-scripts
::
::   Maintained by @WindowsAddict
::
:: =======================================================================================================













::========================================================================================================================================

cls
title Change Windows 10 Edition with Retail/OEM Key
set _elev=
if /i "%~1"=="-el" set _elev=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set "_psc=powershell"
set "nul=1>nul 2>nul"
set "ELine=echo: &echo ==== ERROR ==== &echo:"

::========================================================================================================================================

if %winbuild% LSS 17134 (
%ELine%
echo Unsupported OS version Detected.
echo OS Requirement - Windows 10 [17134] 1803 and later builds.
goto Ced_Done
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

(%nul% cscript //NoLogo "!_vbsf!" /File:"!_batf!" -el) && (
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
goto Ced_Done

:Passed

::========================================================================================================================================

::  Set buffer height independently of window height
::  https://stackoverflow.com/a/13351373
::  Written by @dbenham

mode con: cols=98 lines=31
%nul% %_psc% "&{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.height=36;$W.buffersize=$B;}"

::========================================================================================================================================

set slp=SoftwareLicensingProduct
set sls=SoftwareLicensingService
set wApp=55c92734-d682-4d71-983e-d6ec3f16059f
setlocal EnableDelayedExpansion

::  Check Installation type
set instype=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v InstallationType 2^>nul') do if not errorlevel 1 set "instype=%%b"

if not "%instype%"=="Client" (
%ELine%
echo Unsupported OS version [Server] Detected.
echo OS Requirement - Windows 10 [17134] 1803 and later builds.
goto Ced_Done
)

::  Check Windows Edition
set osedition=
for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%wApp%' and PartialProductKey is not NULL) get LicenseFamily /VALUE" 2^>nul') do if not errorlevel 1 set "osedition=%%a"
if not defined osedition for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition 2^>nul ^| find /i "Current Edition :"') do set "osedition=%%a"

cls
if "%osedition%"=="" (
%ELine%
echo OS Edition was not detected properly. Aborting...
goto Ced_Done
)

::========================================================================================================================================

echo _______________________________________________________________________________________________
echo:
echo  Note 1 - This script can not change 'Core'(Home) to 'Non-Core' (Pro) Editions.
echo           You'll have to do the above manually, Follow these steps.
echo         - Disable internet.
echo         - Go to Settings ^> Update ^& Security ^> Activation and 
echo           Insert 'Pro' Edition Product Key VK7JG-NPHTM-C97JM-9MPGT-3V66T
echo         - Follow on screen instructions, Done. [Incase of errors, restart the system.]
echo:
echo  Note 2 - Following option works only in W10 17134 (RS4) and later builds.
echo _______________________________________________________________________________________________
echo:
echo   You can change the Current Edition '%osedition%' to one of the following :
echo _______________________________________________________________________________________________

REM Thanks to @RPO for the help in codes

echo:
for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do echo %%a
echo:
choice /C:21 /N /M "[1] Continue [2] Exit : "
if %errorlevel%==1 exit /b
echo:

for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (

choice /C:NY /N /M "Do you want to change to the %%a edition? [Y,N] : "
if errorlevel 2 (

call :%%a %nul%
if "!key!"=="" cls &%ELine% &echo [%%a ^| %winbuild%] HWID Activation is Not Supported. &goto Ced_Done

cls
echo ____________________________________________________________________

REM  Thanks to @abbodi1406 for the WMI methods

echo:
echo Changing the Edition to Windows 10 %%a
wmic path %sls% where __CLASS='%sls%' call InstallProductKey ProductKey="!key!" %nul% && (
for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get ProductKeyChannel /value 2^>nul') do set "_channel=%%#"
wmic path %sls% where __CLASS='%sls%' call RefreshLicenseStatus %nul%
echo:
call echo Installing %%_channel%% Key [!key!] 
echo [Successful]
echo:
echo Reboot is required to properly change the Edition.
) || (
%ELine%
echo Installing Retail/OEM Key [!key!]
echo [Unsuccessful]
)
echo ____________________________________________________________________

goto Ced_Done
))

::========================================================================================================================================

:Ced_Done

echo:
echo Press any key to exit...
pause >nul
exit /b

::========================================================================================================================================

:: Retail_OEM Key List

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