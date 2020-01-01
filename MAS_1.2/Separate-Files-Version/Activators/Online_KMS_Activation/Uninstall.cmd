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
title Online KMS Complete Uninstall
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
set "EchoRed=%_psc% write-host -back Black -fore Red"
set "EchoGreen=%_psc% write-host -back Black -fore Green"
set "ELine=echo: & %EchoRed% ==== ERROR ==== &echo:"

::========================================================================================================================================

for %%i in (powershell.exe) do if "%%~$path:i"=="" (
echo: &echo ==== ERROR ==== &echo:
echo Powershell is not installed in the system.
echo Aborting...
goto UN_Done
)

::========================================================================================================================================

if %winbuild% LSS 7600 (
%ELine%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10 and their Server equivalent.
goto UN_Done
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
goto UN_Done

:Passed

::========================================================================================================================================

mode con: cols=98 lines=30
setlocal EnableDelayedExpansion
set "key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks"

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
@RD /s /q "%windir%\Online_KMS_Activation_Script\" %nul%
)

If exist "%ProgramData%\Online_KMS_Activation\" (
echo Deleting [Folder] %ProgramData%\Online_KMS_Activation\
@RD /s /q "%ProgramData%\Online_KMS_Activation\" %nul%
)

if exist "%ProgramData%\Online_KMS_Activation.cmd" (
echo Deleting [File] %ProgramData%\Online_KMS_Activation.cmd
del /f /q "%ProgramData%\Online_KMS_Activation.cmd" %nul%
)

reg query "HKCR\DesktopBackground\shell\Activate Windows - Office" %nul% && (
echo Deleting [Context Menu Registry] HKCR\DesktopBackground\shell\Activate Windows - Office
Reg delete "HKCR\DesktopBackground\shell\Activate Windows - Office" /f %nul%
)

echo:

::========================================================================================================================================

::  Clear-KMS-Cache.cmd  
::  https://forums.mydigitallife.net/posts/1511883
::  Written by @abbodi1406 (MDL)

set "SysPath=%Windir%\System32"
if exist "%Windir%\Sysnative\reg.exe" (set "SysPath=%Windir%\Sysnative")
set "Path=%SysPath%;%Windir%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"
set "OSPP=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
wmic path SoftwareLicensingProduct where (Description like '%%KMSCLIENT%%') get Name 2>nul | findstr /i Windows 1>nul && (set SppHook=1) || (set SppHook=0)
wmic path SoftwareLicensingProduct where (Description like '%%KMSCLIENT%%') get Name 2>nul | findstr /i Office 1>nul && (set SppHook=1)
wmic path OfficeSoftwareProtectionService get Version >nul 2>&1 && (set OsppHook=1) || (set OsppHook=0)
if %SppHook% NEQ 0 call :UN_cKMS SoftwareLicensingProduct SoftwareLicensingService SPP
if %OsppHook% NEQ 0 call :UN_cKMS OfficeSoftwareProtectionProduct OfficeSoftwareProtectionService OSPP
call :UN_cREG >nul 2>&1
echo:
goto UN_Verify

:UN_cKMS
echo Clearing %3 KMS Cache...
set spp=%1
set sps=%2
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%') get ID /VALUE" 2^>nul') do (set app=%%G&call :UN_cAPP)
for /f "tokens=2 delims==" %%A in ('"wmic path %sps% get Version /VALUE"') do set ver=%%A
wmic path %sps% where version='%ver%' call ClearKeyManagementServiceMachine >nul 2>&1
wmic path %sps% where version='%ver%' call ClearKeyManagementServicePort >nul 2>&1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 1 >nul 2>&1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 1 >nul 2>&1
goto :eof

:UN_cAPP
wmic path %spp% where ID='%app%' call ClearKeyManagementServiceMachine >nul 2>&1
wmic path %spp% where ID='%app%' call ClearKeyManagementServicePort >nul 2>&1
goto :eof

:UN_cREG
reg delete "HKLM\%SPPk%\55c92734-d682-4d71-983e-d6ec3f16059f" /f
reg delete "HKLM\%SPPk%\0ff1ce15-a989-479d-af46-f275c6370663" /f
reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort
reg delete "HKU\S-1-5-20\%SPPk%\55c92734-d682-4d71-983e-d6ec3f16059f" /f
reg delete "HKU\S-1-5-20\%SPPk%\0ff1ce15-a989-479d-af46-f275c6370663" /f
reg delete "HKLM\%OSPP%\59a52881-a989-479d-af46-f275c6370663" /f
reg delete "HKLM\%OSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f
reg delete "HKLM\%OSPP%" /f /v KeyManagementServiceName
reg delete "HKLM\%OSPP%" /f /v KeyManagementServicePort
if %OsppHook% NEQ 1 (
reg delete "HKLM\%OSPP%" /f
reg delete "HKU\S-1-5-20\%OSPP%" /f
)
goto :eof

::========================================================================================================================================

:UN_Verify

set error_=
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Renewal" >nul && (set error_=1)
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Run_Once" >nul && (set error_=1)
If exist "%windir%\Online_KMS_Activation_Script\" (set error_=1)
if exist "%ProgramData%\Online_KMS_Activation\" (set error_=1)

reg query "HKCR\DesktopBackground\shell\Activate Windows - Office" >nul 2>&1 && (set error_=1)
if exist "%ProgramData%\Online_KMS_Activation.cmd" (set error_=1)

if defined error_ (
echo ______________________________________
%ELine%
echo Try Again.
echo ______________________________________
) else (
echo _________________________________________________________________
echo:
%EchoGreen% Online KMS Complete Uninstall is done successfully.
echo _________________________________________________________________
)

::========================================================================================================================================

:UN_Done

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