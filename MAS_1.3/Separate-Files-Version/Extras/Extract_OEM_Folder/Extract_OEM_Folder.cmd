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
title Extract $OEM$ Folder
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
goto Done
)

::========================================================================================================================================

if %winbuild% LSS 7600 (
%ELine%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10 and their Server equivalent.
goto Done
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name
::  Written by @abbodi1406

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

mode con cols=98 lines=30

::  Get correct Desktop Location with powershell
::  Written by @dcshoecomp (superuser.com)
::  https://superuser.com/a/1413170

for /f "delims=" %%a in ('%_psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do Set "desktop=%%a"

if exist "%desktop%\$OEM$\" (
echo _____________________________________________________
%ELine%
echo $OEM$ folder already exists on the Desktop.
echo _____________________________________________________
goto Done2
)

set "_dir=%desktop%\$OEM$\$$\Setup\Scripts"
set _nofile=

set "_fdir1=Activators\HWID-KMS38_Activation"
set "HWID_Activation.cmd=%_fdir1%\HWID_Activation.cmd"
set "KMS38_Activation.cmd=%_fdir1%\KMS38_Activation.cmd"
set "ClipUp.exe=%_fdir1%\BIN\ClipUp.exe"
set "gatherosstate.exe=%_fdir1%\BIN\gatherosstate.exe"
set "slc.dll=%_fdir1%\BIN\slc.dll"
set "ARM64_gatherosstate.exe=%_fdir1%\BIN\ARM64_gatherosstate.exe"
set "ARM64_slc.dll=%_fdir1%\BIN\ARM64_slc.dll"

set "_fdir2=Activators\Online_KMS_Activation"
set "Activate.cmd=%_fdir2%\Activate.cmd"
set "Renewal_Setup.cmd=%_fdir2%\Renewal_Setup.cmd"
set "cleanosppx64.exe=%_fdir2%\BIN\cleanosppx64.exe"
set "cleanosppx86.exe=%_fdir2%\BIN\cleanosppx86.exe"

cd /d "!_work!"
pushd "!_work!"
cd ..
cd ..

if not exist "%HWID_Activation.cmd%" set _nofile=1
if not exist "%KMS38_Activation.cmd%" set _nofile=1
if not exist "%ClipUp.exe%" set _nofile=1
if not exist "%gatherosstate.exe%" set _nofile=1
if not exist "%slc.dll%" set _nofile=1
if not exist "%ARM64_gatherosstate.exe%" set _nofile=1
if not exist "%ARM64_slc.dll%" set _nofile=1

if not exist "%Activate.cmd%" set _nofile=1
if not exist "%Renewal_Setup.cmd%" set _nofile=1
if not exist "%cleanosppx64.exe%" set _nofile=1
if not exist "%cleanosppx86.exe%" set _nofile=1

if defined _nofile (
echo _____________________________________________________
%ELine%
echo Some files are missing in the 'Activators' folder.
echo _____________________________________________________
goto Done
)

::========================================================================================================================================

:Menu

cls
echo:
echo:
echo                             Extract the $OEM$ Folder on your desktop.
echo                                  For more details use Read me.
echo                       _______________________________________________________
echo                      ^|                                                       ^|
echo                      ^|                                                       ^|
echo                      ^|   [1] HWID                                            ^|
echo                      ^|                                                       ^|
echo                      ^|   [2] KMS38                                           ^|
echo                      ^|                                                       ^|
echo                      ^|   [3] HWID, Fallback to KMS38                         ^|
echo                      ^|                                                       ^|
echo                      ^|   [4] Online KMS                                      ^|
echo                      ^|                                                       ^|
echo                      ^|   [5] HWID ^+ Online KMS                               ^|
echo                      ^|                                                       ^|
echo                      ^|   [6] KMS38 ^+ Online KMS                              ^|
echo                      ^|                                                       ^|
echo                      ^|   [7] HWID, Fallback to KMS38 ^+ Online KMS            ^|
echo                      ^|                                                       ^|
echo                      ^|   [8] Exit                                            ^|
echo                      ^|                                                       ^|
echo                      ^|_______________________________________________________^|
echo:     
choice /C:12345678 /N /M ">                     Enter Your Choice [1,2,3,4,5,6,7,8] : "

if errorlevel 8 exit /b
if errorlevel 7 goto:$OEM$HWID_FB_KMS38-KMS
if errorlevel 6 goto:$OEM$KMS38KMS
if errorlevel 5 goto:$OEM$HWIDKMS
if errorlevel 4 goto:$OEM$KMS
if errorlevel 3 goto:$OEM$HWID_FB_KMS38
if errorlevel 2 goto:$OEM$KMS38
if errorlevel 1 goto:$OEM$HWID

::========================================================================================================================================

:$OEM$HWID

cls
call :Prep
call :HWIDPrep
call :export HWIDSetup "%_dir%\SetupComplete.cmd"
set error_=
call :HWIDPrep2

if defined error_ goto ErrorFound
set "_oem=HWID"
goto Done

:HWIDSetup:
@echo off

reg query HKU\S-1-5-19 1>nul 2>nul || (
echo ==== Error ====
echo Right click on this file and select 'Run as administrator'
echo Press any key to exit...
pause >nul
exit /b
)

call "%~dp0HWID_Activation.cmd" /u

cd /d "%SystemRoot%\Setup\"
if exist "%SystemRoot%\Setup\Scripts\" @RD /S /Q "%SystemRoot%\Setup\Scripts\"
exit /b
:HWIDSetup:

::========================================================================================================================================

:$OEM$KMS38

cls
call :Prep
call :KMS38Prep
call :export KMS38Setup "%_dir%\SetupComplete.cmd"
set error_=
call :KMS38Prep2

if defined error_ goto ErrorFound
set "_oem=KMS38"
goto Done

:KMS38Setup:
@echo off

reg query HKU\S-1-5-19 1>nul 2>nul || (
echo ==== Error ====
echo Right click on this file and select 'Run as administrator'
echo Press any key to exit...
pause >nul
exit /b
)

call "%~dp0KMS38_Activation.cmd" /u

cd /d "%SystemRoot%\Setup\"
if exist "%SystemRoot%\Setup\Scripts\" @RD /S /Q "%SystemRoot%\Setup\Scripts\"
exit /b
:KMS38Setup:

::========================================================================================================================================

:$OEM$HWID_FB_KMS38

cls
call :Prep

copy /y /b "%HWID_Activation.cmd%" "%_dir%\HWID_Activation.cmd" %nul%
call :KMS38Prep
call :export HWID_FB_KMS38 "%_dir%\SetupComplete.cmd"

set error_=
If not exist "%_dir%\HWID_Activation.cmd" (set error_=1)
call :KMS38Prep2

if defined error_ goto ErrorFound
set "_oem=HWID`, Fallback to KMS38"
goto Done

:HWID_FB_KMS38:
@echo off

setlocal EnableDelayedExpansion
reg query HKU\S-1-5-19 1>nul 2>nul || (
echo ==== Error ====
echo Right click on this file and select 'Run as administrator'
echo Press any key to exit...
pause >nul
exit /b
)

for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

::  Check Windows Edition
set osedition=
for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey is not NULL) get LicenseFamily /VALUE" 2^>nul') do if not errorlevel 1 set "osedition=%%a"
if not defined osedition for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition 2^>nul ^| find /i "Current Edition :"') do set "osedition=%%a"

::  Check Installation type
set instype=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v InstallationType 2^>nul') do if not errorlevel 1 set "instype=%%b"

set KMS38=
if "%winbuild%" GEQ "17763" if "%osedition%"=="EnterpriseS" set KMS38=1
if "%winbuild%" GEQ "17763" if "%osedition%"=="EnterpriseSN" set KMS38=1
if "%osedition%"=="EnterpriseG" set KMS38=1
if "%osedition%"=="EnterpriseGN" set KMS38=1
if not "%instype%"=="Client" echo %osedition%| findstr /I /B Server 1>nul && set KMS38=1

if defined KMS38 (
call "%~dp0KMS38_Activation.cmd" /u
) else (
call "%~dp0HWID_Activation.cmd" /u
)

cd /d "%SystemRoot%\Setup\"
if exist "%SystemRoot%\Setup\Scripts\" @RD /S /Q "%SystemRoot%\Setup\Scripts\"
exit /b
:HWID_FB_KMS38:

::========================================================================================================================================

:$OEM$KMS

cls
call :Prep
call :KMSPrep
call :export KMSSetup "%_dir%\SetupComplete.cmd"
set error_=
call :KMSPrep2

if defined error_ goto ErrorFound
set "_oem=Online KMS"
goto Done

:KMSSetup:
@echo off

============================================================================

:: Change value from 1 to 0 to disable KMS Renewal And Activation Task
set Renewal_And_Activation_Task=1

:: Change value from 1 to 0 to disable KMS activation desktop context menu 
set Desktop_context_menu=1

============================================================================

reg query HKU\S-1-5-19 1>nul 2>nul || (
echo ==== Error ====
echo Right click on this file and select 'Run as administrator'
echo Press any key to exit...
pause >nul
exit /b
)

if %Renewal_And_Activation_Task% EQU 1 call "%~dp0Renewal_Setup.cmd" /rat
if %Desktop_context_menu% EQU 1 call "%~dp0Renewal_Setup.cmd" /dcm

cd /d "%SystemRoot%\Setup\"
if exist "%SystemRoot%\Setup\Scripts\" @RD /S /Q "%SystemRoot%\Setup\Scripts\"
exit /b
:KMSSetup:

::========================================================================================================================================

:$OEM$HWIDKMS

cls
call :Prep
call :HWIDPrep
call :KMSPrep

call :export HWIDKMSSetup "%_dir%\SetupComplete.cmd"

set error_=
call :HWIDPrep2
call :KMSPrep2

if defined error_ goto ErrorFound
set "_oem=HWID `+ Online KMS"
goto Done

:HWIDKMSSetup:
@echo off

============================================================================

:: Change value from 1 to 0 to disable KMS Renewal And Activation Task
set Renewal_And_Activation_Task=1

:: Change value from 1 to 0 to disable KMS activation desktop context menu 
set Desktop_context_menu=1

============================================================================

reg query HKU\S-1-5-19 1>nul 2>nul || (
echo ==== Error ====
echo Right click on this file and select 'Run as administrator'
echo Press any key to exit...
pause >nul
exit /b
)

call "%~dp0HWID_Activation.cmd" /u
if defined HWIDAct set SkipWinAct=/swa

if %Renewal_And_Activation_Task% EQU 1 call "%~dp0Renewal_Setup.cmd" /rat %SkipWinAct%
if %Desktop_context_menu% EQU 1 call "%~dp0Renewal_Setup.cmd" /dcm %SkipWinAct%

cd /d "%SystemRoot%\Setup\"
if exist "%SystemRoot%\Setup\Scripts\" @RD /S /Q "%SystemRoot%\Setup\Scripts\"
exit /b
:HWIDKMSSetup:

::========================================================================================================================================

:$OEM$KMS38KMS

cls
call :Prep
call :KMS38Prep
call :KMSPrep
call :export KMS38KMSSetup "%_dir%\SetupComplete.cmd"
set error_=
call :KMS38Prep2
call :KMSPrep2

if defined error_ goto ErrorFound
set "_oem=KMS38 `+ Online KMS"
goto Done

:KMS38KMSSetup:
@echo off

============================================================================

:: Change value from 1 to 0 to disable KMS Renewal And Activation Task
set Renewal_And_Activation_Task=1

:: Change value from 1 to 0 to disable KMS activation desktop context menu 
set Desktop_context_menu=1

============================================================================

reg query HKU\S-1-5-19 1>nul 2>nul || (
echo ==== Error ====
echo Right click on this file and select 'Run as administrator'
echo Press any key to exit...
pause >nul
exit /b
)

call "%~dp0KMS38_Activation.cmd" /u

if %Renewal_And_Activation_Task% EQU 1 call "%~dp0Renewal_Setup.cmd" /rat
if %Desktop_context_menu% EQU 1 call "%~dp0Renewal_Setup.cmd" /dcm

cd /d "%SystemRoot%\Setup\"
if exist "%SystemRoot%\Setup\Scripts\" @RD /S /Q "%SystemRoot%\Setup\Scripts\"
exit /b
:KMS38KMSSetup:

::========================================================================================================================================

:$OEM$HWID_FB_KMS38-KMS

cls
call :Prep

copy /y /b "%HWID_Activation.cmd%" "%_dir%\HWID_Activation.cmd" %nul%
call :KMS38Prep
call :KMSPrep
call :export HWID_FB_KMS38-KMSSetup "%_dir%\SetupComplete.cmd"

set error_=
If not exist "%_dir%\HWID_Activation.cmd" (set error_=1)
call :KMS38Prep2
call :KMSPrep2

if defined error_ goto ErrorFound
set "_oem=HWID`, Fallback to KMS38 `+ Online KMS"
goto Done

:HWID_FB_KMS38-KMSSetup:
@echo off

============================================================================

:: Change value from 1 to 0 to disable KMS Renewal And Activation Task
set Renewal_And_Activation_Task=1

:: Change value from 1 to 0 to disable KMS activation desktop context menu 
set Desktop_context_menu=1

============================================================================

setlocal EnableDelayedExpansion
reg query HKU\S-1-5-19 1>nul 2>nul || (
echo ==== Error ====
echo Right click on this file and select 'Run as administrator'
echo Press any key to exit...
pause >nul
exit /b
)

for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

::  Check Windows Edition
set osedition=
for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey is not NULL) get LicenseFamily /VALUE" 2^>nul') do if not errorlevel 1 set "osedition=%%a"
if not defined osedition for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition 2^>nul ^| find /i "Current Edition :"') do set "osedition=%%a"

::  Check Installation type
set instype=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v InstallationType 2^>nul') do if not errorlevel 1 set "instype=%%b"

set KMS38=
if "%winbuild%" GEQ "17763" if "%osedition%"=="EnterpriseS" set KMS38=1
if "%winbuild%" GEQ "17763" if "%osedition%"=="EnterpriseSN" set KMS38=1
if "%osedition%"=="EnterpriseG" set KMS38=1
if "%osedition%"=="EnterpriseGN" set KMS38=1
if not "%instype%"=="Client" echo %osedition%| findstr /I /B Server 1>nul && set KMS38=1

if defined KMS38 (
call "%~dp0KMS38_Activation.cmd" /u
) else (
call "%~dp0HWID_Activation.cmd" /u
if defined HWIDAct set SkipWinAct=/swa
)

if %Renewal_And_Activation_Task% EQU 1 call "%~dp0Renewal_Setup.cmd" /rat %SkipWinAct%
if %Desktop_context_menu% EQU 1 call "%~dp0Renewal_Setup.cmd" /dcm %SkipWinAct%

cd /d "%SystemRoot%\Setup\"
if exist "%SystemRoot%\Setup\Scripts\" @RD /S /Q "%SystemRoot%\Setup\Scripts\"
exit /b
:HWID_FB_KMS38-KMSSetup:

::========================================================================================================================================

:ErrorFound

%ELine%
echo $OEM$ Folder was not created successfully...
goto :Done2
:Done
echo ________________________________________________________________________________________________
echo:
%EchoGreen% %_oem% `$OEM`$ folder is successfully created on the Desktop.
echo ________________________________________________________________________________________________
:Done2
echo:
echo Press any key to exit...
pause >nul
exit /b

::========================================================================================================================================

:Prep

cd /d "%desktop%"
md "%desktop%\$OEM$\$$\Setup\Scripts\"
md "%desktop%\$OEM$\$$\Setup\Scripts\BIN"

cd /d "!_work!"
pushd "!_work!"
cd ..
cd ..

exit /b

:HWIDPrep

copy /y /b "%HWID_Activation.cmd%" "%_dir%\HWID_Activation.cmd" %nul%
copy /y /b "%gatherosstate.exe%" "%_dir%\BIN\gatherosstate.exe" %nul%
copy /y /b "%slc.dll%" "%_dir%\BIN\slc.dll" %nul%
copy /y /b "%ARM64_gatherosstate.exe%" "%_dir%\BIN\ARM64_gatherosstate.exe" %nul%
copy /y /b "%ARM64_slc.dll%" "%_dir%\BIN\ARM64_slc.dll" %nul%
exit /b

:KMS38Prep

copy /y /b "%KMS38_Activation.cmd%" "%_dir%\KMS38_Activation.cmd" %nul%
copy /y /b "%ClipUp.exe%" "%_dir%\BIN\ClipUp.exe" %nul%
copy /y /b "%gatherosstate.exe%" "%_dir%\BIN\gatherosstate.exe" %nul%
copy /y /b "%slc.dll%" "%_dir%\BIN\slc.dll" %nul%
copy /y /b "%ARM64_gatherosstate.exe%" "%_dir%\BIN\ARM64_gatherosstate.exe" %nul%
copy /y /b "%ARM64_slc.dll%" "%_dir%\BIN\ARM64_slc.dll" %nul%
exit /b

:KMSPrep

copy /y /b "%Activate.cmd%" "%_dir%\Activate.cmd" %nul%
copy /y /b "%Renewal_Setup.cmd%" "%_dir%\Renewal_Setup.cmd" %nul%
copy /y /b "%cleanosppx64.exe%" "%_dir%\BIN\cleanosppx64.exe" %nul%
copy /y /b "%cleanosppx86.exe%" "%_dir%\BIN\cleanosppx86.exe" %nul%
exit /b

:HWIDPrep2

If not exist "%_dir%\HWID_Activation.cmd" (set error_=1)
If not exist "%_dir%\BIN\gatherosstate.exe" (set error_=1)
If not exist "%_dir%\BIN\slc.dll" (set error_=1)
If not exist "%_dir%\BIN\ARM64_gatherosstate.exe" (set error_=1)
If not exist "%_dir%\BIN\ARM64_slc.dll" (set error_=1)
If not exist "%_dir%\SetupComplete.cmd" (set error_=1)
exit /b

:KMS38Prep2

If not exist "%_dir%\KMS38_Activation.cmd" (set error_=1)
If not exist "%_dir%\BIN\ClipUp.exe" (set error_=1)
If not exist "%_dir%\BIN\gatherosstate.exe" (set error_=1)
If not exist "%_dir%\BIN\slc.dll" (set error_=1)
If not exist "%_dir%\BIN\ARM64_gatherosstate.exe" (set error_=1)
If not exist "%_dir%\BIN\ARM64_slc.dll" (set error_=1)
If not exist "%_dir%\SetupComplete.cmd" (set error_=1)
exit /b

:KMSPrep2

If not exist "%_dir%\Activate.cmd" (set error_=1)
If not exist "%_dir%\Renewal_Setup.cmd" (set error_=1)
If not exist "%_dir%\BIN\cleanosppx64.exe" (set error_=1)
If not exist "%_dir%\BIN\cleanosppx86.exe" (set error_=1)
If not exist "%_dir%\SetupComplete.cmd" (set error_=1)
exit /b

::========================================================================================================================================

::  Extract the text from batch script without character and file encoding issue
::  Thanks to @abbodi1406

:export
%nul% %_psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('%~2',$f[1].Trim(),[System.Text.Encoding]::ASCII);" &exit/b

::========================================================================================================================================