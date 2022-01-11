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



::  To activate with Downlevel method (default), run the script with /a parameter or change 0 to 1 in below line
set _acti=0

::  To only generate GenuineTicket.xml with Downlevel method (default), run the script with /g parameter or change 0 to 1 in below line
set _gent=0

::  To enable LockBox method, run the script with /k parameter or change 0 to 1 in below line
::  You need to use this option with either activation or ticket generation. 
::  Example,
::  HWID_Activation.cmd /a /k
::  HWID_Activation.cmd /g /k
set _lock=0



::  If value is changed in above lines or any parameter is used then script will run in unattended mode
::  Incase if more than one options are used then only one option will be applied



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
title  HWID Activation

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/a"  set _acti=1
if /i "%%A"=="/g"  set _gent=1
if /i "%%A"=="/k"  set _lock=1
if /i "%%A"=="-el" set _elev=1
)
)

for %%A in (%_acti% %_gent% %_lock%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

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
if %~z0 GEQ 1500000 (set "_exitmsg=Go back") else (set "_exitmsg=Exit")
set "notifytocheckupdate=if %winbuild% GTR 19044 echo Make sure you are using updated version of the script."

::========================================================================================================================================

if %winbuild% LSS 10240 (
%eline%
echo Unsupported OS version detected.
echo Project is supported for Windows 10/11.
goto dk_done
)

if %winbuild% GEQ 22483 if not exist %_psc% (
%nceline%
echo Powershell is not installed in the system.
goto dk_done
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

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
goto dk_done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul% reg query HKU\S-1-5-19 || (
if not defined _elev %nul% %_psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto dk_done
)

::========================================================================================================================================

:dl_menu

if %_unattended%==0 (
cls
mode 76, 25
title  HWID Activation

if !_lock!==0 (set "_method=%_Green% "[Downlevel Method]"") else (set "_method=%_Yellow% "  [LockBox Method]"")
echo:
echo:
echo:
echo         ____________________________________________________________
echo:
call :dk_color2 %_White% "                [1] HWID Activation       " !_method!
echo                 ____________________________________________
echo:
call :dk_color2 %_White% "                [2] Generate Ticket       " !_method!
echo                 ____________________________________________
echo:      
echo                 [3] Change Method
echo:
echo                 [4] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "              " %_Green% "Enter a menu option in the Keyboard [1,2,3,4]"
choice /C:1234 /N
set _el=!errorlevel!
if !_el!==4  exit /b
if !_el!==3  (
if !_lock!==0 (
set _lock=1
) else (
set _lock=0
)
cls
echo:
call :dk_color %_Green% " Downlevel Method:"
echo  It creates downlevelGTkey ticket for activation with simplest process.
echo:
call :dk_color %_Yellow% " LockBox Method:"
echo  It creates clientLockboxKey ticket which better mimics genuine activation,
echo  But requires more steps such as,
echo  - Cleaning ClipSVC licences
echo  - Deleting a volatile and protected registry key by taking ownership
echo  - System may need a restart for succesfull activation
echo  - Microsoft Account and Store Apps may need relogin-restart in the system
echo:
call :dk_color2 %_White% " " %Green% "Note:"
echo  Microsoft accepts both types of tickets and that's unlikely to change.
echo  If you are not sure what to choose then select default Downlevel Method.
echo:
call :dk_color %_Yellow% " Press any key to go back..."
pause >nul
goto :dl_menu
)
if !_el!==2  set _gent=1&goto :dl_menu2
if !_el!==1  goto :dl_menu2
goto :dl_menu
)

:dl_menu2

cls
if %_gent%==1 (set _title=title  Generate HWID GenuineTicket.xml) else (set _title=title  HWID Activation)
if %_lock%==0 (%_title% [Downlevel Method] & mode 102, 30) else (%_title% [Lockbox Method] & mode 102, 32)

::========================================================================================================================================

if not exist %_psc% if %_lock%==1 (
set _lock=0
set _gent=0
%nceline%
echo Powershell is not installed in the system.
echo It is required for Lockbox Method of HWID.
echo You need to set the script to the default.
if %_unattended%==0 (
echo:
call :dk_color %_Yellow% "Press any key to go back..."
pause >nul
goto dl_menu
) else (
goto dk_done
)
)

if %_gent%==1 if exist %Systemdrive%\GenuineTicket.xml (
set _gent=0
%eline%
echo File '%Systemdrive%\GenuineTicket.xml' already exist.
if %_unattended%==0 (
echo:
call :dk_color %_Yellow% "Press any key to go back..."
pause >nul
goto dl_menu
) else (
goto dk_done
)
)

::========================================================================================================================================

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
if %_unattended%==0 (
call :dk_color %_Yellow% "Press any key to continue..."
pause >nul
)
for /f "skip=2 tokens=2*" %%a in ('reg query HKLM\SYSTEM\CurrentControlSet\Services\WinMgmt /v Start 2^>nul') do if /i %%b equ 0x4 (sc config WinMgmt start= auto %nul%)
net start WinMgmt /y %nul%
net stop sppsvc /y %nul%
net start sppsvc /y %nul%
cls
)

::========================================================================================================================================

::  Refresh license status, it helps to get correct product name in Windows 17134 and later builds

call :dk_refresh

::  Check product name

set winos=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "winos=%%b"

::========================================================================================================================================

::  Check if system is permanently activated or not

cls
call :dk_checkperm
if defined _perm if not %_gent%==1 (
echo ___________________________________________________________________________________________
echo:
call :dk_color2 %_White% "     " %Green% "Checking: %winos% is Permanently Activated."
call :dk_color2 %_White% "     " %Gray% "Activation is not required."
echo ___________________________________________________________________________________________
if %_unattended%==1 goto dk_done
echo:
choice /C:12 /N /M ">    [1] Activate [2] %_exitmsg% : "
if errorlevel 2 exit /b
)
cls

::========================================================================================================================================

::  Check Evaluation version

set _eval=
set _evalserv=
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _eval=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" set _evalserv=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" set _eval=1 & set _evalserv=1

if defined _eval (
%eline%
echo [%winos% ^| %winbuild%]
if defined _evalserv (
echo Server Evaluation cannot be activated. Convert it to full Server OS.
) else (
echo Evaluation Editions cannot be activated. Install full Windows OS.
echo Check the ReadMe for how to get genuine installation media for full version.
)
goto dk_done
)

::========================================================================================================================================

::  Check SKU value

set osSKU=
for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do set "osSKU=%%a"

if not defined osSKU (
%eline%
echo SKU value was not detected properly. Aborting...
goto dk_done
)

::========================================================================================================================================

::  Check if HWID key (Retail,OEM,MAK) is already installed or not

set _hwidk=
call :dk_channel
for %%A in (Retail,OEM,MAK) do echo: %_channel%| findstr /i "%%A" >nul && set _hwidk=1

::========================================================================================================================================

::  Detect Key

set key=
set notworking=
set actidnotfound=

if defined applist call :hwiddata attempt1
if not defined key call :hwiddata attempt2

::========================================================================================================================================

if not defined key if not defined _hwidk (
%eline%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
echo Unable to find this product in the supported product list.
%notifytocheckupdate%
echo:
echo However, if you would like to try HWID activation on this product then,
echo install any generic Retail, OEM, MAK key for this product and run the script.
goto dk_done
)

if not defined key (
echo:
call :dk_color %Magenta% "====== Info ======"
echo:
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
echo Unable to find this product in the supported product list.
%notifytocheckupdate%
echo:
echo Since %_channel% key is already installed, script will try to activate with HWID.
echo:
echo It may or may not activate it.
if %_unattended%==0 (
echo:
call :dk_color %_Yellow% "Press any key to continue..."
pause >nul
)
cls
)

::========================================================================================================================================

::  Enterprise LTSC 2021 doesn't support HWID (At the time of writing this).
::  To activate it with HWID, script insert the product key of Iot Enterprise LTSC 2021. Restart is required for full effect.

::  If you don't want to change it then comment/delete the below lines.

set changekey=
if /i %key%==KCNVH-YKWX8-GJJB9-H9FDT-6F7W2 (
set _chan=OEM:NONSLP
set changekey=1
set notworking=
set key=QPM6N-7J2WJ-P88HH-P3YRH-YY74H
)

::========================================================================================================================================

::  Check and show info for editions which doesn't support HWID now but may support it in future

if defined notworking (
echo:
call :dk_color %Magenta% "====== Info ======"
echo:
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
echo At the time of writing this, HWID Activation was not supported for this product.
echo:
echo Now it may or may not activate it.
if %_unattended%==0 (
echo:
call :dk_color %_Yellow% "Press any key to continue..."
pause >nul
)
cls
)

::========================================================================================================================================

::  Check Windows Architecture 

set arch=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b

if not defined arch (
%eline%
echo Unable to detect Windows Architecture. Aborting...
goto dk_done
)

::========================================================================================================================================

::  Check files

set ARM64_file=
if /i "%arch%"=="ARM64" set ARM64_file=arm64_

for %%# in (%ARM64_file%gatherosstate.exe %ARM64_file%slc.dll) do (
if not exist "!_work!\BIN\%%#" (
%eline%
echo '%%#' file is missing in 'BIN' folder. Aborting...
goto dk_done
)
)

::  Verify gatherosstate.exe file

set _hash=
if /i "%arch%"=="ARM64" (set _orig=7E449AE5549A0D93CF65F4A1BB2AA7D1DC090D2D) else (set _orig=FABB5A0FC1E6A372219711152291339AF36ED0B5)
for /f "skip=1 tokens=* delims=" %%# in ('certutil -hashfile "!_work!\BIN\%ARM64_file%gatherosstate.exe" SHA1^|findstr /i /v CertUtil') do set "_hash=%%#"
set "_hash=%_hash: =%"

if /i not "%_hash%"=="%_orig%" (
%eline%
echo %ARM64_file%gatherosstate.exe SHA1 hash mismatch found.
echo:
echo Expected: %_orig%
echo Detected: %_hash%
goto dk_done
)

::========================================================================================================================================

::  Check Internet connection

cls
echo:
echo Checking OS Info                        [%winos% ^| %winbuild% ^| %arch%]

if not %_gent%==1 (
set _intcon=
ping -n 1 dns.msftncsi.com 2>nul | find "131.107.255.255" 1>nul || ping -n 1 www.microsoft.com 1>nul
if !errorlevel!==0 ( 
set _intcon=1
echo Checking Internet Connection            [Connected]
) else (
call :dk_color %Red% "Checking Internet Connection            [Not connected]"
)
)

::========================================================================================================================================

echo:
set "_serv=ClipSVC wlidsvc sppsvc LicenseManager Winmgmt wuauserv"

::  Client License Service (ClipSVC)
::  Microsoft Account Sign-in Assistant
::  Software Protection
::  Windows License Manager Service
::  Windows Management Instrumentation
::  Windows Update

echo Checking Services                       [%_serv%]

::  Check disabled services

set serv_ste=
for %%# in (%_serv%) do (
set serv_dis=
reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v Start %nul% || set serv_dis=1
for /f "skip=2 tokens=2*" %%a in ('reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v Start 2^>nul') do if /i %%b equ 0x4 set serv_dis=1
if defined serv_dis (if defined serv_ste (set "serv_ste=!serv_ste! %%#") else (set "serv_ste=%%#"))
)

::  Change disabled services startup type to auto

set serv_csts=
set serv_cste=

if defined serv_ste (
for %%# in (%serv_ste%) do (
sc config %%# start= auto %nul% && (
if defined serv_csts (set "serv_csts=!serv_csts! %%#") else (set "serv_csts=%%#")
) || (
if defined serv_cste (set "serv_cste=!serv_cste! %%#") else (set "serv_cste=%%#")
)
)
)

if defined serv_csts echo Enabling Disabled Services              [Successful] [%serv_csts%]
if defined serv_cste call :dk_color %Red% "Enabling Disabled Services              [Failed] [%serv_cste%]"

::========================================================================================================================================

::  Check if the services are able to run or not

set serv_e=
for %%# in (%_serv%) do (
sc query %%# | find /i "RUNNING" %nul% || net start %%# /y %nul%
sc query %%# | find /i "RUNNING" %nul% || sc start %%# %nul%
sc query %%# | find /i "RUNNING" %nul% || if defined serv_e (set "serv_e=!serv_e! %%#") else (set "serv_e=%%#")
)

if not defined serv_e (
echo Starting Services                       [Successful]
) else (
call :dk_color %Red% "Starting Services                       [Failed] [%serv_e%]"
echo %serv_e% | find /i "wuauserv" %nul% && (
call :dk_color %Magenta% "Windows Update Service [wuauserv] is not working, check if you have blocked it"
)
)

if not defined applist (
call :dk_color %Red% "Checking WMI Query                      [Failed]"
) else (
echo Checking WMI Query                      [Successful]
)

::========================================================================================================================================

::  Install key

echo:
if defined changekey call :dk_color %Magenta% "Windows 10 Iot Enterprise LTSC 2021 Product Key Is Selected For HWID Activation"&echo:

set _partial=
if defined key set _ipartial=%key:~-5%

if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get PartialProductKey /value 2^>nul') do set "_partial=%%#"
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%# in ('%_psc% "(([WMISEARCHER]'SELECT PartialProductKey FROM %slp% WHERE ApplicationID=''%wApp%'' AND PartialProductKey IS NOT NULL').Get()).PartialProductKey | %% {echo ('PartialProductKey='+$_)}" 2^>nul') do set "_partial=%%#"

if defined key if /i "%_partial%"=="%_ipartial%" (
echo Checking Installed Product Key          [%key%] [%_channel%]
)

if not defined key (
echo Checking Installed Product Key          [Partial Key - %_partial%] [%_channel%]
)

set _channel=
if defined key if /i not "%_partial%"=="%_ipartial%" (
if %winbuild% LSS 22483 wmic path %sls% where __CLASS='%sls%' call InstallProductKey ProductKey="%key%" %nul%
if %winbuild% GEQ 22483 %_psc% "(([WMISEARCHER]'SELECT Version FROM %sls%').Get()).InstallProductKey('%key%')" %nul%
if not !errorlevel!==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%

if !errorlevel!==0 (
call :dk_refresh
echo Installing Generic Product Key          [%key%] [%_chan%] [Successful]
) else (
call :dk_color %Red% "Installing Generic Product Key          [%key%] [%_chan%] [Failed]%actidnotfound%"
)
)

::========================================================================================================================================

::  Files are copied to temp to generate ticket to avoid possible issues in case the path contains special character or non English names

echo:
set "temp_=%SystemRoot%\Temp\_Temp"
if exist "%temp_%\.*" rmdir /s /q "%temp_%\" %nul%
md "%temp_%\" %nul%

pushd "!_work!\BIN\"
copy /y /b "%ARM64_file%gatherosstate.exe" "%temp_%\gatherosstate.exe" %nul%
copy /y /b "%ARM64_file%slc.dll" "%temp_%\slc.dll" %nul%
popd

set copyf=
if not exist "%temp_%\gatherosstate.exe" set copyf=1
if not exist "%temp_%\slc.dll" set copyf=1

if defined copyf (
call :dk_color %Red% "Copying Required Files to Temp          [%temp_%] [Failed]"
goto :dl_final
) else (
echo Copying Required Files to Temp          [%temp_%] [Successful]
)

::========================================================================================================================================

::  Modify the Pfn value in gatherosstate with slc.dll as per the system, that way one gatherosstate can be used in all the editions

pushd "%temp_%\"
rundll32 "%temp_%\slc.dll",PatchGatherosstate %nul%
popd
if not exist "%temp_%\gatherosstatemodified.exe" (
call :dk_color %Red% "Creating Modified Gatherosstate         [Failed] Aborting..."
call :dk_color %Magenta% "Most likely Antivirus blocked the process, disable it and/or create proper exclsuions"
goto :dl_final
) else (
echo Creating Modified Gatherosstate         [Successful]
)

::========================================================================================================================================

::  Clean ClipSVC Licences
::  This code runs only if Lockbox method to generate ticket is manually set by the user in this script.

if %_lock%==1 (
for %%# in (ClipSVC) do (
sc query %%# | find /i "STOPPED" %nul% || net stop %%# /y %nul%
sc query %%# | find /i "STOPPED" %nul% || sc stop %%# %nul%
)

rundll32 clipc.dll,ClipCleanUpState

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\*.dat" del /f /q "%ProgramData%\Microsoft\Windows\ClipSVC\*.dat" %nul%

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat" (
call :dk_color %Red% "Cleaning ClipSVC Licences               [Failed]"
) else (
echo Cleaning ClipSVC Licences               [Successful]
)
)

::========================================================================================================================================

::  Below registry key (Volatile & Protected) gets created after the ClipSVC License cleanup command, and gets automatically deleted after 
::  system restart. It needs to be deleted to activate the system without restart.

::  This code runs only if Lockbox method to generate ticket is manually set by the user in this script.

set "RegKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState"

if %_lock%==1 (
call :regown "%RegKey%" %nul% 
reg delete "%RegKey%" /f %nul% 

reg query "%RegKey%" %nul% && (
call :dk_color %Red% "Deleting a Volatile Registry            [Failed]"
call :dk_color %Magenta% "Restart the system, that will delete this registry key automatically"
) || (
echo Deleting a Volatile Registry            [Successful]
)
)

::========================================================================================================================================

::  Multiple attempts to generate the ticket because in some cases, one attempt is not enough.

echo:
set "_noxml=if not exist "%temp_%\GenuineTicket.xml""

start /wait "" "%temp_%/gatherosstatemodified.exe" %nul%
%_noxml% timeout /t 3 %nul%
%_noxml% net stop sppsvc /y %nul%
%_noxml% call "%temp_%/gatherosstatemodified.exe" %nul%
%_noxml% timeout /t 3 %nul%
%_noxml% "%temp_%/gatherosstatemodified.exe" %nul%
%_noxml% timeout /t 3 %nul%

::  Refresh ClipSVC (required after cleanup) with below command, not related to generating tickets

if %_lock%==1 (
for %%# in (wlidsvc LicenseManager sppsvc) do (net stop %%# /y %nul% & net start %%# /y %nul%)
call :dk_refresh
)

%_noxml% (
call :dk_color %Red% "Generating GenuineTicket.xml            [Failed] Aborting..."
goto :dl_final
)

if %_lock%==1 (
find /i "clientLockboxKey" "%temp_%\GenuineTicket.xml" >nul && (
echo Generating GenuineTicket.xml            [Successful] [clientLockboxKey Ticket]
) || (
call :dk_color %Red% "Generating GenuineTicket.xml            [Failed] [downlevelGTkey Ticket created] Aborting..."
call :dk_color %Magenta% "Try again / Restart system"
goto :dl_final
)
) else (
echo Generating GenuineTicket.xml            [Successful]
)

::========================================================================================================================================

::  Copy GenuineTicket.xml to the root of C drive and exit if ticket generation option was used in script

if %_gent%==1 (
echo:
copy /y /b "%temp_%\GenuineTicket.xml" "%Systemdrive%\GenuineTicket.xml" %nul%
if not exist "%Systemdrive%\GenuineTicket.xml" (
call :dk_color %Red% "Copying GenuineTicket.xml to %Systemdrive%\        [Failed]"
) else (
call :dk_color %Green% "Copying GenuineTicket.xml to %Systemdrive%\        [Successful]"
)
goto :dl_final
)

::========================================================================================================================================

::  clipup -v -o -altto <Ticket path> method to apply ticket is not used to avoid the certain issues in case if the username have 
::  spaces / special characters / non English names

set "tdir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
copy /y /b "%temp_%\GenuineTicket.xml" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "Failed to copy Ticket to [%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket\]"
goto :dl_final
)

set "_xmlexist=if exist "%tdir%\GenuineTicket.xml""

net stop ClipSVC /y %nul%
net start ClipSVC /y %nul%
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

%_xmlexist% %_psc% Restart-Service ClipSVC %nul%
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

set fallback_=
%_xmlexist% (
set fallback_=1
%nul% clipup -v -o
%_xmlexist% timeout /t 2 %nul%
)

%_xmlexist% (
call :dk_color %Red% "Installing GenuineTicket.xml            [Failed] Aborting..."
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
goto :dl_final
) else (
if defined fallback_ (call :dk_color %Red% "Installing GenuineTicket.xml            [Successful] [Fallback method: clipup -v -o]"
) else (echo Installing GenuineTicket.xml            [Successful]
)
)

::==========================================================================================================================================

if defined changekey (
set winos=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "winos=%%b"
)

echo:
echo Activating...
echo:

call :dk_act
call :dk_checkperm
if defined _perm (
call :dk_color %Green% "%winos% is permanently activated."
goto :dl_final
)

::  Refresh some services and license status

if %_lock%==1 set _retry=1
if defined _intcon set _retry=1

if defined _retry (
for %%# in (wlidsvc LicenseManager sppsvc) do (net stop %%# /y %nul% & net start %%# /y %nul%)
call :dk_refresh
call :dk_act
cscript //nologo %windir%\system32\slmgr.vbs /ato %nul%
)

::  Check license status reason with wmi query, activation command errorlevel gives incorrect result in older builds of Windows 10

set _status=0
if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%wApp%' and PartialProductKey is not null) get LicenseStatusReason /VALUE" 2^>nul') do set "_status=%%a"
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%a in ('%_psc% "(([WMISEARCHER]'SELECT LicenseStatusReason FROM %slp% WHERE ApplicationID=''%wApp%'' AND PartialProductKey IS NOT NULL').Get()).LicenseStatusReason | %% {echo ('LicenseStatusReason='+$_)}" 2^>nul') do set "_status=%%a"
cmd /c exit /b %_status%

if %_status% NEQ 0 set "error_code=[Error Code: 0x!=ExitCode!]"

call :dk_checkperm

if defined _perm (
call :dk_color %Green% "%winos% is permanently activated."
) else (
call :dk_color %Red% "Activation Failed %error_code%"
call :dk_color %Magenta% "Try again / Restart system / Check troubleshooting steps in ReadMe"
)

::========================================================================================================================================

:dl_final

echo:
if exist "%temp_%\.*" rmdir /s /q "%temp_%\" %nul%
if exist "%temp_%\" (
call :dk_color %Red% "Cleaning Temp Files                     [Failed]"
) else (
echo Cleaning Temp Files                     [Successful]
)

::  Rolling back services startup type back to disabled

set serv_rsts=
set serv_rste=

if defined serv_csts (
for %%# in (%serv_csts%) do (
sc config %%# start= disabled %nul% && (
if defined serv_rsts (set "serv_rsts=!serv_rsts! %%#") else (set "serv_rsts=%%#")
) || (
if defined serv_rste (set "serv_cste=!serv_rste! %%#") else (set "serv_rste=%%#")
)
)
)

if defined serv_rsts echo Reverting Services Back To Disabled     [Successful] [%serv_rsts%]
if defined serv_rste call :dk_color %Red% "Reverting Services Back To Disabled     [Failed] [%serv_rste%]"

goto :dk_done

::========================================================================================================================================

::  A lean and mean snippet to set registry ownership and permission recursively
::  Written by @AveYo aka @BAU
::  pastebin.com/XTPt0JSC

::  Modified by @abbodi1406 to make it work in ARM64 Windows 10 (builds older than 21277) where only x86 version of PowerShell is installed.

::  This code runs only if Lockbox method is manually set by the user in this script.

:regown

%_psc% $A='%~1','%~2','%~3','%~4','%~5','%~6';iex(([io.file]::ReadAllText('!_batp!')-split':Own1\:.*')[1])&exit/b:Own1:
$D1=[uri].module.gettype('System.Diagnostics.Process')."GetM`ethods"(42) |where {$_.Name -eq 'SetPrivilege'} #`:no-ev-warn
'SeSecurityPrivilege','SeTakeOwnershipPrivilege','SeBackupPrivilege','SeRestorePrivilege'|foreach {$D1.Invoke($null, @("$_",2))}
$path=$A[0]; $rk=$path-split'\\',2; switch -regex ($rk[0]){'[mM]'{$hv=2147483650};'[uU]'{$hv=2147483649};default{$hv=2147483648};}
$HK=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hv, 256); $s=$A[1]; $sps=[Security.Principal.SecurityIdentifier]
$u=($A[2],'S-1-5-32-544')[!$A[2]];$o=($A[3],$u)[!$A[3]];$w=$u,$o |% {new-object $sps($_)}; $old=!$A[3];$own=!$old; $y=$s-eq'all'
$rar=new-object Security.AccessControl.RegistryAccessRule( $w[0], ($A[5],'FullControl')[!$A[5]], 1, 0, ($A[4],'Allow')[!$A[4]] )
$x=$s-eq'none';function Own1($k){$t=$HK.OpenSubKey($k,2,'TakeOwnership');if($t){0,4|%{try{$o=$t.GetAccessControl($_)}catch{$old=0}
};if($old){$own=1;$w[1]=$o.GetOwner($sps)};$o.SetOwner($w[0]);$t.SetAccessControl($o); $c=$HK.OpenSubKey($k,2,'ChangePermissions')
$p=$c.GetAccessControl(2);if($y){$p.SetAccessRuleProtection(1,1)};$p.ResetAccessRule($rar);if($x){$p.RemoveAccessRuleAll($rar)}
$c.SetAccessControl($p);if($own){$o.SetOwner($w[1]);$t.SetAccessControl($o)};if($s){$($subkeys=$HK.OpenSubKey($k).GetSubKeyNames()) 2>$null;
foreach($n in $subkeys){Own1 "$k\$n"}}}};Own1 $rk[1];if($env:VO){get-acl Registry::$path|fl} #:Own1: lean & mean snippet by AveYo

::========================================================================================================================================

::  Check Windows permanent activation status

:dk_checkperm

if %winbuild% LSS 22483 wmic path %slp% where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name /value 2>nul | findstr /i "Windows" 1>nul && set _perm=1||set _perm=
if %winbuild% GEQ 22483 %_psc% "(([WMISEARCHER]'SELECT Name FROM %slp% WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL').Get()).Name | %% {echo ('Name='+$_)}" 2>nul | findstr /i "Windows" 1>nul && set _perm=1||set _perm=
exit /b

::  Refresh license status

:dk_refresh

if %winbuild% LSS 22483 wmic path %sls% where __CLASS='%sls%' call RefreshLicenseStatus %nul%
if %winbuild% GEQ 22483 %_psc% "$null=(([WMICLASS]'%sls%').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::  Check Windows installed key channel

:dk_channel

if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get ProductKeyChannel /value 2^>nul') do set "_channel=%%#"
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%# in ('%_psc% "(([WMISEARCHER]'SELECT ProductKeyChannel FROM %slp% WHERE ApplicationID=''%wApp%'' AND PartialProductKey IS NOT NULL').Get()).ProductKeyChannel | %% {echo ('ProductKeyChannel='+$_)}" 2^>nul') do set "_channel=%%#"
exit /b

::  Activation command

:dk_act

if %winbuild% LSS 22483 wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" call Activate %nul%
if %winbuild% GEQ 22483 %_psc% "(([WMISEARCHER]'SELECT ID FROM %slp% WHERE ApplicationID=''%wApp%'' AND PartialProductKey IS NOT NULL').Get()).Activate()" %nul%
exit /b

::========================================================================================================================================

:dk_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
if not exist %_psc% (echo %~3) else (%_psc% write-host -back '%1' -fore '%2' '%3')
)
exit /b

:dk_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
if not exist %_psc% (echo %~3%~6) else (%_psc% write-host -back '%1' -fore '%2' '%3' -NoNewline; write-host -back '%4' -fore '%5' '%6')
)
exit /b

::========================================================================================================================================

:dk_done

echo:
if %_unattended%==1 timeout /t 2 & exit /b
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
pause >nul
exit /b

::========================================================================================================================================

::  1st column = Activation ID
::  2nd column = Generic Retail/OEM/MAK Key
::  3rd column = SKU ID
::  4th column = 1 = activation is not working (at the time of writing this), 0 = activation is working
::  5th column = Key Type
::  6th column = WMI Edition ID
::  7th column = Version name incase same Edition ID is used in different OS versions with different key
::  Separator  = _

::  Key preference is in the following order. Retail > OEM:NONSLP > OEM:DM > Volume:MAK

:hwiddata

for %%# in (
8b351c9c-f398-4515-9900-09df49427262_XGVPP-NMH47-7TTHJ-W3FW7-8HV2C___4_0_OEM:NONSLP_Enterprise
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
d06934ee-5448-4fd1-964a-cd077618aa06_43TBQ-NH92J-XKTM7-KT3KK-P39PB_125_0_OEM:NONSLP_EnterpriseS_2019
706e0cfd-23f4-43bb-a9af-1a492b9f1302_NK96Y-D9CD8-W44CQ-R8YTK-DYJWX_125_0_OEM:NONSLP_EnterpriseS_2016
faa57748-75c8-40a2-b851-71ce92aa8b45_FWN7H-PF93Q-4GGP8-M8RF3-MDWWW_125_0_OEM:NONSLP_EnterpriseS_2015
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
1ca0bfa8-d96b-4815-a732-7756f30c29e2_FV469-WGNG4-YQP66-2B2HY-KD8YX_171_1_OEM:NONSLP_EnterpriseG
8d6f6ffe-0c30-40ec-9db2-aad7b23bb6e3_FW7NV-4T673-HF4VX-9X4MM-B4H4T_172_1_OEM:NONSLP_EnterpriseGN
df96023b-dcd9-4be2-afa0-c6c871159ebe_NJCF7-PW8QT-3324D-688JX-2YV66_175_0_____Retail_ServerRdsh
d4ef7282-3d2c-4cf0-9976-8854e64a8d1e_V3WVW-N2PV2-CGWC3-34QGF-VMJ2C_178_0_____Retail_Cloud
af5c9381-9240-417d-8d35-eb40cd03e484_NH9J3-68WK7-6FB93-4K3DF-DJ4F6_179_0_____Retail_CloudN
c7051f63-3a76-4992-bce5-731ec0b1e825_2HN6V-HGTM8-6C97C-RK67V-JQPFD_183_1_____Retail_CloudE
8ab9bdd1-1f67-4997-82d9-8878520837d9_XQQYW-NFFMW-XJPBH-K8732-CKFFD_188_0_____OEM:DM_IoTEnterprise
ed655016-a9e8-4434-95d9-4345352c2552_QPM6N-7J2WJ-P88HH-P3YRH-YY74H_191_0_OEM:NONSLP_IoTEnterpriseS
d4bdc678-0a4b-4a32-a5b3-aaa24c3b0f24_K9VKN-3BGWV-Y624W-MCRMQ-BHDCD_202_0_____Retail_CloudEditionN
92fb8726-92a8-4ffc-94ce-f82e07444653_KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W_203_0_____Retail_CloudEdition
) do (
for /f "tokens=1-7 delims=_" %%A in ("%%#") do if %osSKU%==%%C (

if %1==attempt1 if not defined key echo "!applist!" | find /i "%%A" 1>nul && (set "key=%%B" & set "_chan=%%E" & if %%D==1 set notworking=1)

if %1==attempt2 if not defined key (
set "actidnotfound= [Mismatched Act-ID]"
set 7th=%%G
if not defined 7th (
set "key=%%B" & set "_chan=%%E" & if %%D==1 set notworking=1
) else (
echo "%winos%" | find "%%G" 1>nul && (set "key=%%B" & set "_chan=%%E" & if %%D==1 set notworking=1)
)
)
)
)
exit /b

::========================================================================================================================================