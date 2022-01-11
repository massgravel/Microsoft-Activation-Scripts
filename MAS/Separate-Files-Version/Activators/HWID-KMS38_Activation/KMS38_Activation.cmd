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



::  To activate, run the script with /a parameter or change 0 to 1 in below line
set _acti=0

::  To activate along with KMS38 protection (from being replaced by 180 days KMS activation), 
::  run the script with /ap parameter or change 0 to 1 in below line
set _prot=0

::  To only generate GenuineTicket.xml, run the script with /g parameter or change 0 to 1 in below line
set _gent=0

::  To uninstall the KMS38, run the script with /x parameter or change 0 to 1 in below line
set _unin=0



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
title  KMS38 Activation

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/a"  set _acti=1
if /i "%%A"=="/ap" set _prot=1
if /i "%%A"=="/g"  set _gent=1
if /i "%%A"=="/x"  set _unin=1
if /i "%%A"=="-el" set _elev=1
)
)

for %%A in (%_acti% %_prot% %_gent% %_unin%) do (if "%%A"=="1" set _unattended=1)

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

set _k38=
set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
set slp=SoftwareLicensingProduct
set sls=SoftwareLicensingService
set wApp=55c92734-d682-4d71-983e-d6ec3f16059f
if %~z0 GEQ 1500000 (set "_exitmsg=Go back") else (set "_exitmsg=Exit")
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set "notifytocheckupdate=if %winbuild% GTR 19044 echo Make sure you are using updated version of the script."

::========================================================================================================================================

if %winbuild% LSS 14393 (
%eline%
echo Unsupported OS version detected.
echo Project is supported for Windows 10/11/Server, build 14393 and later.
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

if %_unin%==1 goto :k_uninstall

:k_menu

if %_unattended%==0 (
cls
mode 76, 25
title  KMS38 Activation

echo:
echo:
echo:
echo         ____________________________________________________________
echo:
echo                 [1] KMS38 Activation
echo:
echo                 [2] KMS38 Activation ^+ Protection
echo                 ____________________________________________
echo:
echo                 [3] Generate Ticket
echo                 ____________________________________________
echo:
echo                 [4] Complete Uninstall
echo:
echo                 [5] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "              " %_Green% "Enter a menu option in the Keyboard [1,2,3,4,5]"
choice /C:12345 /N
set _el=!errorlevel!
if !_el!==5  exit /b
if !_el!==4  (
cls
echo:
call :dk_color2 %_White% " " %_Green% "KMS38 Complete Uninstall:"
echo:
echo  It requires below steps,
echo  - Cleaning ClipSVC licences
echo  - Deleting a volatile and protected registry key by taking ownership
echo  - Microsoft Account and Store Apps may need relogin-restart in the system
echo  - Removing Specific KMS Host for Windows
echo:
echo  Its recommended to restart the system after this process.
echo:
echo:
choice /C:12 /N /M "> [1] Continue [2] Go back : "
if errorlevel 2 goto :k_menu
if errorlevel 1 goto :k_uninstall
)
if !_el!==3  (set _gent=1&goto :k_menu2)
if !_el!==2  (
cls
echo:
call :dk_color2 %_White% " " %_Green% "KMS38 Protection:"
echo:
echo  It stops 180 days KMS Activation from replacing KMS38 activation.
echo  Protection requires permission altering of a registry key.
echo:
echo  If you are going to use KMS_VL_ALL or MAS's KMS activation for Office,
echo  then you don't need to enable this protection.  
echo  For more info, check readme.
echo:
echo:
choice /C:12 /N /M "> [1] Continue [2] Go back : "
if errorlevel 2 goto :k_menu
if errorlevel 1 (set _prot=1&goto :k_menu2)
)
if !_el!==1  (set _prot=0&goto :k_menu2)
goto :k_menu
)

:k_menu2

cls
mode 102, 30
if %_gent%==1 (set _title=title  Generate KMS38 GenuineTicket.xml) else (set _title=title  KMS38 Activation)
%_title%

::========================================================================================================================================

if not exist %_psc% if %_prot%==1 (
set _prot=0
set _gent=0
%nceline%
echo Powershell is not installed in the system.
echo It is required to protect KMS38 activation.
echo You need to set the script to the default.
if %_unattended%==0 (
echo:
call :dk_color %_Yellow% "Press any key to go back..."
pause >nul
goto k_menu
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
goto k_menu
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

::  Check if GVLK (KMS key) is already installed or not

set _gvlk=
call :dk_channel
echo: %_channel%| findstr /i "GVLK" 1>nul && set _gvlk=1

::========================================================================================================================================

::  Detect Key

set app=
set key=

if defined applist call :kms38data getkey

::========================================================================================================================================

if not defined key if not defined _gvlk (
%eline%
if not defined applist (
echo Failed running WMI query check, verify that these services are working correctly
echo Windows Management Instrumentation [WinMgmt], Software Protection [sppsvc]
echo:
echo Try troublshooting steps listed in the ReadMe.
goto dk_done
)
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
echo Unable to find this product in the supported product list.
%notifytocheckupdate%
echo:
echo However, if you would like to try KMS38 activation on this product then,
echo install KMS key [GVLK] for this product and run the script.
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
echo Since KMS key is already installed, script will activate with KMS38.
if %_unattended%==0 (
echo:
call :dk_color %_Yellow% "Press any key to continue..."
pause >nul
)
cls
)

::========================================================================================================================================

::  Making sure we have activation ID if not defined yet

if not defined key set app=
if not defined app (
if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get ID /VALUE" 2^>nul') do call set "app=%%a"
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%a in ('%_psc% "(([WMISEARCHER]'SELECT ID FROM %slp% WHERE ApplicationID=''%wApp%'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL').Get()).ID | %% {echo ('ID='+$_)}" 2^>nul') do call set "app=%%a"
)

if not defined app (
%eline%
echo [%winos% ^| SKU:%osSKU% ^| %winbuild%]
echo Unable to detect activation ID.
%notifytocheckupdate%
goto dk_done
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

::========================================================================================================================================

:: Check clipup.exe for the detection and activation of server cor and acor editions

set a_cor=
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" if not exist "%systemroot%\System32\clipup.exe" set a_cor=1

if defined a_cor (
if not exist "!_work!\BIN\clipup.exe" (
%eline%
echo 'clipup.exe' file is missing in 'BIN' folder. Aborting...
goto dk_done
)
)

::========================================================================================================================================

cls
echo:
echo Checking OS Info                        [%winos% ^| %winbuild% ^| %arch%]

::========================================================================================================================================

echo:
set "_serv=ClipSVC sppsvc Winmgmt"

::  Client License Service (ClipSVC)
::  Software Protection
::  Windows Management Instrumentation

echo Checking Services                       [%_serv%]

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
)

if not defined applist (
call :dk_color %Red% "Checking WMI Query                      [Failed]"
) else (
echo Checking WMI Query                      [Successful]
)

::========================================================================================================================================

::  Install key

echo:
set _partial=
if defined key set _ipartial=%key:~-5%

if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get PartialProductKey /value 2^>nul') do set "_partial=%%#"
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%# in ('%_psc% "(([WMISEARCHER]'SELECT PartialProductKey FROM %slp% WHERE ApplicationID=''%wApp%'' AND PartialProductKey IS NOT NULL').Get()).PartialProductKey | %% {echo ('PartialProductKey='+$_)}" 2^>nul') do set "_partial=%%#"

if defined key if /i "%_partial%"=="%_ipartial%" (
echo Checking Installed Product Key          [%key%] [Volume:GVLK]
)

if not defined key (
echo Checking Installed Product Key          [Partial Key - %_partial%] [Volume:GVLK]
)

if defined key if /i not "%_partial%"=="%_ipartial%" (
if %winbuild% LSS 22483 wmic path %sls% where __CLASS='%sls%' call InstallProductKey ProductKey="%key%" %nul%
if %winbuild% GEQ 22483 %_psc% "(([WMISEARCHER]'SELECT Version FROM %sls%').Get()).InstallProductKey('%key%')" %nul%
if not !errorlevel!==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%

if !errorlevel!==0 (
call :dk_refresh
echo Installing KMS Client Setup Key         [%key%] [Successful]
) else (
call :dk_color %Red% "Installing KMS Client Setup Key         [%key%] [Failed]"
)
)

::========================================================================================================================================

::  Set specific KMS host to Local Host
::  By doing this, global KMS IP can not replace KMS38 activation but can be used with Office and other Windows Editions

set regadd=
set k_error=

if not %_gent%==1 (
echo:
set regadd=1
%nul% reg delete "HKLM\%SPPk%\%wApp%" /f
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%wApp%" /f

%nul% reg query "HKLM\%SPPk%\%wApp%" && (
call :regown "HKLM\%SPPk%\%wApp%"
%nul% reg delete "HKLM\%SPPk%\%wApp%" /f
)

%nul% reg add "HKLM\%SPPk%\%wApp%\%app%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2" || set k_error=1
%nul% reg add "HKLM\%SPPk%\%wApp%\%app%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" || set k_error=1

if not defined k_error (
echo Adding Specific KMS Host                [LocalHost 127.0.0.2] [Successful]
) else (
call :dk_color %Red% "Adding Specific KMS Host                [LocalHost 127.0.0.2] [Failed]"
)
)

::========================================================================================================================================

::  Files are copied to temp to generate ticket to avoid possible issues in case the path contains special character or non English names

echo:
set "temp_=%SystemRoot%\Temp\_Temp"
set "_clipup=%systemroot%\System32\clipup.exe"
if exist "%temp_%\.*" rmdir /s /q "%temp_%\" %nul%
md "%temp_%\" %nul%

pushd "!_work!\BIN\"
copy /y /b "%ARM64_file%gatherosstate.exe" "%temp_%\gatherosstate.exe" %nul%
copy /y /b "%ARM64_file%slc.dll" "%temp_%\slc.dll" %nul%
if defined a_cor copy /y /b "ClipUp.exe" "%_clipup%" %nul%
popd

set copyf=
if not exist "%temp_%\gatherosstate.exe" set copyf=1
if not exist "%temp_%\slc.dll" set copyf=1

if defined copyf (
call :dk_color %Red% "Copying Required Files to Temp          [%temp_%] [Failed]"
goto :k_final
) else (
echo Copying Required Files to Temp          [%temp_%] [Successful]
)

if defined a_cor (
if exist "%_clipup%" (
echo Copying clipup.exe File to              [%systemroot%\System32\] [Successful]
) else (
call :dk_color %Red% "Copying clipup.exe File to              [%systemroot%\System32\] [Failed] Aborting..."
goto :k_final
)
)

::========================================================================================================================================

::  Multiple attempts to generate the ticket because in some cases, one attempt is not enough.

set "_noxml=if not exist "%temp_%\GenuineTicket.xml""

start /wait "" "%temp_%/gatherosstate.exe" %nul%
%_noxml% timeout /t 3 %nul%
%_noxml% net stop sppsvc /y %nul%
%_noxml% call "%temp_%/gatherosstate.exe" %nul%
%_noxml% timeout /t 3 %nul%
%_noxml% "%temp_%/gatherosstate.exe" %nul%
%_noxml% timeout /t 3 %nul%

%_noxml% (
call :dk_color %Red% "Generating GenuineTicket.xml            [Failed] Aborting..."
goto :k_final
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
goto :k_final
)

::========================================================================================================================================

::  clipup -v -o -altto <Ticket path> method to apply ticket is not used to avoid the certain issues in case if the username have 
::  spaces / special characters / non English names

set "tdir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
copy /y /b "%temp_%\GenuineTicket.xml" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "Failed to copy Ticket to [%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket\]"
goto :k_final
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
goto :k_final
) else (
if defined fallback_ (call :dk_color %Red% "Installing GenuineTicket.xml            [Successful] [Fallback method: clipup -v -o]"
) else (echo Installing GenuineTicket.xml            [Successful]
)
)

::==========================================================================================================================================

echo:
echo Activating...
echo:

call :k_checkexp
if defined _k38 (
call :k_actinfo
goto :k_final
)

::  Clear 180 Days KMS Activation lock with Windows SKU specific rearm and without the need to restart the system

if %winbuild% LSS 22483 wmic path %slp% where ID='%app%' call ReArmsku %nul%
if %winbuild% GEQ 22483 %_psc% "$null=([WMI]'%slp%=''%app%''').ReArmsku()" %nul%

if %errorlevel%==0 (
echo Applying SKU-ID Rearm                   [Successful]
) else (
call :dk_color %Red% "Applying SKU-ID Rearm                   [Failed]"
)

echo:
call :k_checkexp
if defined _k38 (
call :k_actinfo
goto :k_final
)

::  Restart software protection service to refresh itself and run refresh license status and activation commands

net stop sppsvc /y %nul%
net start sppsvc /y %nul%
call :dk_refresh
call :dk_act

call :k_checkexp
if defined _k38 (
call :k_actinfo
goto :k_final
)

call :dk_color %Red% "Activation Failed"
call :dk_color %Magenta% "Try again / Restart system / Check troubleshooting steps in ReadMe"

::========================================================================================================================================

:k_final

::  Remove the added Specific KMS Host (Local Host) if activation is not completed

echo:
set k_error=
if defined regadd if not defined _k38 (
%nul% reg delete "HKLM\%SPPk%\%wApp%" /f
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%wApp%" /f
%nul% reg query "HKLM\%SPPk%\%wApp%" && set k_error=1
%nul% reg query "HKU\S-1-5-20\%SPPk%\%wApp%" && set k_error=1
if not defined k_error (
echo Removing The Added Specific KMS Host    [Successful]
) else (
call :dk_color %Red% "Removing The Added Specific KMS Host    [Failed]"
)
)

::  Protect KMS38 if opted by the user and conditions are correct

if defined regadd if defined _k38 if %_prot%==1 (
%nul% call :regown "HKLM\%SPPk%\%wApp%" "" S-1-5-32-544 "" Deny "SetValue,Delete"
%nul% reg delete "HKLM\%SPPk%\%wApp%" /f
%nul% reg query "HKLM\%SPPk%\%wApp%" && (
call :dk_color %Gray% "Locking a Registry To Protect KMS38     [Successful]"
) || (
call :dk_color %Red% "Locking a Registry To Protect KMS38     [Failed]"
)
)

::  clipup.exe does not exist in server cor and acor editions by default, it was copied there with this script

if exist "%temp_%\.*" rmdir /s /q "%temp_%\" %nul%
if defined a_cor if exist "%_clipup%" del /f /q "%_clipup%" %nul%

if exist "%temp_%\" (
call :dk_color %Red% "Cleaning Temp Files                     [Failed]"
) else (
echo Cleaning Temp Files                     [Successful]
)

if defined a_cor (
if exist "%_clipup%" (
call :dk_color %Red% "Deleting copied clipup.exe file         [Failed]"
) else (
echo Deleting copied clipup.exe file         [Successful]
)
)

goto :dk_done

::========================================================================================================================================

:k_uninstall

cls
mode 76, 25
title  KMS38 Complete Uninstall
set "RegKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState"

echo:
call :k_checkexp

if not defined _k38 (
echo Windows is not activated with KMS38. Skipping cleaning ClipSVC Licences...
echo:
)

if defined _k38 (
for %%# in (ClipSVC) do (
sc query %%# | find /i "STOPPED" %nul% || net stop %%# /y %nul%
sc query %%# | find /i "STOPPED" %nul% || sc stop %%# %nul%
)

REM  Thanks to @mspaintmsi for informing this command info

rundll32 clipc.dll,ClipCleanUpState

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat" (
call :dk_color %Red% "Cleaning ClipSVC Licences               [Failed]"
) else (
echo Cleaning ClipSVC Licences               [Successful]
)

REM  Below registry key (Volatile & Protected) gets created after the ClipSVC License cleanup command, and gets automatically deleted after 
REM  system restart. It needs to be deleted to activate the system without restart.

call :regown "%RegKey%" %nul%
reg delete "%RegKey%" /f %nul%

reg query "%RegKey%" %nul% && (
call :dk_color %Red% "Deleting a Volatile Registry            [Failed]"
call :dk_color %Magenta% "Restart the system, that will delete this registry key automatically"
) || (
echo Deleting a Volatile Registry            [Successful]
)

for %%# in (ClipSVC wlidsvc LicenseManager sppsvc) do (net stop %%# /y %nul% & net start %%# /y %nul%)
call :dk_refresh
)

set exist_=
%nul% reg query "HKLM\%SPPk%\%wApp%" && (
set exist_=1
%nul% reg delete "HKLM\%SPPk%\%wApp%" /f
)
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%wApp%" /f

%nul% reg query "HKLM\%SPPk%\%wApp%" && (
%nul% call :regown "HKLM\%SPPk%\%wApp%"
%nul% reg delete "HKLM\%SPPk%\%wApp%" /f
)

%nul% reg query "HKLM\%SPPk%\%wApp%" && (
call :dk_color %Red% "Removing Specific KMS Host              [Failed]"
) || (
if defined exist_ (
echo Removing Specific KMS Host              [Successful]
) else (
echo Removing Specific KMS Host              [Already Removed]
)
)

if defined _k38 (
echo:
call :dk_color %Magenta% "Its recommended to restart the system now."
)

goto :dk_done

::========================================================================================================================================

::  A lean and mean snippet to set registry ownership and permission recursively
::  Written by @AveYo aka @BAU
::  pastebin.com/XTPt0JSC

::  Modified by @abbodi1406 to make it work in ARM64 Windows 10 (builds older than 21277) where only x86 version of PowerShell is installed.

::  This code runs only if KMS38 protection option or complete uninstall option is used by the user in this script.

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

::  Check KMS activation status

:k_actinfo

set xpr=
if not exist %_psc% (
cscript //nologo %windir%\system32\slmgr.vbs /xpr
) else (
for /f "tokens=* delims=" %%# in ('%_psc% "$([DateTime]::Now.addMinutes(%gpr%)).ToString('yyyy-MM-dd HH:mm:ss')" 2^>nul') do set "xpr=%%#"
call :dk_color %Green% "%winos% is activated till !xpr!"
)
exit /b

::  Check remaining KMS activation grace period

:k_checkexp

set gpr=0
if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%# in ('"wmic path %slp% where (ApplicationID='%wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" 2^>nul') do set "gpr=%%#"
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%# in ('%_psc% "(([WMISEARCHER]'SELECT GracePeriodRemaining FROM %slp% WHERE ApplicationID=''%wApp%'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL').Get()).GracePeriodRemaining | %% {echo ('GracePeriodRemaining='+$_)}" 2^>nul') do set "gpr=%%#"
if %gpr% GTR 259200 (set _k38=1) else (set _k38=)
exit /b

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
::  2nd column = GVLK (Generic volume licensing key)
::  3rd column = SKU ID
::  4th column = WMI Edition ID (For reference only)
::  5th column = Build Branch name incase same Edition ID is used in different OS versions with different key (For reference only)
::  Separator  = "_"

:kms38data

for %%# in (
73111121-5638-40f6-bc11-f1d7b0d64300_NPPR9-FWDCX-D2C8J-H872K-2YT43___4_Enterprise
9bd77860-9b31-4b7b-96ad-2564017315bf_VDYBN-27WPP-V4HQT-9VMD4-VMK7H___7_ServerStandard_FE
de32eafd-aaee-4662-9444-c1befb41bde2_N69G4-B89J2-4G8F4-WWYCC-J464C___7_ServerStandard_RS5
8c1c5410-9f39-4805-8c9d-63a07706358f_WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY___7_ServerStandard_RS1
ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03_WX4NM-KYWYW-QJJR4-XV3QB-6VM33___8_ServerDatacenter_FE
34e1ae55-27f8-4950-8877-7a03be5fb181_WMDGN-G9PQG-XVVXX-R3X43-63DFG___8_ServerDatacenter_RS5
21c56779-b449-4d20-adfc-eece0e1ad74b_CB7KF-BWN84-R7R2Y-793K2-8XDDG___8_ServerDatacenter_RS1
e272e3e2-732f-4c65-a8f0-484747d0d947_DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4__27_EnterpriseN
2de67392-b7a7-462a-b1ca-108dd189f588_W269N-WFGWX-YVC9B-4J6C9-T83GX__48_Professional
a80b5abf-76ad-428b-b05d-a47d2dffeebf_MH37W-N47XK-V7XM9-C7227-GCQG9__49_ProfessionalN
034d3cbb-5d4b-4245-b3f8-f84571314078_WVDHN-86M7X-466P6-VHXV7-YY726__50_ServerSolution_RS5
2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283_JCKRF-N37P4-C2D82-9YXRT-4M63B__50_ServerSolution_RS1
7b9e1751-a8da-4f75-9560-5fadfe3d8e38_3KHY7-WNT83-DGQKR-F7HPR-844BM__98_CoreN
a9107544-f4a0-4053-a96a-1479abdef912_PVMJN-6DFY6-9CCP6-7BKTT-D3WVR__99_CoreCountrySpecific
cd918a57-a41b-4c82-8dce-1a538e221a83_7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH_100_CoreSingleLanguage
58e97c99-f377-4ef1-81d5-4ad5522b5fd8_TX9XD-98N7V-6WMQ6-BX7FG-H8Q99_101_Core
7b4433f4-b1e7-4788-895a-c45378d38253_QN4C6-GBJD2-FB422-GHWJK-GJG2R_110_ServerCloudStorage
8de8eb62-bbe0-40ac-ac17-f75595071ea3_GRFBW-QNDC4-6QBHG-CCK3B-2PR88_120_ServerARM64_RS5
43d9af6e-5e86-4be8-a797-d072a046896c_K9FYF-G6NCK-73M32-XMVPY-F9DRR_120_ServerARM64_RS4
e0c42288-980c-4788-a014-c080d2e1926e_NW6C2-QMPVW-D7KKK-3GKT6-VCFB2_121_Education
3c102355-d027-42c6-ad23-2e7ef8a02585_2WH4N-8QGBV-H22JP-CT43Q-MDWWJ_122_EducationN
32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7XTQ-FN8P6-TTKYV-9D4CC-J462D_125_EnterpriseS_RS5,VB
2d5a5a60-3040-48bf-beb0-fcd770c20ce0_DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ_125_EnterpriseS_RS1
7b51a46c-0c04-4e8f-9af4-8496cca90d5e_WNMTR-4C88C-JK8YV-HQ7T2-76DF9_125_EnterpriseS_TH1
7103a333-b8c8-49cc-93ce-d37c09687f92_92NFX-8DJQP-P6BBQ-THF9C-7CG2H_126_EnterpriseSN_RS5,VB
9f776d83-7156-45b2-8a5c-359b9c9f22a3_QFFDN-GRT3P-VKWWX-X7T3R-8B639_126_EnterpriseSN_RS1
87b838b7-41b6-4590-8318-5797951d8529_2F77B-TNFGY-69QQF-B8YKP-D69TJ_126_EnterpriseSN_TH1
39e69c41-42b4-4a0a-abad-8e3c10a797cc_QFND9-D3Y9C-J3KKY-6RPVP-2DPYV_145_ServerDatacenterACor_FE
90c362e5-0da1-4bfd-b53b-b87d309ade43_6NMRW-2C8FM-D24W7-TQWMY-CWH2D_145_ServerDatacenterACor_RS5
e49c08e7-da82-42f8-bde2-b570fbcae76c_2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG_145_ServerDatacenterACor_RS3
f5e9429c-f50b-4b98-b15c-ef92eb5cff39_67KN8-4FYJW-2487Q-MQ2J7-4C4RG_146_ServerStandardACor_FE
73e3957c-fc0c-400d-9184-5f7b6f2eb409_N2KJX-J94YW-TQVFB-DG9YT-724CC_146_ServerStandardACor_RS5
61c5ef22-f14f-4553-a824-c4b31e84b100_PTXN8-JFHJM-4WC78-MPCBR-9W4KR_146_ServerStandardACor_RS3
82bbc092-bc50-4e16-8e18-b74fc486aec3_NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J_161_ProfessionalWorkstation
4b1571d3-bafb-4b40-8087-a961be2caf65_9FNHH-K3HBT-3W4TD-6383H-6XYWF_162_ProfessionalWorkstationN
3f1afc82-f8ac-4f6c-8005-1d233e606eee_6TP4R-GNPTD-KYYHQ-7B7DP-J447Y_164_ProfessionalEducation
5300b18c-2e33-4dc2-8291-47ffcec746dd_YVWGF-BXNMC-HTQYQ-CPQ99-66QFC_165_ProfessionalEducationN
8c8f0ad3-9a43-4e05-b840-93b8d1475cbc_6N379-GGTMK-23C6M-XVVTC-CKFRQ_168_ServerAzureCor_FE
a99cc1f0-7719-4306-9645-294102fbff95_FDNH6-VW9RW-BXPJ7-4XTYG-239TB_168_ServerAzureCor_RS5
3dbf341b-5f6c-4fa7-b936-699dce9e263f_VP34G-4NPPG-79JTQ-864T4-R3MQX_168_ServerAzureCor_RS1
e0b2d383-d112-413f-8a80-97f373a5820c_YYVX9-NTFWV-6MDM3-9PT4T-4M68B_171_EnterpriseG
e38454fb-41a4-4f59-a5dc-25080e354730_44RPN-FTY23-9VTTB-MP9BX-T84FV_172_EnterpriseGN
ec868e65-fadf-4759-b23e-93fe37f2cc29_CPWHC-NT2C7-VYW78-DHDB2-PG3GK_175_ServerRdsh_RS5
e4db50ea-bda1-4566-b047-0ca50abc6f07_7NBT4-WGBQX-MP4H7-QXFF8-YP3KX_175_ServerRdsh_RS3
0df4f814-3f57-4b8b-9a9d-fddadcd69fac_NBTWJ-3DR69-3C4V8-C26MC-GQ9M6_183_CloudE
d30136fc-cb4b-416e-a23d-87207abc44a9_6XN7V-PCBDC-BDBRH-8DQY7-G6R44_202_CloudEditionN
ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69_37D7F-N49CB-WQR8W-TBJ73-FM8RX_203_CloudEdition
19b5e0fb-4431-46bc-bac1-2f1873e4ae73_NTBV8-9K7Q8-V27C6-M2BTV-KHMXV_407_ServerTurbine
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do if %osSKU%==%%C (
if %1==getkey if not defined key echo "!applist!" | find /i "%%A" >nul && (set "app=%%A" & set "key=%%B")
)
)
exit /b

::========================================================================================================================================