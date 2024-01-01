@set masver=2.5
@setlocal DisableDelayedExpansion
@echo off



::============================================================================
::
::   This script is a part of 'Microsoft_Activation_Scripts' (MAS) project.
::
::   Homepage: mass grave[.]dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================



::  To stage current edition while changing edition with CBS Upgrade Method, change 0 to 1 in below line
set _stg=0



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
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "1" /f %nul1%
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
title  Change Windows Edition %masver%

set _args=
set _elev=

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="-el"                    set _elev=1
)
)

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

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
set "line=echo ___________________________________________________________________________________________"
if %~z0 GEQ 200000 (set "_exitmsg=Go back") else (set "_exitmsg=Exit")

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected [%winbuild%].
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
goto ced_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto ced_done
)

::========================================================================================================================================

::  Fix special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%userprofile%\AppData\Local\Temp"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
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

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo To do so, right click on this script and select 'Run as administrator'.
goto ced_done
)

::========================================================================================================================================

::  This code disables QuickEdit for this cmd.exe session only without making permanent changes to the registry
::  It is added because clicking on the script window pauses the operation and leads to the confusion that script stopped due to an error

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
echo [1] Get Latest MAS
echo [0] Continue Anyway
echo:
call :dk_color %_Green% "Enter a menu option in the Keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
)

::========================================================================================================================================

cls
mode 98, 30

echo:
echo Initializing...
echo:
call :dk_product
call :dk_ckeckwmic

::  Show info for potential script stuck scenario

sc start sppsvc %nul%
if %errorlevel% NEQ 1056 if %errorlevel% NEQ 0 (
echo:
echo Error code: %errorlevel%
call :dk_color %Red% "Failed to start [sppsvc] service, rest of the process may take a long time..."
echo:
)

::========================================================================================================================================

::  Check Activation IDs

call :dk_actids
if not defined applist (
net stop sppsvc /y %nul%
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if !errorlevel! NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
call :dk_refresh
call :dk_actids
if not defined applist (
%eline%
echo Activation IDs not found. Aborting...
echo:
echo Check this page for help. %mas%troubleshoot
goto ced_done
)
)

::========================================================================================================================================

call :dk_checksku

if not defined osSKU (
%eline%
echo SKU value was not detected properly. Aborting...
goto ced_done
)

::========================================================================================================================================

::  Check Windows Edition

set osedition=
set dismedition=
set dismnotworking=

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SLGetWindowsInformation', 'slc.dll', 22, 1, [int], @([String], [int], [int].MakeByRefType(), [String].MakeByRefType()), 1, 3);
set d1=%d1% $editionName = 0; [void]$TypeBuilder.CreateType()::SLGetWindowsInformation('Kernel-EditionName', 0, [ref]0, [ref]$editionName); $editionName
if %winbuild% GEQ 14393 for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set osedition=%%s)
if "%osedition%"=="0" set osedition=

if not defined osedition (
for /f "skip=2 tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul6%') do set "osedition=%%a"
)

::  Workaround for an issue in builds between 1607 and 1709 where ProfessionalEducation is shown as Professional

if %osSKU%==164 set osedition=ProfessionalEducation
if %osSKU%==165 set osedition=ProfessionalEducationN

for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition %nul6% ^| find /i "Current Edition :"') do set "dismedition=%%a"
if not defined dismedition set dismnotworking=1

if defined dismedition if not defined osedition set osedition=%dismedition%

if not defined osedition (
%eline%
DISM /English /Online /Get-CurrentEdition %nul%
cmd /c exit /b !errorlevel!
echo DISM command failed [Error Code - 0x!=ExitCode!]
echo OS Edition was not detected properly. Aborting...
echo:
echo Check this page for help. %mas%troubleshoot
goto ced_done
)

::========================================================================================================================================

set branch=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildBranch %nul6%') do set "branch=%%b"

::  Check PowerShell

%psc% $ExecutionContext.SessionState.LanguageMode %nul2% | find /i "Full" %nul1% || (
%eline%
%psc% $ExecutionContext.SessionState.LanguageMode
echo:
echo PowerShell is not working. Aborting...
echo If you have applied restrictions on Powershell then undo those changes.
echo:
echo Check this page for help. %mas%troubleshoot
goto ced_done
)

::========================================================================================================================================

::  Get Target editions list

set _target=
set _dtarget=
set _ptarget=
set _ntarget=
set _wtarget=

if %winbuild% GEQ 10240 for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (if defined _dtarget (set "_dtarget= !_dtarget! %%a ") else (set "_dtarget= %%a "))
if %winbuild% LSS 10240 for /f "tokens=4" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':cbsxml\:.*';& ([ScriptBlock]::Create($f[1])) -GetTargetEditions;" ^| findstr /i /c:"Target Edition : "') do (if defined _ptarget (set "_ptarget= !_ptarget! %%a ") else (set "_ptarget= %%a "))

if %winbuild% GEQ 10240 if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
call :ced_edilist
if /i "%osedition:~0,4%"=="Core" (set "_wtarget= Professional !_wtarget! ")
set "_dtarget= %_dtarget% !_wtarget! "
)

::========================================================================================================================================

::  Block the change to/from CloudEdition editions

for %%# in (202 203) do if %osSKU%==%%# (
%eline%
echo [%winos% ^| SKU:%osSKU% ^| %winbuild%]
echo It's not recommended to change this installed edition to any other.
echo Aborting...
goto ced_done
)

for %%# in ( %_dtarget% %_ptarget% ) do if /i not "%%#"=="%osedition%" (
echo "!_target!" | find /i " %%# " %nul1% || set "_target= !_target! %%# "
)

if defined _target (
for %%# in (%_target%) do (
echo %%# | findstr /i "CountrySpecific CloudEdition ServerRdsh" %nul% || (set "_ntarget=!_ntarget! %%#")
)
)

if not defined _ntarget (
%line%
echo:
if defined dismnotworking call :dk_color %Red% "DISM.exe is not responding."
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
call :dk_color %Gray% "You can change the Edition [%osedition%] [%winbuild%] to one of the following."
if defined dismnotworking (
call :dk_color %_Yellow% "Note - DISM.exe is not responding."
if /i "%osedition:~0,4%"=="Core" call :dk_color %_Yellow% "     - You will see more edition options to choose once its changed to Pro."
)
%line%
echo:

for %%A in (%_ntarget%) do (
set /a counter+=1
echo [!counter!]  %%A
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  %_exitmsg%
echo:
call :dk_color %_Green% "Enter option number in keyboard, and press "Enter":"
set /p inpt=
if "%inpt%"=="" goto cedmenu2
if "%inpt%"=="0" exit /b
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto cedmenu2

::========================================================================================================================================

if %winbuild% LSS 10240 goto :cbsmethod
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" goto :ced_change_server

cls
mode con cols=105 lines=32

set key=
set _chan=
set _dismapi=0

::  Check if DISM Api or slmgr.vbs is required for edition upgrade

if not exist "%SystemRoot%\System32\spp\tokens\skus\%targetedition%\" (
set _dismapi=1
)

set "keyflow=Retail OEM:NONSLP OEM:DM Volume:MAK Volume:GVLK"

call :ced_targetSKU %targetedition%
if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%
if not defined key call :changeeditiondata

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Unable to get product key from pkeyhelper.dll
echo:
echo Check this page for help. %mas%troubleshoot
goto ced_done
)

::========================================================================================================================================

::  Changing from Core to Non-Core & Changing editions in Windows build older than 17134 requires "changepk /productkey" or DISM Api method and restart
::  In other cases, editions can be changed instantly with "slmgr /ipk"

if %_dismapi%==1 (
mode con cols=105 lines=40
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':checkrebootflag\:.*';iex ($f[1]);" | find /i "True" %nul% && (
%eline%
echo Pending Reboot flags found.
echo:
echo Restart the system and try again.
goto ced_done
)
)

cls
%line%
echo:
if defined dismnotworking call :dk_color %_Yellow% "DISM.exe is not responding."
echo Changing the Current Edition [%osedition%] %winbuild% to [%targetedition%]
echo:

if %_dismapi%==1 (
call :dk_color %Green% "Notes-"
echo:
echo  - Save your work before continue, system will auto restart.
echo:
echo  - You will need to activate with HWID option once the edition is changed.
%line%
echo:
choice /C:21 /N /M "[1] Continue [2] %_exitmsg% : "
if !errorlevel!==1 exit /b
)

::========================================================================================================================================

if %_dismapi%==0 (
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
echo Check this page for help. %mas%troubleshoot
)
)

if %_dismapi%==1 (
echo:
echo Applying the DISM API method with %_chan% Key %key%. Please wait...
echo:
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':dismapi\:.*';& ([ScriptBlock]::Create($f[1])) %targetedition% %key%;"
timeout /t 3 %nul1%
echo:
call :dk_color %Blue% "Incase of errors, you must restart your system before trying again."
echo Check this page for help. %mas%troubleshoot
)
%line%

goto ced_done

::========================================================================================================================================

:cbsmethod

cls
mode con cols=105 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':checkrebootflag\:.*';iex ($f[1]);" | find /i "True" %nul% && (
%eline%
echo Pending reboot flags found.
echo:
echo Restart the system and try again.
goto ced_done
)

echo:
if defined dismnotworking call :dk_color %_Yellow% "Note - DISM.exe is not responding."
echo Changing the Current Edition [%osedition%] %winbuild% to [%targetedition%]
echo:
call :dk_color %Blue% "Important - Save your work before continue, system will auto reboot."
echo:
choice /C:01 /N /M "[1] Continue [0] %_exitmsg% : "
if %errorlevel%==1 exit /b

echo:
echo Initializing...
echo:

if %_stg%==0 (set stage=) else (set stage=-StageCurrent)
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':cbsxml\:.*';& ([ScriptBlock]::Create($f[1])) -SetEdition %targetedition% %stage%;"
echo:
call :dk_color %Blue% "Incase of errors, you must restart your system before trying again."
echo Check this page for help. %mas%troubleshoot
%line%

goto ced_done

::========================================================================================================================================

:ced_change_server

cls
mode con cols=105 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

set key=
set _chan=
set "keyflow=Volume:GVLK Retail Volume:MAK OEM:NONSLP OEM:DM"

call :ced_targetSKU %targetedition%
if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%
if not defined key call :changeeditiondata

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Unable to get product key from pkeyhelper.dll
echo:
echo Check this page for help. %mas%troubleshoot
goto ced_done
)

::========================================================================================================================================

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':checkrebootflag\:.*';iex ($f[1]);" | find /i "True" %nul% && (
%eline%
echo Pending reboot flags found.
echo:
echo Restart the system and try again.
goto ced_done
)

cls
echo:
if defined dismnotworking call :dk_color %_Yellow% "Note - DISM.exe is not responding."
echo Changing the Current Edition [%osedition%] %winbuild% to [%targetedition%]
echo:
echo Applying the command with %_chan% Key
echo DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula
DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula

call :dk_color %Blue% "You must restart the system at this stage."
echo Help: %mas%troubleshoot

::========================================================================================================================================

:ced_done

echo:
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
pause %nul1%
exit /b

::========================================================================================================================================

::  Check SKU value

:dk_checksku

set osSKU=
set slcSKU=
set wmiSKU=
set regSKU=

if %winbuild% GEQ 14393 (set info=Kernel-BrandingInfo) else (set info=Kernel-ProductInfo)
set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SLGetWindowsInformationDWORD', 'slc.dll', 'Public, Static', 1, [int], @([String], [int].MakeByRefType()), 1, 3);
set d1=%d1% $Sku = 0; [void]$TypeBuilder.CreateType()::SLGetWindowsInformationDWORD('%info%', [ref]$Sku); $Sku
for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set slcSKU=%%s)
if "%slcSKU%"=="0" set slcSKU=
if 1%slcSKU% NEQ +1%slcSKU% set slcSKU=

for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn %nul6%') do set "regSKU=%%a"
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" %nul6%') do if not errorlevel 1 set "wmiSKU=%%a"
if %_wmic% EQU 0 for /f "tokens=1" %%a in ('%psc% "([WMI]'Win32_OperatingSystem=@').OperatingSystemSKU" %nul6%') do if not errorlevel 1 set "wmiSKU=%%a"

set osSKU=%slcSKU%
if not defined osSKU set osSKU=%wmiSKU%
if not defined osSKU set osSKU=%regSKU%
exit /b

::  Refresh license status

:dk_refresh

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call RefreshLicenseStatus %nul%
if %_wmic% EQU 0 %psc% "$null=(([WMICLASS]'SoftwareLicensingService').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::  Get Windows Activation IDs

:dk_actids

set applist=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f') get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f''').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))
exit /b

::  Get Edition list

:ced_edilist

if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f') get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f''').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
%chkedi% do (
call if exist %Systemdrive%\Windows\System32\spp\tokens\skus\%%a (
call set "_wtarget= !_wtarget! %%a "
)
)
exit /b

::  Check wmic.exe

:dk_ckeckwmic

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value %nul2% | find /i "computersystem" %nul1% && set _wmic=1
)
exit /b

::  Get Product name (WMI/REG methods are not reliable in all conditions, hence winbrand.dll method is used)

:dk_product

call :dk_reflection

set d1=%ref% $meth = $TypeBuilder.DefinePInvokeMethod('BrandingFormatString', 'winbrand.dll', 'Public, Static', 1, [String], @([String]), 1, 3);
set d1=%d1% $meth.SetImplementationFlags(128); $TypeBuilder.CreateType()::BrandingFormatString('%%WINDOWS_LONG%%')

set winos=
for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set winos=%%s)
echo "%winos%" | find /i "Windows" %nul1% || (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName %nul6%') do set "winos=%%b"
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

::  Check pending reboot flags

:checkrebootflag:
function Test-PendingReboot
{
 if (Test-Path -Path "$env:windir\WinSxS\pending.xml") { return $true }
 if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA SilentlyContinue) { return $true }
 if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA SilentlyContinue) { return $true }
 try { 
   $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
   $status = $util.DetermineIfRebootPending()
   if(($status -ne $null) -and $status.RebootPending){
     return $true
   }
 }catch{}
 
 return $false
}
Test-PendingReboot
:checkrebootflag:

::========================================================================================================================================

::  Get Product Key from pkeyhelper.dll for future new editions
::  It works on Windows 10 1803 (17134) and later builds.

:dk_pkey

call :dk_reflection

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SkuGetProductKeyForEdition', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([int], [String], [String].MakeByRefType(), [String].MakeByRefType()), 1, 3);
set d1=%d1% $out = ''; [void]$TypeBuilder.CreateType()::SkuGetProductKeyForEdition(%1, %2, [ref]$out, [ref]$null); $out

set pkey=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkey=%%a)
exit /b


::  Get channel name for the key which was extracted from pkeyhelper.dll

:dk_pkeychannel

set k=%1
set m=[Runtime.InteropServices.Marshal]
set p=%SystemRoot%\System32\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('PidGenX', 'pidgenx.dll', 'Public, Static', 1, [int], @([String], [String], [String], [int], [IntPtr], [IntPtr], [IntPtr]), 1, 3);
set d1=%d1% $r = [byte[]]::new(0x04F8); $r[0] = 0xF8; $r[1] = 0x04; $f = %m%::AllocHGlobal(0x04F8); %m%::Copy($r, 0, $f, 0x04F8);
set d1=%d1% [void]$TypeBuilder.CreateType()::PidGenX('%k%', '%p%', '00000', 0, 0, 0, $f); %m%::Copy($f, $r, 0, 0x04F8); %m%::FreeHGlobal($f); [Text.Encoding]::Unicode.GetString($r, 1016, 128)

set pkeychannel=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkeychannel=%%a)
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

call :dk_reflection

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('GetEditionIdFromName', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([String], [int].MakeByRefType()), 1, 3);
set d1=%d1% $out = 0; [void]$TypeBuilder.CreateType()::GetEditionIdFromName('%k%', [ref]$out); $out

for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set targetSKU=%%a)
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

::  https://github.com/Gamers-Against-Weed/Set-WindowsCbsEdition

:cbsxml:[
param (
    [Parameter()]
    [String]$SetEdition,

    [Parameter()]
    [Switch]$GetTargetEditions,

    [Parameter()]
    [Switch]$StageCurrent
)

function Get-AssemblyIdentity {
    param (
        [String]$PackageName
    )

    $PackageName = [String]$PackageName
    $packageData = ($PackageName -split '~')

    if($packageData[3] -eq '') {
        $packageData[3] = 'neutral'
    }

    return "<assemblyIdentity name=`"$($packageData[0])`" version=`"$($packageData[4])`" processorArchitecture=`"$($packageData[2])`" publicKeyToken=`"$($packageData[1])`" language=`"$($packageData[3])`" />"
}

function Get-SxsName {
    param (
        [String]$PackageName
    )

    $name = ($PackageName -replace '[^A-z0-9\-\._]', '')

    if($name.Length -gt 40) {
        $name = ($name[0..18] -join '') + '\.\.' + ($name[-19..-1] -join '')
    }

    return $name.ToLower()
}

function Find-EditionXmlInSxs {
    param (
        [String]$Edition
    )

    $candidates = @($Edition, 'Client', 'Server')
    $winSxs = $Env:SystemRoot + '\WinSxS'
    $allInSxs = Get-ChildItem -Path $winSxs | select Name

    foreach($candidate in $candidates) {
        $name = Get-SxsName -PackageName "Microsoft-Windows-Editions-$candidate"
        $packages = $allInSxs | where name -Match ('^.*_'+$name+'_31bf3856ad364e35')

        if($packages.Length -eq 0) {
            continue
        }

        $package = $packages[-1].Name
        $testPath = $winSxs + "\$package\" + $Edition + 'Edition.xml'

        if(Test-Path -Path $testPath -PathType Leaf) {
            return $testPath
        }
    }

    return $null
}

function Find-EditionXml {
    param (
        [String]$Edition
    )

    $servicingEditions = $Env:SystemRoot + '\servicing\Editions'
    $editionXml = $Edition + 'Edition.xml'

    $editionXmlInServicing = $servicingEditions + '\' + $editionXml

    if(Test-Path -Path $editionXmlInServicing -PathType Leaf) {
        return $editionXmlInServicing
    }

    return Find-EditionXmlInSxs -Edition $Edition
}

function Write-UpgradeCandidates {
    param (
        [HashTable]$InstallCandidates
    )

    $editionCount = 0
    Write-Host 'Editions that can be upgraded to:'
    foreach($candidate in $InstallCandidates.Keys) {
        Write-Host "Target Edition : $candidate"
        $editionCount++
    }

    if($editionCount -eq 0) {
        Write-Host '(no editions are available)'
    }
}

function Write-UpgradeXml {
    param (
        [Array]$RemovalCandidates,
        [Array]$InstallCandidates,
        [Boolean]$Stage
    )

    $removeAction = 'remove'
    if($Stage) {
        $removeAction = 'stage'
    }

    Write-Output '<?xml version="1.0"?>'
    Write-Output '<unattend xmlns="urn:schemas-microsoft-com:unattend">'
    Write-Output '<servicing>'

    foreach($package in $InstallCandidates) {
        Write-Output '<package action="install">'
        Write-Output (Get-AssemblyIdentity -PackageName $package)
        Write-Output '</package>'
    }

    foreach($package in $RemovalCandidates) {
        Write-Output "<package action=`"$removeAction`">"
        Write-Output (Get-AssemblyIdentity -PackageName $package)
        Write-Output '</package>'
    }

    Write-Output '</servicing>'
    Write-Output '</unattend>'
}

function Write-Usage {
    Get-Help $script:MyInvocation.MyCommand.Path -detailed
}

$version = '1.0'
$getTargetsParam = $GetTargetEditions.IsPresent
$stageCurrentParam = $StageCurrent.IsPresent

if($SetEdition -eq '' -and ($false -eq $getTargetsParam)) {
    Write-Usage
    Exit 1
}

$removalCandidates = @();
$installCandidates = @{};

$packages = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages' | select Name | where { $_.name -match '^.*\\Microsoft-Windows-.*Edition~' }
foreach($package in $packages) {
    $state = (Get-ItemProperty -Path "Registry::$($package.Name)").CurrentState
    $packageName = ($package.Name -split '\\')[-1]
    $packageEdition = (($packageName -split 'Edition~')[0] -split 'Microsoft-Windows-')[-1]

    if($state -eq 0x40) {
        if($null -eq $installCandidates[$packageEdition]) {
            $installCandidates[$packageEdition] = @()
        }

        if($false -eq ($installCandidates[$packageEdition] -contains $packageName)) {
            $installCandidates[$packageEdition] = $installCandidates[$packageEdition] + @($packageName)
        }
    }

    if((($state -eq 0x50) -or ($state -eq 0x70)) -and ($false -eq ($removalCandidates -contains $packageName))) {
        $removalCandidates = $removalCandidates + @($packageName)
    }
}

if($getTargetsParam) {
    Write-UpgradeCandidates -InstallCandidates $installCandidates
    Exit
}

if($false -eq ($installCandidates.Keys -contains $SetEdition)) {
    Write-Error "The system cannot be upgraded to `"$SetEdition`""
    Exit 1
}

$xmlPath = $Env:SystemRoot + '\Temp' + '\CbsUpgrade.xml'

Write-UpgradeXml -RemovalCandidates $removalCandidates `
    -InstallCandidates $installCandidates[$SetEdition] `
    -Stage $stageCurrentParam >$xmlPath

$editionXml = Find-EditionXml -Edition $SetEdition
if($null -eq $editionXml) {
    Write-Warning 'Unable to find edition specific settings XML. Proceeding without it...'
}

Write-Host 'Starting the upgrade process. This may take a while...'

DISM.EXE /English /NoRestart /Online /Apply-Unattend:$xmlPath
$dismError = $LASTEXITCODE

Remove-Item -Path $xmlPath -Force

if(($dismError -ne 0) -and ($dismError -ne 3010)) {
    Write-Error 'Failed to upgrade to the target edition'
    Exit $dismError
}

if($null -ne $editionXml) {
    $destination = $Env:SystemRoot + '\' + $SetEdition + '.xml'
    Copy-Item -Path $editionXml -Destination $destination

    DISM.EXE /English /NoRestart /Online /Apply-Unattend:$editionXml
    $dismError = $LASTEXITCODE

    if(($dismError -ne 0) -and ($dismError -ne 3010)) {
        Write-Error 'Failed to apply edition specific settings'
        Exit $dismError
    }
}

Restart-Computer
:cbsxml:]

::========================================================================================================================================

::  Change edition using DISM API
::  Thanks to Alex (aka may, ave9858)

:dismapi:[
param (
    [Parameter()]
    [String]$TargetEdition,

    [Parameter()]
    [String]$Key
)

$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TB = $ModuleBuilder.DefineType(0)

[void]$TB.DefinePInvokeMethod('DismInitialize', 'DismApi.dll', 22, 1, [int], @([int], [IntPtr], [IntPtr]), 1, 3)
[void]$TB.DefinePInvokeMethod('DismOpenSession', 'DismApi.dll', 22, 1, [int], @([String], [IntPtr], [IntPtr], [UInt32].MakeByRefType()), 1, 3)
[void]$TB.DefinePInvokeMethod('_DismSetEdition', 'DismApi.dll', 22, 1, [int], @([UInt32], [String], [String], [IntPtr], [IntPtr], [IntPtr]), 1, 3)
$Dism = $TB.CreateType()

[void]$Dism::DismInitialize(2, 0, 0)
$Session = 0
[void]$Dism::DismOpenSession('DISM_{53BFAE52-B167-4E2F-A258-0A37B57FF845}', 0, 0, [ref]$Session)
if (!$Dism::_DismSetEdition($Session, "$TargetEdition", "$Key", 0, 0, 0)) {
    Restart-Computer
}
:dismapi:]

::========================================================================================================================================

::  1st column = Generic Retail/OEM/MAK/GVLK Key
::  2nd column = Key Type
::  3rd column = WMI Edition ID
::  4th column = Version name incase same Edition ID is used in different OS versions with different key
::  Separator  = _

::  For Windows 10/11 editions, HWID key is listed where ever possible, in Server versions, KMS key is listed where ever possible.
::  Only RS3 and older version Generic keys are stored here, later ones are extracted from the pkeyhelper.dll itself

:changeeditiondata

if %winbuild% GTR 17763 exit /b
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" (set Cor=Cor) else (set Cor=)

set h=
for %%# in (
XGV%h%PP-NM%h%H47-7TT%h%HJ-W%h%3FW7-8HV%h%2C__OEM:NONSLP_Enterprise
D6R%h%D9-D4%h%N8T-RT9%h%QX-Y%h%W6YT-FCW%h%WJ______Retail_Starter
3V6%h%Q6-NQ%h%XCX-V8Y%h%XR-9%h%QCYV-QPF%h%CT__Volume:MAK_EnterpriseN
3NF%h%XW-2T%h%27M-2BD%h%W6-4%h%GHRV-68X%h%RX______Retail_StarterN
VK7%h%JG-NP%h%HTM-C97%h%JM-9%h%MPGT-3V6%h%6T______Retail_Professional
2B8%h%7N-8K%h%FHP-DKV%h%6R-Y%h%2C8J-PKC%h%KT______Retail_ProfessionalN
4CP%h%RK-NM%h%3K3-X6X%h%XQ-R%h%XX86-WXC%h%HW______Retail_CoreN
N24%h%34-X9%h%D7W-8PF%h%6X-8%h%DV9T-8TY%h%MD______Retail_CoreCountrySpecific
BT7%h%9Q-G7%h%N6G-PGB%h%YW-4%h%YWX6-6F4%h%BT______Retail_CoreSingleLanguage
YTM%h%G3-N6%h%DKC-DKB%h%77-7%h%M9GH-8HV%h%X7______Retail_Core
XKC%h%NC-J2%h%6Q9-KFH%h%D2-F%h%KTHY-KD7%h%2Y__OEM:NONSLP_PPIPro
YNM%h%GQ-8R%h%YV3-4PG%h%Q3-C%h%8XTP-7CF%h%BY______Retail_Education
84N%h%GF-MH%h%BT6-FXB%h%X8-Q%h%WJK7-DRR%h%8H______Retail_EducationN
NK9%h%6Y-D9%h%CD8-W44%h%CQ-R%h%8YTK-DYJ%h%WX__OEM:NONSLP_EnterpriseS_RS1
FWN%h%7H-PF%h%93Q-4GG%h%P8-M%h%8RF3-MDW%h%WW__OEM:NONSLP_EnterpriseS_TH
2DB%h%W3-N2%h%PJG-MVH%h%W3-G%h%7TDK-9HK%h%R4__Volume:MAK_EnterpriseSN_RS1
NTX%h%6B-BR%h%YC2-K67%h%86-F%h%6MVQ-M7V%h%2X__Volume:MAK_EnterpriseSN_TH
G3K%h%NM-CH%h%G6T-R36%h%X3-9%h%QDG6-8M8%h%K9______Retail_ProfessionalSingleLanguage
HNG%h%CC-Y3%h%8KG-QVK%h%8D-W%h%MWRK-X86%h%VK______Retail_ProfessionalCountrySpecific
DXG%h%7C-N3%h%6C4-C4H%h%TG-X%h%4T3X-2YV%h%77______Retail_ProfessionalWorkstation
WYP%h%NQ-8C%h%467-V2W%h%6J-T%h%X4WX-WT2%h%RQ______Retail_ProfessionalWorkstationN
8PT%h%T6-RN%h%W4C-6V7%h%J2-C%h%2D3X-MHB%h%PB______Retail_ProfessionalEducation
GJT%h%YN-HD%h%MQY-FRR%h%76-H%h%VGC7-QPF%h%8P______Retail_ProfessionalEducationN
C4N%h%TJ-CX%h%6Q2-VXD%h%MR-X%h%VKGM-F9D%h%JC__Volume:MAK_EnterpriseG
46P%h%N6-R9%h%BK9-CVH%h%KB-H%h%WQ9V-MBJ%h%Y8__Volume:MAK_EnterpriseGN
NJC%h%F7-PW%h%8QT-332%h%4D-6%h%88JX-2YV%h%66______Retail_ServerRdsh
V3W%h%VW-N2%h%PV2-CGW%h%C3-3%h%4QGF-VMJ%h%2C______Retail_Cloud
NH9%h%J3-68%h%WK7-6FB%h%93-4%h%K3DF-DJ4%h%F6______Retail_CloudN
2HN%h%6V-HG%h%TM8-6C9%h%7C-R%h%K67V-JQP%h%FD______Retail_CloudE
WC2%h%BQ-8N%h%RM3-FDD%h%YY-2%h%BFGV-KHK%h%QY_Volume:GVLK_ServerStandard%Cor%_RS1
CB7%h%KF-BW%h%N84-R7R%h%2Y-7%h%93K2-8XD%h%DG_Volume:GVLK_ServerDatacenter%Cor%_RS1
JCK%h%RF-N3%h%7P4-C2D%h%82-9%h%YXRT-4M6%h%3B_Volume:GVLK_ServerSolution_RS1
QN4%h%C6-GB%h%JD2-FB4%h%22-G%h%HWJK-GJG%h%2R_Volume:GVLK_ServerCloudStorage_RS1
VP3%h%4G-4N%h%PPG-79J%h%TQ-8%h%64T4-R3M%h%QX_Volume:GVLK_ServerAzureCor_RS1
9JQ%h%NQ-V8%h%HQ6-PKB%h%8H-G%h%GHRY-R62%h%H6______Retail_ServerAzureNano_RS1
VN8%h%D3-PR%h%82H-DB6%h%BJ-J%h%9P4M-92F%h%6J______Retail_ServerStorageStandard_RS1
48T%h%QX-NV%h%K3R-D8Q%h%R3-G%h%THHM-8FH%h%XC______Retail_ServerStorageWorkgroup_RS1
2HX%h%DN-KR%h%XHB-GPY%h%C7-Y%h%CKFJ-7FV%h%DG_Volume:GVLK_ServerDatacenterACor_RS3
PTX%h%N8-JF%h%HJM-4WC%h%78-M%h%PCBR-9W4%h%KR_Volume:GVLK_ServerStandardACor_RS3
) do (
for /f "tokens=1-4 delims=_" %%A in ("%%#") do if /i %targetedition%==%%C (

if not defined key (
set 4th=%%D
if not defined 4th (
set "key=%%A" & set "_chan=%%B"
) else (
echo "%branch%" | find /i "%%D" %nul1% && (set "key=%%A" & set "_chan=%%B")
)
)
)
)
exit /b

::========================================================================================================================================
:: Leave empty line below
