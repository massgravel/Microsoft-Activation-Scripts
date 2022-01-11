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
title  Activation Troubleshoot

set _elev=
if /i "%~1"=="-el" set _elev=1

set winbuild=1
set "nul=>nul 2>&1"
set "_psc=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 2>nul | find /i "0x0" 1>nul && (set _NCS=0)

call :_colorprep

set slp=SoftwareLicensingProduct
set sls=SoftwareLicensingService
set ospp=OfficeSoftwareProtectionProduct
set wApp=55c92734-d682-4d71-983e-d6ec3f16059f
set cbs_log=%SystemRoot%\logs\cbs\cbs.log
set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :_color %Red% "==== ERROR ====" &echo:"
set "line=_________________________________________________________________________________________________"

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected.
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
goto at_done
)

if not exist %_psc% (
%nceline%
echo Powershell is not installed in the system.
echo Aborting...
goto at_done
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
%nceline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto at_done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul% reg query HKU\S-1-5-19 || (
if not defined _elev %nul% %_psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%nceline%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto at_done
)

::========================================================================================================================================

setlocal DisableDelayedExpansion

::  Check desktop location

set desktop=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "desktop=%%b"
if not defined desktop for /f "delims=" %%a in ('%_psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"

if not defined desktop (
%eline%
echo Desktop location was not detected, aborting...
goto at_done
)

setlocal EnableDelayedExpansion

::========================================================================================================================================

:at_menu

cls
color 07
title  Activation Troubleshoot
mode con cols=77 lines=30

echo:
echo:
echo:
echo:       _______________________________________________________________
echo:                                                   
echo:             [1] Troubleshoot ReadMe - HWID          
echo:             [2] Troubleshoot ReadMe - KMS38         
echo:             [3] Troubleshoot ReadMe - Online KMS    
echo:             ___________________________________________________      
echo:                                                                      
echo:             [4] Dism RestoreHealth                                   
echo:             [5] SFC Scannow                                          
echo:                                                                      
echo:             [6] Windows Rearm - Specific SKU ID
echo:             [7] Office  Rearm - Specific KMS SKU ID
echo:
echo:             [8] Clean ClipSVC Licences                               
echo:             [9] Rebuild Licensing Tokens                             
echo:                                                                      
echo:             [F] Office License Is Not Genuine - Notification
echo:
echo:             [0] Exit                                                 
echo:       _______________________________________________________________
echo:          
call :_color2 %_White% "            " %_Green% "Enter a menu option in the Keyboard :"
choice /C:123456789F0 /N
set _erl=%errorlevel%

if %_erl%==11 exit /b
if %_erl%==10 start https://windowsaddict.ml/office-license-is-not-genuine &goto at_menu
if %_erl%==9 goto:retokens
if %_erl%==8 goto:cleanclipsvc
if %_erl%==7 goto:officerearm
if %_erl%==6 goto:rearmwin
if %_erl%==5 goto:sfcscan
if %_erl%==4 goto:dism_rest
if %_erl%==3 start https://windowsaddict.ml/readme-troubleshoot-onlinekms.html &goto at_menu
if %_erl%==2 start https://windowsaddict.ml/readme-troubleshoot-kms38.html &goto at_menu
if %_erl%==1 start https://windowsaddict.ml/readme-troubleshoot-hwid.html &goto at_menu
goto :at_menu

::========================================================================================================================================

:dism_rest

cls
mode 98, 30
title  Dism /Online /Cleanup-Image /RestoreHealth

if %winbuild% LSS 9200 (
%eline%
echo Unsupported OS version Detected.
echo This command is supported only for Windows 8/8.1/10/11 and their Server equivalent.
goto :at_back
)

set _int=
for %%a in (dns.msftncsi.com,www.microsoft.com,one.one.one.one,resolver1.opendns.com) do (
if not defined _int (for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do if not [%%#]==[] set _int=1))

echo:
if defined _int (
echo      Checking Internet Connection  [Connected]
) else (
call :_color2 %_White% "     " %Red% "Checking Internet Connection  [Not connected]"
)

echo %line%
echo:
echo      Dism uses Windows Update to provide the files required to fix corruption.
echo      This will take 5-15 minutes or more..
echo %line%
echo:
echo      Notes:
echo:
call :_color2 %_White% "     - " %Gray% "Make sure the Internet is connected."
call :_color2 %_White% "     - " %Gray% "Make sure the Windows update is properly working."
echo:
echo %line%
echo:
choice /C:29 /N /M ">    [9] Continue [2] Go back : "
if %errorlevel%==1 goto at_menu

cls
mode 110, 30
echo:

call :_stopservice TrustedInstaller
del /s /f /q "%SystemRoot%\logs\cbs\*.*"

set _time=
for /f %%a in ('%_psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
echo:
echo Applying the command,
echo dism /online /cleanup-image /restorehealth /Logpath:"%SystemRoot%\Temp\RHealth_DISM_%_time%.txt" /loglevel:4
echo:
dism /online /cleanup-image /restorehealth /Logpath:"%SystemRoot%\Temp\RHealth_DISM_%_time%.txt" /loglevel:4

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%
copy /y /b "%SystemRoot%\Temp\RHealth_DISM_%_time%.txt" "!desktop!\AT_Logs\RHealth_DISM_%_time%.txt" %nul%
copy /y /b "%cbs_log%" "!desktop!\AT_Logs\RHealth_CBS_%_time%.txt" %nul%
del /f /q "%SystemRoot%\Temp\RHealth_DISM_%_time%.txt" %nul%

echo:
call :_color %Gray% "CBS and DISM logs are copied to the AT_Logs folder on the dekstop."
goto :at_back

::========================================================================================================================================

:sfcscan

cls
mode 98, 30
title  sfc /scannow

echo:
echo %line%
echo:    
echo      System File Checker will repair missing or corrupted system files.
echo      This will take 10-15 minutes or more..
echo:
echo      If SFC could not fix something, then run the command again to see if it may be able 
echo      to the next time. Sometimes it may take running the sfc /scannow command 3 times
echo      restarting the PC after each time to completely fix everything that it's able to.
echo:   
echo %line%
echo:
choice /C:29 /N /M ">    [9] Continue [2] Go back : "
if %errorlevel%==1 goto at_menu

cls
echo:

call :_stopservice TrustedInstaller
del /s /f /q "%SystemRoot%\logs\cbs\*.*"

set _time=
for /f %%a in ('%_psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
echo:
echo Applying the command,
echo sfc /scannow
echo:
sfc /scannow

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%

copy /y /b "%cbs_log%" "!desktop!\AT_Logs\SFC_CBS_%_time%.txt" %nul%
findstr /i /c:"[SR]" %cbs_log% | findstr /i /v /c:verify >"!desktop!\AT_Logs\SFC_Main_%_time%.txt"

echo:
call :_color %Gray% "CBS and main extracted logs are copied to the AT_Logs folder on the dekstop."
goto :at_back

::========================================================================================================================================

:rearmwin

cls
mode 98, 30
title  Windows Rearm - Specific SKU ID

if %winbuild% LSS 9600 (
%eline%
echo Unsupported OS version Detected.
echo This command is supported only for Windows 8/8.1/10/11 and their Server equivalent.
goto :at_back
)

echo:
echo %line%
echo:    
echo      Notes:
echo:
echo       - Rearm helps in troubleshooting activation issues.
echo:
echo       - Windows KMS activation will reset.
echo:
echo %line%
echo:
choice /C:29 /N /M ">    [9] Continue [2] Go back : "
if %errorlevel%==1 goto at_menu

cls
mode 105, 30
echo:
echo:
set app=
net start sppsvc /y %nul%
if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%wApp%' and PartialProductKey is not NULL) get ID /VALUE" 2^>nul') do call set "app=%%a"
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%a in ('%_psc% "(([WMISEARCHER]'SELECT ID FROM %slp% WHERE ApplicationID=''%wApp%'' AND PartialProductKey IS NOT NULL').Get()).ID | %% {echo ('ID='+$_)}" 2^>nul') do call set "app=%%a"

if defined app (
if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%x in ('"wmic path %slp% where ID='%app%' get Name /VALUE" 2^>nul') do echo Rearming: %%x
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%x in ('%_psc% "(([WMISEARCHER]'SELECT Name FROM %slp% WHERE ID=''%app%''').Get()).Name | %% {echo ('Name='+$_)}" 2^>nul') do echo Rearming: %%x
echo:
echo Applying the command,
if %winbuild% LSS 22483 (
echo wmic path %slp% where ID='%app%' call ReArmsku
wmic path %slp% where ID='%app%' call ReArmsku %nul%
)
if %winbuild% GEQ 22483 (
echo Powershell "$null=([WMI]'%slp%=''%app%''').ReArmsku()"
%_psc% "$null=([WMI]'%slp%=''%app%''').ReArmsku()" %nul%
)
if !errorlevel!==0 (
call :_color %Green% "[Successful]"
) else (
call :_color %Red% "[Failed]"
)
) else (
call :_color %Red% "Error- Activation ID not found"
)

goto :at_back

::========================================================================================================================================

:officerearm

cls
mode 98, 30
title  Office Rearm - Specific KMS SKU ID

if %winbuild% LSS 9600 (
%eline%
echo Unsupported OS version Detected.
echo This command is supported only for Windows 8/8.1/10/11 and their Server equivalent.
goto :at_back
)

echo:
echo %line%
echo:   
echo      Notes:
echo:
echo       - Rearm helps in troubleshooting activation issues.
echo:
echo       - Office KMS activation will reset.
echo:
call :_color2 %_White% "      - " %Gray% "Office rearm can be applied only a certain number of times."
echo:
echo %line%
echo:
choice /C:29 /N /M ">    [9] Continue [2] Go back : "
if %errorlevel%==1 goto at_menu

cls
mode 105, 30
echo:

net start sppsvc /y %nul%
call :getapplist %slp%

if defined applist (
for %%# in (%applist%) do (
echo:
if %winbuild% LSS 22483 for /f "tokens=2 delims==" %%x in ('"wmic path %slp% where ID='%%#' get Name /VALUE" 2^>nul') do echo Rearming: %%x
if %winbuild% GEQ 22483 for /f "tokens=2 delims==" %%x in ('%_psc% "(([WMISEARCHER]'SELECT Name FROM %slp% WHERE ID=''%%#''').Get()).Name | %% {echo ('Name='+$_)}" 2^>nul') do echo Rearming: %%x
echo:
echo Applying the command,
if %winbuild% LSS 22483 (
echo wmic path %slp% where ID='%%#' call ReArmsku
wmic path %slp% where ID='%%#' call ReArmsku %nul%
)
if %winbuild% GEQ 22483 (
echo Powershell "$null=([WMI]'%slp%=''%%#''').ReArmsku()"
%_psc% "$null=([WMI]'%slp%=''%%#''').ReArmsku()" %nul%
)
if !errorlevel!==0 (
call :_color %Green% "[Successful]"
) else (
call :_color %Red% "[Failed]"
)
)
) else (
echo:
echo Checking: Volume version of Office 2013-2021 is not found.
)

call :getapplist %ospp%

if defined applist (
if %winbuild% LSS 9200 (set _off=Office) else (set _off=Office 2010)
echo:
echo Skipping the Rearm of OSPP based '!_off!'
)

goto :at_back

:getapplist

set applist=
if %winbuild% LSS 22483 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %1 where (Name like '%%office%%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get ID /VALUE" 2^>nul')"
if %winbuild% GEQ 22483 set "chkapp=for /f "tokens=2 delims==" %%a in ('%_psc% "(([WMISEARCHER]'SELECT ID FROM %1 WHERE Name like ''%%office%%'' and Description like ''%%KMSCLIENT%%'' and PartialProductKey is not NULL').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))
exit /b

::========================================================================================================================================

:retokens

cls
mode 98, 30
title  Rebuild Licensing Tokens ^& Re-install System License Files

echo:
echo %line%
echo:   
echo      Notes:
echo:
echo       - Rebuild Licensing Tokens ^& Re-install System License Files
echo         It helps in troubleshooting activation issues.
echo:
call :_color2 %_White% "      - " %Gray% "Windows and Office activation may reset, you may need to activate them again."
echo:
call :_color2 %_White% "      - " %Magenta% "This option will uninstall Office licenses and keys."
call :_color2 %_White% "        " %Magenta% "Installed Office will need to repair itself ones upon opening an office app,"
call :_color2 %_White% "        " %Magenta% "you may also need to repair Office from Apps and Features in Windows Settings."
echo:
call :_color2 %_White% "      - " %Gray% "Script is designed to skip rebuilding tokens where products may not be able to"
call :_color2 %_White% "        " %Gray% "restore their license."
echo:
call :_color2 %_White% "      - " %Red% "Apply it only when it is necessary."
echo:
echo %line%
echo:
choice /C:24 /N /M ">    [4] Continue [2] Go back : "
if %errorlevel%==1 goto at_menu

cls
mode 98, 30

set nosup=
set 68=HKLM\SOFTWARE\Microsoft\Office
set 86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

%nul% reg query %68%\16.0\Common\InstallRoot /v Path                && set nosup=1 REM Office 2016 MSI x86-x64
%nul% reg query %86%\16.0\Common\InstallRoot /v Path                && set nosup=1 REM Office 2016 MSI x86
%nul% reg query %68%\15.0\Common\InstallRoot /v Path                && set nosup=1 REM Office 2013 MSI x86-x64
%nul% reg query %86%\15.0\Common\InstallRoot /v Path                && set nosup=1 REM Office 2013 MSI x86
%nul% reg query %68%\14.0\Common\InstallRoot /v Path                && set nosup=1 REM Office 2010 MSI x86-x64
%nul% reg query %86%\14.0\Common\InstallRoot /v Path                && set nosup=1 REM Office 2010 MSI x86
%nul% reg query %68%\14.0\CVH /f Click2run /k                       && set nosup=1 REM Office 2010 C2R x86-x64
%nul% reg query %86%\14.0\CVH /f Click2run /k                       && set nosup=1 REM Office 2010 C2R x86

if %winbuild% GEQ 10240 reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msoxmled.exe" %nul% && (
dir /b "%ProgramFiles%\WindowsApps\Microsoft.Office.Desktop*" %nul% && set nosup=1 REM Office UWP x86-x64
dir /b "%ProgramW6432%\WindowsApps\Microsoft.Office.Desktop*" %nul% && set nosup=1 REM Office UWP x86
)

sc qc osppsvc %nul% && (
if %winbuild% LSS 9200 (set _off=Office) else (set _off=Office 2010)
echo:
echo Skipping rebuilding OSPP tokens for detected '!_off!'
)

echo:
if defined nosup (
echo Detected Office may not be able to repair itself, hence skipping SPP tokens rebuilding...
goto :at_back
)

net start sppsvc /y %nul%

call :at_permcheck Office
if defined _perm (
echo Office is permanently activated, token rebuilding may deactivate it, hence skipping...
goto :at_back
)

if %winbuild% LSS 10240 (
call :at_permcheck Windows
if defined _perm (
echo Windows is permanently activated, token rebuilding may deactivate it, hence skipping...
goto :at_back
)
)

set token=
for %%# in (
%Systemdrive%\Windows\System32\spp\store_test\2.0\tokens.dat
%Systemdrive%\Windows\System32\spp\store\tokens.dat
%Systemdrive%\Windows\System32\spp\store\2.0\tokens.dat
%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat
) do if exist %%# set token=%%#

echo %line%
echo:
call :_color %Gray% "Rebuilding SoftwareProtectionPlatform tokens.dat"
echo %line%
echo:

if not exist "%token%" (
%eline%
echo tokens.dat file not found.
echo Restart the system and try again.
goto :at_back
) else (
echo Detected tokens.dat file [%token%]
)

echo Stopping sppsvc service...
call :_stopservice sppsvc

::  data.dat and cache files are not deleted since doing that may corrupt the office license in a way that only reinstallation can fix

del /f /q %token% %nul%
if exist %token% (
call :_stopservice sppsvc
del /f /q %token% %nul%
)

echo:
if exist %token% (
call :_color %Red% "Failed to delete the tokens.dat file."
) else (
echo tokens.dat file was successfully deleted.
)

echo:
echo Reinstalling System Licenses [slmgr /rilc]...
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul% && (
echo [Successful]
) || (
call :_color %Red% "[Failed]"
)

echo:
if exist %token% (
call :_color %Green% "tokens.dat file was rebuilt successfully."
) else (
call :_color %Red% "Failed to rebuilt tokens.dat file."
)

goto :at_back

:at_permcheck

set _perm=
if %winbuild% LSS 22483 wmic path %slp% where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name /value 2>nul | findstr /i "%1" 1>nul && set _perm=1||set _perm=
if %winbuild% GEQ 22483 %_psc% "(([WMISEARCHER]'SELECT Name FROM %slp% WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL').Get()).Name | %% {echo ('Name='+$_)}" 2>nul | findstr /i "%1" 1>nul && set _perm=1||set _perm=
exit /b

::========================================================================================================================================

:cleanclipsvc

cls
mode 98, 30
title  Clean ClipSVC Licences

if %winbuild% LSS 10240 (
%eline%
echo Unsupported OS version Detected.
echo This command is supported only for Windows 10/11 and their Server equivalent..
goto :at_back
)

echo:
echo %line%
echo:
echo      Notes:
echo:
echo       - Cleaning ClipSVC Licences helps in troubleshooting HWID-KMS38 activation issues.
echo:
call :_color2 %_White% "      - " %Gray% "All installed HWID-KMS38 licences will be cleaned from the system."
echo         This will remove KMS38 license from the system but HWID license can't be removed.
echo:
echo       - System restart is recommended after applying it.
echo:
echo %line%
echo:
choice /C:29 /N /M ">    [9] Continue [2] Go back : "
if %errorlevel%==1 goto at_menu

cls
echo:

echo Stopping ClipSVC service...
call :_stopservice ClipSVC
timeout /t 2 %nul%

::  Thanks to @mspaintmsi for informing this command info

echo:
echo Applying the command to Clean ClipSVC Licences...
echo rundll32 clipc.dll,ClipCleanUpState

rundll32 clipc.dll,ClipCleanUpState

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\*.dat" del /f /q "%ProgramData%\Microsoft\Windows\ClipSVC\*.dat" %nul%

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat" (
call :_color %Red% "[Failed]"
) else (
call :_color %Green% "[Successful]"
)

::  Below registry key (Volatile & Protected) gets created after the ClipSVC License cleanup command, and gets automatically deleted after 
::  system restart. It needs to be deleted to activate the system without restart.

set "RegKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState"

call :regown "%RegKey%" %nul%
reg delete "%RegKey%" /f %nul%

echo:
echo Deleting a Volatile ^& Protected Registry Key...
echo [%RegKey%]
reg query "%RegKey%" %nul% && (
call :_color %Red% "[Failed]"
echo Restart the system, that will delete this registry key automatically.
) || (
call :_color %Green% "[Successful]"
)

echo:
echo Restarting [ClipSVC wlidsvc LicenseManager sppsvc] services...
for %%# in (ClipSVC wlidsvc LicenseManager sppsvc) do (net stop %%# /y %nul% & net start %%# /y %nul%)

::========================================================================================================================================

:at_back

echo:
echo %line%
echo:
call :_color %_Yellow% "Press any key to go back..."
pause >nul
goto :at_menu

::========================================================================================================================================

:at_done

echo:
echo Press any key to exit...
pause >nul
exit /b

::========================================================================================================================================

:_stopservice

for %%# in (%1) do (
sc query %%# | find /i "STOPPED" %nul% || net stop %%# /y %nul%
sc query %%# | find /i "STOPPED" %nul% || sc stop %%# %nul%
)
exit /b

::========================================================================================================================================\

::  A lean and mean snippet to set registry ownership and permission recursively
::  Written by @AveYo aka @BAU
::  pastebin.com/XTPt0JSC

::  Modified by @abbodi1406 to make it work in ARM64 Windows 10 (builds older than 21277) where only x86 version of PowerShell is installed.

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

:_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
call :batcol %~1 "%~2"
)
exit /b

:_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
call :batcol %~1 "%~2" %~3 "%~4"
)
exit /b

::=======================================

:: Colored text with pure batch method
:: Thanks to @dbenham and @jeb
:: stackoverflow.com/a/10407642

:batcol

pushd %_coltemp%
if not exist "'" (<nul >"'" set /p "=.")
setlocal
set "s=%~2"
set "t=%~4"
call :_batcol %1 s %3 t
del /f /q "'"
del /f /q "`.txt"
popd
exit /b

:_batcol

setlocal EnableDelayedExpansion
set "s=!%~2!"
set "t=!%~4!"
for /f delims^=^ eol^= %%i in ("!s!") do (
  if "!" equ "" setlocal DisableDelayedExpansion
    >`.txt (echo %%i\..\')
    findstr /a:%~1 /f:`.txt "."
    <nul set /p "=%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%"
)
if "%~4"=="" echo(&exit /b
setlocal EnableDelayedExpansion
for /f delims^=^ eol^= %%i in ("!t!") do (
  if "!" equ "" setlocal DisableDelayedExpansion
    >`.txt (echo %%i\..\')
    findstr /a:%~3 /f:`.txt "."
    <nul set /p "=%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%"
)
echo(
exit /b

::=======================================

:_colorprep

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"

set     "Red="41;97m""
set    "Gray="100;97m""
set   "Black="30m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "Yellow="43;97m""
set "Magenta="45;97m""

set    "_Red="40;91m""
set  "_Green="40;92m""
set   "_Blue="40;94m""
set  "_White="40;37m""
set "_Yellow="40;93m""

exit /b
)

for /f %%A in ('"prompt $H&for %%B in (1) do rem"') do set "_BS=%%A %%A"
set "_coltemp=%SystemRoot%\Temp"

set     "Red="CF""
set    "Gray="8F""
set   "Black="00""
set   "Green="2F""
set    "Blue="1F""
set  "Yellow="6F""
set "Magenta="5F""

set    "_Red="0C""
set  "_Green="0A""
set   "_Blue="09""
set  "_White="07""
set "_Yellow="0E""

exit /b

::========================================================================================================================================