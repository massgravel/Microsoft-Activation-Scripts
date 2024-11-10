@set masver=2.8
@echo off



::============================================================================
::
::   Homepage: mass grave[.]dev
::      Email: mas.help@outlook.com
::
::============================================================================



::========================================================================================================================================

::  Set environment variables, it helps if they are misconfigured in the system

setlocal EnableExtensions
setlocal DisableDelayedExpansion

set "PathExt=.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC"

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)

set "ComSpec=%SysPath%\cmd.exe"
set "PSModulePath=%ProgramFiles%\WindowsPowerShell\Modules;%SysPath%\WindowsPowerShell\v1.0\Modules"

set re1=
set re2=
set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="re1" set re1=1
if /i "%%#"=="re2" set re2=1
)

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

if exist %SystemRoot%\Sysnative\cmd.exe if not defined re1 (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* re1"
exit /b
)

:: Re-launch the script with ARM32 process if it was initiated by x64 process on ARM64 Windows

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 if not defined re2 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* re2"
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
echo Help - %mas%troubleshoot
echo:
echo:
ping 127.0.0.1 -n 20
)
cls

::  Check LF line ending

pushd "%~dp0"
>nul findstr /v "$" "%~nx0" && (
echo:
echo Error - Script either has LF line ending issue or an empty line at the end of the script is missing.
echo:
echo:
echo Help - %mas%troubleshoot
echo:
echo:
ping 127.0.0.1 -n 20 >nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title  Troubleshoot %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args set _args=%_args:re1=%
if defined _args set _args=%_args:re2=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="-el"                    set _elev=1
)
)

set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"

call :dk_setvar
set "line=_________________________________________________________________________________________________"

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected [%winbuild%].
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalents.
goto dk_done
)

::========================================================================================================================================

::  Fix special character limitations in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%
set _PSarg=%_PSarg:'=''%

set "_ttemp=%userprofile%\AppData\Local\Temp"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
if /i not "!_work!"=="!_ttemp!" (
%eline%
echo The script was launched from the temp folder.
echo You are most likely running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto dk_done
)
)

::========================================================================================================================================

::  Check PowerShell

REM :PowerShellTest: $ExecutionContext.SessionState.LanguageMode :PowerShellTest:

cmd /c "%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':PowerShellTest:\s*';iex ($f[1])"" | find /i "FullLanguage" %nul1% || (
%eline%
cmd /c "%psc% "$ExecutionContext.SessionState.LanguageMode""
echo:
cmd /c "%psc% "$ExecutionContext.SessionState.LanguageMode"" | find /i "FullLanguage" %nul1% && (
echo Failed to run Powershell command but Powershell is working.
echo:
cmd /c "%psc% ""$av = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct; $n = @(); foreach ($i in $av) { if ($i.displayName -notlike '*windows*') { $n += $i.displayName } }; if ($n) { Write-Host ('Installed 3rd party Antivirus might be blocking the script - ' + ($n -join ', ')) -ForegroundColor White -BackgroundColor Blue }"""
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
) || (
echo PowerShell is not working. Aborting...
echo If you have applied restrictions on Powershell then undo those changes.
echo:
set fixes=%fixes% %mas%fix_powershell
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%fix_powershell"
)
goto dk_done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo Right click on this script and select 'Run as administrator'.
goto dk_done
)

::========================================================================================================================================

::  Disable QuickEdit and launch from conhost.exe to avoid Terminal app

if %winbuild% GEQ 17763 (
set terminal=1
) else (
set terminal=
)

::  Check if script is running in Terminal app

set r1=$TB = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0);
set r2=%r1% [void]$TB.DefinePInvokeMethod('GetConsoleWindow', 'kernel32.dll', 22, 1, [IntPtr], @(), 1, 3).SetImplementationFlags(128);
set r3=%r2% [void]$TB.DefinePInvokeMethod('SendMessageW', 'user32.dll', 22, 1, [IntPtr], @([IntPtr], [UInt32], [IntPtr], [IntPtr]), 1, 3).SetImplementationFlags(128);
set d1=%r3% $hIcon = $TB.CreateType(); $hWnd = $hIcon::GetConsoleWindow();
set d2=%d1% echo $($hIcon::SendMessageW($hWnd, 127, 0, 0) -ne [IntPtr]::Zero);

if defined terminal (
%psc% "%d2%" %nul2% | find /i "True" %nul1% && set terminal=
)

if defined ps32onArm goto :skipQE
if %_unattended%==1 goto :skipQE
for %%# in (%_args%) do (if /i "%%#"=="-qedit" goto :skipQE)

if defined terminal (
set "launchcmd=start conhost.exe %psc%"
) else (
set "launchcmd=%psc%"
)

::  Disable QuickEdit in current session

set "d1=$t=[AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0);"
set "d2=$t.DefinePInvokeMethod('GetStdHandle', 'kernel32.dll', 22, 1, [IntPtr], @([Int32]), 1, 3).SetImplementationFlags(128);"
set "d3=$t.DefinePInvokeMethod('SetConsoleMode', 'kernel32.dll', 22, 1, [Boolean], @([IntPtr], [Int32]), 1, 3).SetImplementationFlags(128);"
set "d4=$k=$t.CreateType(); $b=$k::SetConsoleMode($k::GetStdHandle(-10), 0x0080);"

%launchcmd% "%d1% %d2% %d3% %d4% & cmd.exe '/c' '!_PSarg! -qedit'" && (exit /b) || (set terminal=1)
:skipQE

::========================================================================================================================================

::  Check for updates

set -=
set old=

for /f "delims=[] tokens=2" %%# in ('ping -4 -n 1 updatecheck.mass%-%grave.dev') do (
if not "%%#"=="" (echo "%%#" | find "127.69" %nul1% && (echo "%%#" | find "127.69.%masver%" %nul1% || set old=1))
)

if defined old (
echo ________________________________________________
%eline%
echo Your version of MAS [%masver%] is outdated.
echo ________________________________________________
echo:
if not %_unattended%==1 (
echo [1] Get Latest MAS
echo [0] Continue Anyway
echo:
call :dk_color %_Green% "Choose a menu option using your keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
)
)

::========================================================================================================================================

setlocal DisableDelayedExpansion

::  Check desktop location

set desktop=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "desktop=%%b"
if not defined desktop for /f "delims=" %%a in ('%psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"

if not defined desktop (
%eline%
echo Unable to detect Desktop location, aborting...
goto dk_done
)

setlocal EnableDelayedExpansion

::========================================================================================================================================

:at_menu

cls
color 07
title  Troubleshoot %masver%
if not defined terminal mode 77, 30

echo:
echo:
echo:
echo:
echo:       _______________________________________________________________
echo:                                                   
call :dk_color2 %_White% "             [1] " %_Green% "Help"
echo:             ___________________________________________________
echo:                                                                      
echo:             [2] Dism RestoreHealth
echo:             [3] SFC Scannow
echo:                                                                      
echo:             [4] Fix WMI
echo:             [5] Fix Licensing
echo:             [6] Fix WPA Registry
echo:             ___________________________________________________
echo:
echo:             [0] %_exitmsg%
echo:       _______________________________________________________________
echo:          
call :dk_color2 %_White% "            " %_Green% "Choose a menu option using your keyboard :"
choice /C:1234560 /N
set _erl=%errorlevel%

if %_erl%==7 exit /b
if %_erl%==6 start %mas%fix-wpa-registry &goto at_menu
if %_erl%==5 goto:retokens
if %_erl%==4 goto:fixwmi
if %_erl%==3 goto:sfcscan
if %_erl%==2 goto:dism_rest
if %_erl%==1 start %mas%troubleshoot.html &goto at_menu
goto :at_menu

::========================================================================================================================================

:dism_rest

cls
if not defined terminal mode 98, 30
title  Dism /English /Online /Cleanup-Image /RestoreHealth

if %winbuild% LSS 9200 (
%eline%
echo Unsupported OS version detected.
echo This command only works on Windows 8/8.1/10/11 and their Server equivalents.
goto :at_back
)

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not "%%#"=="" set _int=1)
)

echo:
if defined _int (
echo      Checking Internet Connection  [Connected]
) else (
call :dk_color2 %_White% "     " %Red% "Checking Internet Connection  [Not connected]"
)

echo %line%
echo:
echo      DISM uses Windows Update to provide replacement files required to fix corruption.
echo      This will take 5-15 minutes or more..
echo %line%
echo:
echo      Notes:
echo:
call :dk_color2 %_White% "     - " %Gray% "Make sure the internet is connected."
call :dk_color2 %_White% "     - " %Gray% "Make sure that Windows update is properly working."
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
if not defined terminal mode 110, 30

for /f %%a in ('%psc% "(Get-Date).ToString('yyyyMMdd-HHmmssfff')"') do set _time=%%a

%psc% Stop-Service TrustedInstaller -force %nul%

copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\backup_cbs_%_time%.log" %nul%
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "%SystemRoot%\logs\DISM\backup_dism_%_time%.log" %nul%
del /f /q "%SystemRoot%\logs\cbs\cbs.log" %nul%
del /f /q "%SystemRoot%\logs\DISM\dism.log" %nul%

echo:
echo Applying the command...
echo dism /english /online /cleanup-image /restorehealth
dism /english /online /cleanup-image /restorehealth

timeout /t 5 %nul1%
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\cbs_%_time%.log" %nul%
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "%SystemRoot%\logs\DISM\dism_%_time%.log" %nul%

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%
call :compresslog cbs\cbs_%_time%.log AT_Logs\RHealth_CBS %nul%
call :compresslog DISM\dism_%_time%.log AT_Logs\RHealth_DISM %nul%

if not exist "!desktop!\AT_Logs\RHealth_CBS_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "!desktop!\AT_Logs\RHealth_CBS_%_time%.log" %nul%
)

if not exist "!desktop!\AT_Logs\RHealth_DISM_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "!desktop!\AT_Logs\RHealth_DISM_%_time%.log" %nul%
)

echo:
call :dk_color %Gray% "CBS and DISM logs are copied to the AT_Logs folder on your desktop."
goto :at_back

::========================================================================================================================================

:sfcscan

cls
if not defined terminal mode 98, 30
title  sfc /scannow

echo:
echo %line%
echo:    
echo      SFC will repair missing or corrupted system files.
echo      It is recommended you run the DISM option first before this one.
echo      This will take 10-15 minutes or more..
echo:
echo      If SFC could not fix something, then run the command again to see if it may be able 
echo      to the next time. Sometimes it may take running the sfc /scannow command 3 times
echo      restarting the PC after each time to completely fix everything that it's able to.
echo:   
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
for /f %%a in ('%psc% "(Get-Date).ToString('yyyyMMdd-HHmmssfff')"') do set _time=%%a

%psc% Stop-Service TrustedInstaller -force %nul%

copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\backup_cbs_%_time%.log" %nul%
del /f /q "%SystemRoot%\logs\cbs\cbs.log" %nul%

echo:
echo Applying the command...
echo sfc /scannow
sfc /scannow

timeout /t 5 %nul1%
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\cbs_%_time%.log" %nul%

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%
call :compresslog cbs\cbs_%_time%.log AT_Logs\SFC_CBS %nul%

if not exist "!desktop!\AT_Logs\SFC_CBS_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "!desktop!\AT_Logs\SFC_CBS_%_time%.log" %nul%
)

echo:
call :dk_color %Gray% "The CBS log was copied to the AT_Logs folder on your Desktop."
goto :at_back

::========================================================================================================================================

:retokens

cls
if not defined terminal (
mode 125, 32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)
title  Fix Licensing ^(ClipSVC ^+ SPP ^+ OSPP^)

echo:
echo %line%
echo:   
echo      Notes:
echo:
echo       - This option helps in troubleshooting activation issues.
echo:
echo       - This option will:
echo            - Deactivate Windows and Office, you may need to reactivate.
echo              If Windows is activated with motherboard / OEM / Digital license
echo              then Windows will activate itself again.
echo:
echo            - Clear ClipSVC, SPP and OSPP licenses.
echo            - Fix permissions of SPP tokens folder and registries.
echo            - Trigger the repair option for Office.
echo:
call :dk_color2 %_White% "      - " %Red% "Apply this option only when it is necessary."
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

::========================================================================================================================================

::  Rebuild ClipSVC Licences

cls
:cleanlicensing

echo:
echo %line%
echo:
call :dk_color %Blue% "Rebuilding ClipSVC Licenses..."
echo:

if %winbuild% LSS 10240 (
echo ClipSVC license rebuilding is supported only on Windows 10/11 and their Server equivalents.
echo Skipping...
goto :rebuildspptok
)

%psc% "(([WMISEARCHER]'SELECT Name FROM SoftwareLicensingProduct WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL AND LicenseDependsOn is NULL').Get()).Name" %nul2% | findstr /i "Windows" %nul1% && (
echo Windows is permanently activated.
echo Skipping...
goto :rebuildspptok
)

echo Stopping ClipSVC service...
%psc% Stop-Service ClipSVC -force %nul%
timeout /t 2 %nul%

echo:
echo Applying the command to clean ClipSVC Licenses...
echo rundll32 clipc.dll,ClipCleanUpState

rundll32 clipc.dll,ClipCleanUpState

if %winbuild% LEQ 10240 (
echo [Successful]
) else (
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat" (
call :dk_color %Red% "[Failed]"
) else (
echo [Successful]
)
)

::  Below registry key (Volatile & Protected) gets created after the ClipSVC License cleanup command, and gets automatically deleted after 
::  system restart. It needs to be deleted to activate the system without restart.

set "RegKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState"
set "_ident=HKU\S-1-5-19\SOFTWARE\Microsoft\IdentityCRL"

reg query "%RegKey%" %nul% && %nul% call :regownstart
reg delete "%RegKey%" /f %nul% 

echo:
echo Deleting a Volatile ^& Protected Registry Key...
echo [%RegKey%]
reg query "%RegKey%" %nul% && (
call :dk_color %Red% "[Failed]"
echo Reboot your machine using the restart option, that will delete this registry key automatically.
) || (
echo [Successful]
)

::   Clear HWID token related registry to fix activation incase there is any corruption

echo:
echo Deleting IdentityCRL Registry Key...
echo [%_ident%]
reg delete "%_ident%" /f %nul%
reg query "%_ident%" %nul% && (
call :dk_color %Red% "[Failed]"
) || (
echo [Successful]
)

%psc% Stop-Service ClipSVC -force %nul%

::  Rebuild ClipSVC folder to fix permission issues

echo:
if %winbuild% GTR 10240 (
echo Deleting folder %ProgramData%\Microsoft\Windows\ClipSVC\
rmdir /s /q "C:\ProgramData\Microsoft\Windows\ClipSvc" %nul%

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\" (
call :dk_color %Red% "[Failed]"
) else (
echo [Successful]
)

echo:
echo Rebuilding the %ProgramData%\Microsoft\Windows\ClipSVC\ folder...
%psc% Start-Service ClipSVC %nul%
timeout /t 3 %nul%
if not exist "%ProgramData%\Microsoft\Windows\ClipSVC\" timeout /t 5 %nul%
if not exist "%ProgramData%\Microsoft\Windows\ClipSVC\" (
call :dk_color %Red% "[Failed]"
) else (
echo [Successful]
)
)

echo:
echo Restarting wlidsvc ^& LicenseManager services...
for %%# in (wlidsvc LicenseManager) do (%psc% "Start-Job { Restart-Service %%# } | Wait-Job -Timeout 20 | Out-Null")

::========================================================================================================================================

::  Rebuild SPP Tokens

:rebuildspptok

echo:
echo %line%
echo:
call :dk_color %Blue% "Rebuilding SPP licensing tokens..."
echo:

call :scandat check

if not defined token (
call :dk_color %Red% "tokens.dat file not found."
) else (
echo tokens.dat file: [%token%]
)

set tokenstore=
set badregistry=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore %nul6%') do call set "tokenstore=%%b"
if %winbuild% GEQ 9200 if /i not "%tokenstore%"=="%SysPath%\spp\store" if /i not "%tokenstore%"=="%SysPath%\spp\store\2.0" if /i not "%tokenstore%"=="%SysPath%\spp\store_test\2.0" (
set badregistry=1
echo:
call :dk_color %Red% "Correct path not found in TokenStore Registry [%tokenstore%]"
)

::  Check sppsvc permissions and apply fixes

if %winbuild% GEQ 9200 if not defined badregistry (
echo:
echo Checking SPP permission related issues...
call :checkperms
if defined permerror (
call :dk_color %Red% "[!permerror!]"
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':fixsppperms\:.*';iex ($f[1])" %nul%
call :checkperms
if defined permerror (
call :dk_color %Red% "[!permerror!] [Failed To Fix]"
) else (
call :dk_color %Green% "[Successfully Fixed]"
)
) else (
echo [No Error Found]
)
)

echo:
echo Stopping sppsvc service...
%psc% Stop-Service sppsvc -force %nul%

set w=
set _sppint=
for %%# in (SppEx%w%tComObj.exe sppsvc.exe) do (reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%w%ge File Execu%w%tion Options\%%#" %nul% && (set _sppint=1))
if defined _sppint (
echo:
echo Removing SPP IFEO registry keys...
for %%# in (SppE%w%xtComObj.exe sppsvc.exe) do (reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%w%ge File Execu%w%tion Options\%%#" /f %nul%)
)

if %winbuild% LSS 9200 (
REM Fix issues caused by Update KB971033 in Windows 7
REM https://support.microsoft.com/help/4487266
echo:
echo Checking Update KB971033...
%psc% "if (Get-Hotfix -Id KB971033 -ErrorAction SilentlyContinue) {Exit 3}" %nul%
if !errorlevel!==3 (
echo Found, uninstalling it...
wusa /uninstall /quiet /norestart /kb:971033
) else (
echo [Not Found]
)
%psc% Stop-Service sppuinotify -force %nul%
sc config sppuinotify start= disabled
del /f /q %SysPath%\7B296FB0-376B-497e-B012-9C450E1B7327-*.C7483456-A289-439d-8115-601632D005A0 /ah
)

::  Delete registry keys that are not deleted by activation scripts

echo:
echo Cleaning some licensing-related registry keys...
%nul% reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "ServiceSessionId" /f
%nul% reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "LicStatusArray" /f
%nul% reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "PolicyValuesArray" /f
%nul% reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "actionlist" /f
%nul% reg delete "HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\data" /f

echo:
call :scandat delete
call :scandat check

if defined token (
echo:
call :dk_color %Red% "Failed to delete .dat files."
echo:
)

echo:
echo Reinstalling system licenses...
%psc% "Stop-Service sppsvc -force; $sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
if %errorlevel% NEQ 0 %psc% "$sls = Get-WmiObject SoftwareLicensingService; $f=[io.file]::ReadAllText('!_batp!') -split ':xrm\:.*';iex ($f[1]); ReinstallLicenses" %nul%
if %errorlevel% EQU 0 (
echo [Successful]
) else (
call :dk_color %Red% "[Failed]"
)

call :scandat check

echo:
if not defined token (
call :dk_color %Red% "Failed to rebuild tokens.dat file."
) else (
echo tokens.dat file was rebuilt successfully.
)

if %winbuild% LSS 9200 (
sc config sppuinotify start= demand
)

::========================================================================================================================================

::  Rebuild OSPP Tokens

echo:
echo %line%
echo:
call :dk_color %Blue% "Rebuilding OSPP licensing tokens..."
echo:

sc qc osppsvc %nul% || (
echo OSPP-based Office is not installed.
echo Skipping rebuilding OSPP tokens...
goto :repairoffice
)

call :scandatospp check

if not defined token (
call :dk_color %Red% "tokens.dat file not found."
) else (
echo tokens.dat file: [%token%]
)

echo:
echo Stopping osppsvc service...
%psc% Stop-Service osppsvc -force %nul%

echo:
call :scandatospp delete
call :scandatospp check

if defined token (
echo:
call :dk_color %Red% "Failed to delete .dat files."
echo:
)

echo:
echo Starting osppsvc service to generate tokens.dat...
%psc% Start-Service osppsvc %nul%
call :scandatospp check
if not defined token (
%psc% Stop-Service osppsvc -force %nul%
%psc% Start-Service osppsvc %nul%
timeout /t 3 %nul%
)

call :scandatospp check

echo:
if not defined token (
call :dk_color %Red% "Failed to rebuild tokens.dat file."
) else (
echo tokens.dat file was rebuilt successfully.
)

::========================================================================================================================================

:repairoffice

echo:
echo %line%
echo:
call :dk_color %Blue% "Repairing Office licenses..."
echo:

for %%# in (68 86) do (
for %%A in (msi14 msi15 msi16 c2r14 c2r15 c2r16) do (set %%A_%%#=&set %%Arepair%%#=)
)

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

reg query %_68%\14.0\CVH /f Click2run /k %nul% && (set "c2r14_68=Office 14.0 C2R x86/x64"  & set "c2r14repair68=")
reg query %_86%\14.0\CVH /f Click2run /k %nul% && (set "c2r14_86=Office 14.0 C2R x86"      & set "c2r14repair86=")

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi14_86=Office 14.0 MSI x86"      & call :getrepairsetup msi14repair86 14)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi14_68=Office 14.0 MSI x86/x64"  & call :getrepairsetup msi14repair68 14)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi15_86=Office 15.0 MSI x86"      & call :getrepairsetup msi15repair86 15)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi15_68=Office 15.0 MSI x86/x64"  & call :getrepairsetup msi15repair68 15)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi16_86=Office 16.0 MSI x86"      & call :getrepairsetup msi16repair86 16)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi16_68=Office 16.0 MSI x86/x64"  & call :getrepairsetup msi16repair68 16)

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set "c2r15_86=Office 15.0 C2R x86"      & call :getc2rrepair c2r15repair86 integratedoffice.exe)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set "c2r15_68=Office 15.0 C2R x86/x64"  & call :getc2rrepair c2r15repair68 integratedoffice.exe)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set "c2r16_86=Office 16.0 C2R x86"      & call :getc2rrepair c2r16repair86 OfficeClickToRun.exe)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set "c2r16_68=Office 16.0 C2R x86/x64"  & call :getc2rrepair c2r16repair68 OfficeClickToRun.exe)

set uwp16=
if %winbuild% GEQ 10240 (
%psc% "Get-AppxPackage -name "Microsoft.Office.Desktop"" | find /i "Office" %nul1% && set uwp16=Office 16.0 UWP
)

set /a counter=0
echo Checking installed Office versions...
echo:

for %%# in (
"%msi14_68%"
"%msi14_86%"
"%msi15_68%"
"%msi15_86%"
"%msi16_68%"
"%msi16_86%"
"%c2r14_68%"
"%c2r14_86%"
"%c2r15_68%"
"%c2r15_86%"
"%c2r16_68%"
"%c2r16_86%"
"%uwp16%"
) do (
if not "%%#"=="""" (
set insoff=%%#
set insoff=!insoff:"=!
echo [!insoff!]
set /a counter+=1
)
)

if %counter% GTR 1 (
%eline%
echo Multiple Office versions found.
echo It is recommended to only install one version of Office.
echo ________________________________________________________________
echo:
)

if %counter% EQU 0 (
echo:
echo Office ^(2010 and later^) is not installed.
goto :repairend
echo:
) else (
echo:
call :dk_color %_Yellow% "A new window will appear, in that window you need to select [Quick Repair] option."
if defined terminal (
call :dk_color %_Yellow% "Press [0] to continue..."
choice /c 0 /n
) else (
call :dk_color %_Yellow% "Press any key to continue..."
pause %nul1%
)
)

if defined uwp16 (
echo:
echo Skipping repair for Office 16.0 UWP... 
echo:
)

set c2r14=
if defined c2r14_68 set c2r14=1
if defined c2r14_86 set c2r14=1

if defined c2r14 (
echo:
echo Skipping repair for Office 14.0 C2R...
echo:
)

if defined msi14_68 if exist "%msi14repair68%" echo Running - "%msi14repair68%"                    & "%msi14repair68%"
if defined msi14_86 if exist "%msi14repair86%" echo Running - "%msi14repair86%"                    & "%msi14repair86%"
if defined msi15_68 if exist "%msi15repair68%" echo Running - "%msi15repair68%"                    & "%msi15repair68%"
if defined msi15_86 if exist "%msi15repair86%" echo Running - "%msi15repair86%"                    & "%msi15repair86%"
if defined msi16_68 if exist "%msi16repair68%" echo Running - "%msi16repair68%"                    & "%msi16repair68%"
if defined msi16_86 if exist "%msi16repair86%" echo Running - "%msi16repair86%"                    & "%msi16repair86%"
if defined c2r15_68 if exist "%c2r15repair68%" echo Running - "%c2r15repair68%" REPAIRUI RERUNMODE & "%c2r15repair68%" REPAIRUI RERUNMODE
if defined c2r15_86 if exist "%c2r15repair86%" echo Running - "%c2r15repair86%" REPAIRUI RERUNMODE & "%c2r15repair86%" REPAIRUI RERUNMODE
if defined c2r16_68 if exist "%c2r16repair68%" echo Running - "%c2r16repair68%" scenario=Repair    & "%c2r16repair68%" scenario=Repair
if defined c2r16_86 if exist "%c2r16repair86%" echo Running - "%c2r16repair86%" scenario=Repair    & "%c2r16repair86%" scenario=Repair

:repairend

echo:
echo %line%
echo:
echo:
call :dk_color %Green% "Finished"
goto :at_back

:getc2rrepair

for %%# in (X86 X64) do (
if exist "%systemdrive%\Program Files\Microsoft Office 15\Client%%#\%2" (
set "%1=%systemdrive%\Program Files\Microsoft Office 15\Client%%#\%2"
)
)
exit /b

:getrepairsetup

set "_common86=%systemdrive%\Program Files (x86)\Common Files\Microsoft Shared\OFFICE%2\Office Setup Controller\setup.exe"
set "_common68=%systemdrive%\Program Files\Common Files\Microsoft Shared\OFFICE%2\Office Setup Controller\setup.exe"

if exist "%_common86%" set "%1=%_common86%"
if exist "%_common68%" set "%1=%_common68%"
exit /b

::========================================================================================================================================

:fixwmi

cls
if not defined terminal mode 98, 34
title  Fix WMI

::  https://techcommunity.microsoft.com/t5/ask-the-performance-team/wmi-repository-corruption-or-not/ba-p/375484

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo Rebuilding WMI is not recommended on Windows Server, aborting...
goto :at_back
)

echo:
echo Checking WMI
call :checkwmi

::  Apply basic fix first and check

if defined error (
%psc% Stop-Service Winmgmt -force %nul%
winmgmt /salvagerepository %nul%
call :checkwmi
)

if not defined error (
echo [Working]
echo No need to apply this option, aborting...
goto :at_back
)

call :dk_color %Red% "[Not Responding]"

set _corrupt=
sc start Winmgmt %nul%
if %errorlevel% EQU 1060 set _corrupt=1
sc query Winmgmt %nul% || set _corrupt=1
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type) do if not defined _corrupt (reg query HKLM\SYSTEM\CurrentControlSet\Services\Winmgmt /v %%G %nul% || set _corrupt=1)

echo:
if defined _corrupt (
%eline%
echo Winmgmt service is corrupted, aborting...
goto :at_back
)

echo Disabling Winmgmt service
sc config Winmgmt start= disabled %nul%
if %errorlevel% EQU 0 (
echo [Successful]
) else (
call :dk_color %Red% "[Failed] Aborting..."
sc config Winmgmt start= auto %nul%
goto :at_back
)

echo:
echo Stopping Winmgmt service
%psc% Stop-Service Winmgmt -force %nul%
%psc% Stop-Service Winmgmt -force %nul%
%psc% Stop-Service Winmgmt -force %nul%
sc query Winmgmt | find /i "STOPPED" %nul% && (
echo [Successful]
) || (
call :dk_color %Red% "[Failed]"
echo:
call :dk_color %Blue% "Its recommended to select [Restart] option and then apply Fix WMI option again."
echo %line%
echo:
choice /C:21 /N /M "> [1] Restart  [2] Revert Back Changes :"
if !errorlevel!==1 (sc config Winmgmt start= auto %nul%&goto :at_back)
echo:
echo Restarting...
shutdown -t 5 -r
exit
)

echo:
echo Deleting WMI repository
rmdir /s /q "%SysPath%\wbem\repository\" %nul%
if exist "%SysPath%\wbem\repository\" (
call :dk_color %Red% "[Failed]"
) else (
echo [Successful]
)

echo:
echo Enabling Winmgmt service
sc config Winmgmt start= auto %nul%
if %errorlevel% EQU 0 (
echo [Successful]
) else (
call :dk_color %Red% "[Failed]"
)

call :checkwmi
if not defined error (
echo:
echo Checking WMI
call :dk_color %Green% "[Working]"
goto :at_back
)

echo:
echo Registering .dll's and Compiling .mof's, .mfl's
call :registerobj %nul%

echo:
echo Checking WMI
call :checkwmi
if defined error (
call :dk_color %Red% "[Not Responding]"
echo:
echo Run [Dism RestoreHealth] and [SFC Scannow] options and make sure there are no errors.
) else (
call :dk_color %Green% "[Working]"
)

goto :at_back

:registerobj

::  https://eskonr.com/2012/01/how-to-fix-wmi-issues-automatically/

%psc% Stop-Service Winmgmt -force %nul%
cd /d %SysPath%\wbem\
regsvr32 /s %SysPath%\scecli.dll
regsvr32 /s %SysPath%\userenv.dll
mofcomp cimwin32.mof
mofcomp cimwin32.mfl
mofcomp rsop.mof
mofcomp rsop.mfl
for /f %%s in ('dir /b /s *.dll') do regsvr32 /s %%s
for /f %%s in ('dir /b *.mof') do mofcomp %%s
for /f %%s in ('dir /b *.mfl') do mofcomp %%s

winmgmt /salvagerepository
winmgmt /resetrepository
exit /b

:checkwmi

::  https://learn.microsoft.com/en-us/windows/win32/wmisdk/wmi-error-constants

set error=
%psc% "Get-WmiObject -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" %nul2% | find /i "computersystem" %nul1%
if %errorlevel% NEQ 0 (set error=1& exit /b)
winmgmt /verifyrepository %nul%
if %errorlevel% NEQ 0 (set error=1& exit /b)

%psc% "try { $null=([WMISEARCHER]'SELECT * FROM SoftwareLicensingService').Get().Version; exit 0 } catch { exit $_.Exception.InnerException.HResult }" %nul%
cmd /c exit /b %errorlevel%
echo "0x%=ExitCode%" | findstr /i "0x800410 0x800440 0x80131501" %nul1%
if %errorlevel% EQU 0 set error=1
exit /b

::========================================================================================================================================

:at_back

echo:
echo %line%
echo:
if defined terminal (
call :dk_color %_Yellow% "Press [0] key to %_exitmsg%..."
choice /c 0 /n
) else (
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
pause %nul1%
)
goto :at_menu

::========================================================================================================================================

:compresslog

::  https://stackoverflow.com/a/46268232

set "ddf="%SystemRoot%\Temp\ddf""
%nul% del /q /f %ddf%
echo/.New Cabinet>%ddf%
echo/.set Cabinet=ON>>%ddf%
echo/.set CabinetFileCountThreshold=0;>>%ddf%
echo/.set Compress=ON>>%ddf%
echo/.set CompressionType=LZX>>%ddf%
echo/.set CompressionLevel=7;>>%ddf%
echo/.set CompressionMemory=21;>>%ddf%
echo/.set FolderFileCountThreshold=0;>>%ddf%
echo/.set FolderSizeThreshold=0;>>%ddf%
echo/.set GenerateInf=OFF>>%ddf%
echo/.set InfFileName=nul>>%ddf%
echo/.set MaxCabinetSize=0;>>%ddf%
echo/.set MaxDiskFileCount=0;>>%ddf%
echo/.set MaxDiskSize=0;>>%ddf%
echo/.set MaxErrors=1;>>%ddf%
echo/.set RptFileName=nul>>%ddf%
echo/.set UniqueFiles=ON>>%ddf%
for /f "tokens=* delims=" %%D in ('dir /a:-D/b/s "%SystemRoot%\logs\%1"') do (
 echo/"%%~fD"  /inf=no;>>%ddf%
)
makecab /F %ddf% /D DiskDirectory1="" /D CabinetNameTemplate="!desktop!\%2_%_time%.cab"
del /q /f %ddf%
exit /b

::========================================================================================================================================

:checkperms

::  This code checks if SPP has permission access to tokens folder and required registry keys. Incorrect permissions are often set by gaming spoofers.

set permerror=
if not exist "%tokenstore%\" set "permerror=Error Found In Token Folder"

if defined ps32onArm exit /b

for %%# in (
"%tokenstore%+FullControl"
"HKLM:\SYSTEM\WPA+QueryValues, EnumerateSubKeys, WriteKey"
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform+SetValue"
) do for /f "tokens=1,2 delims=+" %%A in (%%#) do if not defined permerror (
%psc% "$acl = (Get-Acl '%%A' | fl | Out-String); if (-not ($acl -match 'NT SERVICE\\sppsvc Allow  %%B') -or ($acl -match 'NT SERVICE\\sppsvc Deny')) {Exit 2}" %nul%
if !errorlevel!==2 (
if "%%A"=="%tokenstore%" (
set "permerror=Error Found In Token Folder"
) else (
set "permerror=Error Found In SPP Registries"
)
)
)

REM  https://learn.microsoft.com/office/troubleshoot/activation/license-issue-when-start-office-application

if not defined permerror (
reg query "HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion" %nul% && (
set "pol=HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies"
reg query "!pol!" %nul% || reg add "!pol!" %nul%
%psc% "$netServ = (New-Object Security.Principal.SecurityIdentifier('S-1-5-20')).Translate([Security.Principal.NTAccount]).Value; $aclString = Get-Acl 'Registry::!pol!' | Format-List | Out-String; if (-not ($aclString.Contains($netServ + ' Allow  FullControl') -or $aclString.Contains('NT SERVICE\sppsvc Allow  FullControl')) -or ($aclString.Contains('Deny'))) {Exit 3}" %nul%
if !errorlevel!==3 set "permerror=Error Found In S-1-5-20 SPP"
)
)
exit /b

::========================================================================================================================================

::  Fix SPP related registry and folder permissions

:fixsppperms:
# Fix perms for Token Folder

if ($env:permerror -eq 'Error Found In Token Folder') {
    New-Item -Path $env:tokenstore -ItemType Directory -Force
    $sddl = 'O:BAG:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICIIO;GR;;;BU)(A;;FR;;;BU)(A;OICI;FA;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)'
    $AclObject = New-Object System.Security.AccessControl.DirectorySecurity
    $AclObject.SetSecurityDescriptorSddlForm($sddl)
    Set-Acl -Path $env:tokenstore -AclObject $AclObject
    exit
}

# Fix perms for SPP registries

if ($env:permerror -eq 'Error Found In SPP Registries') {
    $acl = Get-Acl 'HKLM:\SYSTEM\WPA'
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc', 'QueryValues, EnumerateSubKeys, WriteKey', 'ContainerInherit, ObjectInherit', 'None', 'Allow')
    $acl.ResetAccessRule($rule)
    $acl.SetAccessRule($rule)
    Set-Acl -Path 'HKLM:\SYSTEM\WPA' -AclObject $acl
	
    $acl = Get-Acl 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc', 'SetValue', 'ContainerInherit, ObjectInherit', 'None', 'Allow')
    $acl.ResetAccessRule($rule)
    $acl.SetAccessRule($rule)
    Set-Acl -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform' -AclObject $acl
    exit
}

# Fix perms for SPP in HKU\S-1-5-20
# https://learn.microsoft.com/office/troubleshoot/activation/license-issue-when-start-office-application

if ($env:permerror -ne 'Error Found In S-1-5-20 SPP') {
    exit
}
if (-not (Test-Path 'Registry::HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform')) {
    exit
}

# https://stackoverflow.com/a/35843420

function Take-Permissions {
    param($rootKey, $key, [System.Security.Principal.SecurityIdentifier]$sid = 'S-1-5-32-545', $recurse = $true)
    
    switch -regex ($rootKey) {
        'HKCU|HKEY_CURRENT_USER' { $rootKey = 'CurrentUser' }
        'HKLM|HKEY_LOCAL_MACHINE' { $rootKey = 'LocalMachine' }
        'HKCR|HKEY_CLASSES_ROOT' { $rootKey = 'ClassesRoot' }
        'HKCC|HKEY_CURRENT_CONFIG' { $rootKey = 'CurrentConfig' }
        'HKU|HKEY_USERS' { $rootKey = 'Users' }
    }

    ### Step 1 - escalate current process's privilege
    # get SeTakeOwnership, SeBackup and SeRestore privileges before executes next lines, script needs Admin privilege
    $AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
    $TypeBuilder = $ModuleBuilder.DefineType(0)
    $TypeBuilder.DefinePInvokeMethod('RtlAdjustPrivilege', 'ntdll.dll', 'Public, Static', 1, [int], @([int], [bool], [bool], [bool].MakeByRefType()), 1, 3) | Out-Null
    9, 17, 18 | ForEach-Object { $TypeBuilder.CreateType()::RtlAdjustPrivilege($_, $true, $false, [ref]$false) | Out-Null }

    function Take-KeyPermissions {
        param($rootKey, $key, $sid, $recurse, $recurseLevel = 0)

        ### Step 2 - get ownerships of key - it works only for current key
        $regKey = [Microsoft.Win32.Registry]::$rootKey.OpenSubKey($key, 'ReadWriteSubTree', 'TakeOwnership')
        $acl = New-Object System.Security.AccessControl.RegistrySecurity
        $acl.SetOwner($sid)
        $regKey.SetAccessControl($acl)

        ### Step 3 - enable inheritance of permissions (not ownership) for current key from parent
        $acl.SetAccessRuleProtection($false, $false)
        $regKey.SetAccessControl($acl)

        ### Step 4 - only for top-level key, change permissions for current key and propagate it for subkeys
        # to enable propagations for subkeys, it needs to execute Steps 2-3 for each subkey (Step 5)
        if ($recurseLevel -eq 0) {
            $regKey = $regKey.OpenSubKey('', 'ReadWriteSubTree', 'ChangePermissions')
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule($sid, 'FullControl', 'ContainerInherit', 'None', 'Allow')
            $acl.ResetAccessRule($rule)
            $regKey.SetAccessControl($acl)
        }

        ### Step 5 - recursively repeat steps 2-5 for subkeys
        if ($recurse) {
            foreach ($subKey in $regKey.OpenSubKey('').GetSubKeyNames()) {
                Take-KeyPermissions $rootKey ($key + '\' + $subKey) $sid $recurse ($recurseLevel + 1)
            }
        }
    }

    Take-KeyPermissions $rootKey $key $sid $recurse
}

Take-Permissions "Users" "S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" "S-1-5-20"
:fixsppperms:

::========================================================================================================================================

::  Install License files using Powershell/WMI instead of slmgr.vbs

:xrm:
function InstallLicenseFile($Lsc) {
    try {
        $null = $sls.InstallLicense([IO.File]::ReadAllText($Lsc))
    } catch {
        $host.SetShouldExit($_.Exception.HResult)
    }
}
function InstallLicenseArr($Str) {
    $a = $Str -split ';'
    ForEach ($x in $a) {InstallLicenseFile "$x"}
}
function InstallLicenseDir($Loc) {
    dir $Loc *.xrm-ms -af -s | select -expand FullName | % {InstallLicenseFile "$_"}
}
function ReinstallLicenses() {
    $Oem = "$env:SysPath\oem"
    $Spp = "$env:SysPath\spp\tokens"
    InstallLicenseDir "$Spp"
    If (Test-Path $Oem) {InstallLicenseDir "$Oem"}
}
:xrm:

::========================================================================================================================================

:scandat

set token=
for %%# in (
%SysPath%\spp\store_test\2.0\
%SysPath%\spp\store\
%SysPath%\spp\store\2.0\
%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\
) do (

if %1==check (
if exist %%#tokens.dat set token=%%#tokens.dat
)

if %1==delete (
if exist %%# (
%nul% dir /a-d /s "%%#*.dat" && (
attrib -r -s -h "%%#*.dat" /S
del /S /F /Q "%%#*.dat"
)
)
)
)
exit /b

:scandatospp

set token=
for %%# in (
%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform\
) do (

if %1==check (
if exist %%#tokens.dat set token=%%#tokens.dat
)

if %1==delete (
if exist %%# (
%nul% dir /a-d /s "%%#*.dat" && (
attrib -r -s -h "%%#*.dat" /S
del /S /F /Q "%%#*.dat"
)
)
)
)
exit /b

::========================================================================================================================================

:regownstart

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regown\:.*';iex ($f[1]);"
exit /b

::  Below code takes ownership of a volatile registry key and deletes it
::  HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState

:regown:
$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TypeBuilder = $ModuleBuilder.DefineType(0)

$TypeBuilder.DefinePInvokeMethod('RtlAdjustPrivilege', 'ntdll.dll', 'Public, Static', 1, [int], @([int], [bool], [bool], [bool].MakeByRefType()), 1, 3) | Out-Null
$TypeBuilder.CreateType()::RtlAdjustPrivilege(9, $true, $false, [ref]$false) | Out-Null

$SID = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
$IDN = ($SID.Translate([System.Security.Principal.NTAccount])).Value
$Admin = New-Object System.Security.Principal.NTAccount($IDN)

$path = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState'
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey($path, 'ReadWriteSubTree', 'takeownership')

$acl = $key.GetAccessControl()
$acl.SetOwner($Admin)
$key.SetAccessControl($acl)

$rule = New-Object System.Security.AccessControl.RegistryAccessRule($Admin,"FullControl","Allow")
$acl.SetAccessRule($rule)
$key.SetAccessControl($acl)
:regown:

::========================================================================================================================================

:dk_done

echo:
if defined fixes (
call :dk_color %White% "Follow ALL the ABOVE blue lines.   "
call :dk_color2 %Blue% "Press [1] to Open Support Webpage " %Gray% " Press [0] to Ignore"
choice /C:10 /N
if !errorlevel!==1 (for %%# in (%fixes%) do (start %%#))
)

if defined terminal (
call :dk_color %_Yellow% "Press [0] key to %_exitmsg%..."
choice /c 0 /n
) else (
call :dk_color %_Yellow% "Press any key to %_exitmsg%..."
pause %nul1%
)
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

::  Set variables

:dk_setvar

set psc=powershell.exe
set winbuild=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 %nul2% | find /i "0x0" %nul1% && (set _NCS=0)

echo "%PROCESSOR_ARCHITECTURE% %PROCESSOR_ARCHITEW6432%" | find /i "ARM64" %nul1% && (if %winbuild% LSS 21277 set ps32onArm=1)

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set   "White="107;91m""
set    "_Red="40;91m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set   "White="White" "Red""
set    "_Red="Black" "Red""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
if %~z0 GEQ 200000 (
set "_exitmsg=Go back"
set "_fixmsg=Go back to Main Menu, select Troubleshoot and run Fix Licensing option."
) else (
set "_exitmsg=Exit"
set "_fixmsg=In MAS folder, run Troubleshoot script and select Fix Licensing option."
)
exit /b

::========================================================================================================================================
:: Leave empty line below
