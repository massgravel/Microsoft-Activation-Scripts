@set masver=2.8
@echo off



::============================================================================
::
::   Homepage: mass grave[.]dev
::      Email: mas.help@outlook.com
::
::============================================================================



::  To stage current edition while changing edition with CBS Upgrade Method, change 0 to 1 in below line
set _stg=0



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
title  Change Windows Edition %masver%

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
set "line=echo ___________________________________________________________________________________________"

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

cls
if not defined terminal mode 98, 30
title  Change Windows Edition %masver%

echo:
echo Initializing...
echo:

for %%# in (
sppsvc.exe
dism.exe
) do (
if not exist %SysPath%\%%# (
%eline%
echo [%SysPath%\%%#] file is missing, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)
)

::========================================================================================================================================

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_reflection
call :dk_ckeckwmic
call :dk_sppissue

for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if not "%%j"=="" (
set fullbuild=%%i.%%j
) else (
for /f "tokens=3" %%G in ('"reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR" %nul6%') do if not errorlevel 1 set /a "UBR=%%G"
for /f "skip=2 tokens=3,4 delims=. " %%G in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildLabEx') do (
if defined UBR (set "fullbuild=%%G.!UBR!") else (set "fullbuild=%%G.%%H")
)
)

::========================================================================================================================================

::  Check Activation ID

call :dk_actid 55c92734-d682-4d71-983e-d6ec3f16059f
if not defined apps (
%eline%
echo Either key is not insalled or script failed to get installed key's activation ID. Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

::  Check Windows Edition and branch

set osedition=
set dismnotworking=

for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition %nul6% ^| find /i "Current Edition :"') do set "osedition=%%a"
if not defined osedition set dismnotworking=1

if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL) get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
if not defined osedition %chkedi% do if not errorlevel 1 (call set "osedition=%%a")

if not defined osedition (
%eline%
echo Failed to detect OS edition, aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

for /f "skip=2 tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul6%') do set "regedition=%%a"
if /i not "%osedition%"=="%regedition%" (
set "showeditionerror=call :dk_color %_Yellow% "[%osedition%] [Reg-%regedition%].""
)

set branch=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildBranch %nul6%') do set "branch=%%b"

::========================================================================================================================================

::  Get target editions list

set _target=
set _dtarget=
set _ptarget=
set _ntarget=
set _wtarget=

if %winbuild% GEQ 10240 for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (if defined _dtarget (set "_dtarget= !_dtarget! %%a ") else (set "_dtarget= %%a "))
if %winbuild% LSS 10240 for /f "tokens=4" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':cbsxml\:.*';& ([ScriptBlock]::Create($f[1])) -GetTargetEditions;" ^| findstr /i /c:"Target Edition : "') do (if defined _ptarget (set "_ptarget= !_ptarget! %%a ") else (set "_ptarget= %%a "))

if %winbuild% GEQ 10240 if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
call :ced_edilist
if /i "%osedition:~0,4%"=="Core" set _pro=Professional
if /i "%osedition%"=="CoreN" set _pro=ProfessionalN
set "_dtarget= %_dtarget% !_wtarget! !_pro! "
)

::========================================================================================================================================

for %%# in (CloudEdition CloudEditionN ServerRdsh) do if /i %osedition%==%%# (
cls
echo:
call :dk_color %Red% "==== Note ===="
echo:
echo [EditionID:%osedition% ^| %fullbuild%]
echo:
echo Changing this edition may not remove "%osedition%"-specific features.
echo:
call :dk_color %_Yellow% "Press [7] to continue anyway..."
choice /c 7 /n
cls
)

for %%# in ( %_dtarget% %_ptarget% ) do if /i not "%%#"=="%osedition%" (
echo "!_target!" | find /i " %%# " %nul1% || set "_target= !_target! %%# "
)

if defined _target (
for %%# in (%_target%) do (
echo %%# | findstr /i "CountrySpecific CloudEdition" %nul% || (set "_ntarget=!_ntarget! %%#")
)
)

if not defined _ntarget (
%line%
echo:
if defined dismnotworking call :dk_color %Red% "DISM.exe is not working."
call :dk_color %Gray% "Target editions not found."
echo Current Edition [%osedition% ^| %winbuild%] can not be changed to any other Edition.
%line%
goto dk_done
)

::========================================================================================================================================

:cedmenu2

cls
if not defined terminal mode 98, 30
set inpt=
set counter=0
set verified=0
set targetedition=

%line%
echo:
call :dk_color %Gray% "You can change the edition [%osedition%] [%fullbuild%] to one of the following."
%showeditionerror%
if defined dismnotworking (
call :dk_color %_Yellow% "Note - DISM.exe is not working."
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
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
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
if not defined terminal mode con cols=105 lines=32

if /i "%targetedition%"=="ServerRdsh" (
echo:
call :dk_color %Red% "==== Note ===="
echo:
echo Once the edition is changed to "%targetedition%", 
echo the system may not be able to properly change edition later.
echo:
echo [1] Continue Anyway
echo [0] Go Back
echo:
call :dk_color %_Green% "Choose a menu option using your keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 goto cedmenu2
if !errorlevel!==1 rem
)

cls
set key=
set _chan=
set _dismapi=0

::  Check if DISM API or slmgr.vbs is required for edition upgrade

if not exist "%SysPath%\spp\tokens\skus\%targetedition%\" (
echo %_wtarget% | find /i " %targetedition% " || (
set _dismapi=1
)
)

set "keyflow=Retail Volume:GVLK Volume:MAK OEM:NONSLP OEM:DM PGS:TB Retail:TB:Eval"

call :ced_targetSKU %targetedition%
if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%
if not defined key call :changeeditiondata
if not defined key if %_dismapi%==1 if /i "%targetedition%"=="Professional" (
set key=VK7JG-NPHTM-C97JM-9MPGT-3V66T
set _chan=Retail
)

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Failed to get product key from pkeyhelper.dll.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

::  Changing from Core to Non-Core & Changing editions in Windows build older than 17134 requires "changepk /productkey" or DISM Api method and restart
::  In other cases, editions can be changed instantly with "slmgr /ipk"

if %_dismapi%==1 (
if not defined terminal mode con cols=105 lines=40
call :ced_rebootflag
if defined rebootreq goto dk_done
)

cls
%line%
echo:
%showeditionerror%
if defined dismnotworking call :dk_color %_Yellow% "DISM.exe is not working."
echo Changing the current edition [%osedition%] %fullbuild% to [%targetedition%]...
echo:

if %_dismapi%==1 (
call :dk_color %Green% "Notes -"
echo:
echo  - Save your work before continuing, the system will auto-restart.
echo:
echo  - You will need to activate with HWID option once the edition is changed.
%line%
echo:
choice /C:21 /N /M "[1] Continue [2] %_exitmsg% : "
if !errorlevel!==1 exit /b
)

::========================================================================================================================================

if %_dismapi%==0 (
echo Installing %_chan% key [%key%]
echo:
if %_wmic% EQU 1 wmic path %sps% where __CLASS='%sps%' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "try { $null=(([WMISEARCHER]'SELECT Version FROM %sps%').Get()).InstallProductKey('%key%'); exit 0 } catch { exit $_.Exception.InnerException.HResult }" %nul%
set keyerror=!errorlevel!
cmd /c exit /b !keyerror!
if !keyerror! NEQ 0 set "keyerror=[0x!=ExitCode!]"

if !keyerror! EQU 0 (
call :dk_refresh
call :dk_color %Green% "[Successful]"
echo:
call :dk_color %Gray% "Reboot is required to fully change the edition."
) else (
call :dk_color %Red% "[Unsuccessful] [Error Code: !keyerror!]"
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)
)

if %_dismapi%==1 (
echo:
echo Applying the DISM API method with %_chan% key %key%. Please wait...
echo:

call :ced_prep
if defined preperror goto dk_done

%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':dismapi\:.*';& ([ScriptBlock]::Create($f[1])) %targetedition% %key%"
call :ced_postprep
)
%line%

goto dk_done

::========================================================================================================================================

:cbsmethod

cls
if not defined terminal (
mode con cols=105 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)

call :ced_rebootflag
if defined rebootreq goto dk_done

echo:
%showeditionerror%
if defined dismnotworking call :dk_color %_Yellow% "Note - DISM.exe is not working."
echo Changing the current edition [%osedition%] %fullbuild% to [%targetedition%]...
echo:
call :dk_color %Blue% "Important - Save your work before continuing, the system will auto-restart."
echo:
choice /C:01 /N /M "[1] Continue [0] %_exitmsg% : "
if %errorlevel%==1 exit /b

echo:
echo Initializing...
echo:

call :ced_prep
if defined preperror goto dk_done

if %_stg%==0 (set stage=) else (set stage=-StageCurrent)
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':cbsxml\:.*';& ([ScriptBlock]::Create($f[1])) -SetEdition %targetedition% %stage%"
call :ced_postprep
%line%

goto dk_done

::========================================================================================================================================

:ced_change_server

cls
if not defined terminal (
mode con cols=105 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)

set key=
set _chan=
set "keyflow=Volume:GVLK Retail Volume:MAK OEM:NONSLP OEM:DM PGS:TB Retail:TB:Eval"

call :ced_targetSKU %targetedition%
if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%
if not defined key call :changeeditiondata

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo Failed to get product key from pkeyhelper.dll.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

call :ced_rebootflag
if defined rebootreq goto dk_done

cls
echo:
%showeditionerror%
if defined dismnotworking call :dk_color %_Yellow% "Note - DISM.exe is not working."
echo Changing the current edition [%osedition%] %fullbuild% to [%targetedition%]...
echo:

call :ced_prep
if defined preperror goto dk_done

echo Applying the command with %_chan% key...
echo DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula
DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula

call :ced_postprep
%line%

goto dk_done

::========================================================================================================================================

:ced_prep

set _time=
set preperror=

for /f %%a in ('%psc% "(Get-Date).ToString('yyyyMMdd-HHmmssfff')"') do set _time=%%a

%psc% Stop-Service TrustedInstaller -force %nul%

sc query TrustedInstaller | find /i "RUNNING" %nul% && (
%eline%
echo Failed to stop the TrustedInstaller service.
echo Reboot your machine using the restart option and try again.
set preperror=1
exit /b
)

copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\backup_cbs_%_time%.log" %nul%
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "%SystemRoot%\logs\DISM\backup_dism_%_time%.log" %nul%

del /f /q "%SystemRoot%\logs\cbs\cbs.log" %nul%
del /f /q "%SystemRoot%\logs\DISM\dism.log" %nul%

:: Initiate this to appear in fresh logs

dism /online /english /Get-CurrentEdition %nul%
dism /online /english /Get-TargetEditions %nul%
exit /b

::========================================================================================================================================

:ced_postprep

timeout /t 5 %nul1%
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "%SystemRoot%\logs\cbs\cbs_%_time%.log" %nul%
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "%SystemRoot%\logs\DISM\dism_%_time%.log" %nul%

if not exist "!desktop!\ChangeEdition_Logs\" md "!desktop!\ChangeEdition_Logs\" %nul%
call :compresslog cbs\cbs_%_time%.log ChangeEdition_Logs\CBS %nul%
call :compresslog DISM\dism_%_time%.log ChangeEdition_Logs\DISM %nul%

echo:
if %winbuild% GEQ 9200 %psc% "if ((Get-WindowsOptionalFeature -Online -FeatureName NetFx3).State -eq 'Enabled') {Write-Host 'Checking .NET Framework 3.5 Status - Enabled'}"
echo Log files are copied to the ChangeEdition_Logs folder on your desktop.
echo:
call :dk_color %Blue% "In case there are errors, you should restart the system before trying again."
echo:
set fixes=%fixes% %mas%change_edition_issues
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%change_edition_issues"
exit /b

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

::  Refresh license status

:dk_refresh

if %_wmic% EQU 1 wmic path %sps% where __CLASS='%sps%' call RefreshLicenseStatus %nul%
if %_wmic% EQU 0 %psc% "$null=(([WMICLASS]'%sps%').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::  Get installed products Activation IDs

:dk_actid

set apps=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='%1' and PartialProductKey is not null) get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %spp% WHERE ApplicationID=''%1'' AND PartialProductKey IS NOT NULL').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined apps (call set "apps=!apps! %%a") else (call set "apps=%%a"))
exit /b

::  Get Edition list

:ced_edilist

if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL) get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND LicenseDependsOn is NULL').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
%chkedi% do call set "_wtarget= !_wtarget! %%a "
exit /b

::  Check wmic.exe

:dk_ckeckwmic

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
cmd /c "wmic path Win32_ComputerSystem get CreationClassName /value" %nul2% | find /i "computersystem" %nul1% && set _wmic=1
)
exit /b

::  Show info for potential script stuck scenario

:dk_sppissue

sc start sppsvc %nul%
set spperror=%errorlevel%

if %spperror% NEQ 1056 if %spperror% NEQ 0 (
%eline%
echo sc start sppsvc [Error Code: %spperror%]
)

echo:
%psc% "$job = Start-Job { (Get-WmiObject -Query 'SELECT * FROM %sps%').Version }; if (-not (Wait-Job $job -Timeout 30)) {write-host 'sppsvc is not working correctly. Help - %mas%troubleshoot'}"
exit /b

::  Common lines used in PowerShell reflection code

:dk_reflection

set ref=$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1);
set ref=%ref% $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False);
set ref=%ref% $TypeBuilder = $ModuleBuilder.DefineType(0);
exit /b

::========================================================================================================================================

::  Check pending reboot flags

:ced_rebootflag

set rebootreq=
reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" %nul% && set rebootreq=1
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" %nul% && set rebootreq=1

if defined rebootreq (
%eline%
echo Pending reboot flags found.
echo:
echo Make sure Windows is fully updated, restart the system and try again.
)
exit /b

::========================================================================================================================================

::  Get Product Key from pkeyhelper.dll for future new editions
::  It works on Windows 10 1803 (17134) and later builds.

:k_pkey

call :dk_reflection

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SkuGetProductKeyForEdition', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([int], [String], [String].MakeByRefType(), [String].MakeByRefType()), 1, 3);
set d1=%d1% $out = ''; [void]$TypeBuilder.CreateType()::SkuGetProductKeyForEdition(%1, %2, [ref]$out, [ref]$null); $out

set pkey=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkey=%%a)
exit /b

::  Get channel name for the key which was extracted from pkeyhelper.dll

:k_pkeychannel

set k=%1
set m=[Runtime.InteropServices.Marshal]
set p=%SysPath%\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('PidGenX', 'pidgenx.dll', 'Public, Static', 1, [int], @([String], [String], [String], [int], [IntPtr], [IntPtr], [IntPtr]), 1, 3);
set d1=%d1% $r = [byte[]]::new(0x04F8); $r[0] = 0xF8; $r[1] = 0x04; $f = %m%::AllocHGlobal(0x04F8); %m%::Copy($r, 0, $f, 0x04F8);
set d1=%d1% [void]$TypeBuilder.CreateType()::PidGenX('%k%', '%p%', '00000', 0, 0, 0, $f); %m%::Copy($f, $r, 0, 0x04F8); %m%::FreeHGlobal($f); [Text.Encoding]::Unicode.GetString($r, 1016, 128)

set pkeychannel=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkeychannel=%%a)
exit /b

:ced_windowskey

for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b
for %%# in (%keyflow%) do (
call :k_pkey %targetSKU% '%%#'
if defined pkey call :k_pkeychannel !pkey!
if /i "!pkeychannel!"=="%%#" (
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

:dk_done

echo:
if %_unattended%==1 timeout /t 2 & exit /b

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

::  https://github.com/asdcorp/Set-WindowsCbsEdition

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
::  For Windows, generic keys are mentioned till 22000 and for Server, generic keys are mentioned till 17763, later ones are extracted from the pkeyhelper.dll

:changeeditiondata

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
if %winbuild% GTR 17763 exit /b
) else (
if %winbuild% GEQ 22000 exit /b
)
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" (set Cor=Cor) else (set Cor=)

set h=
for %%# in (
XGVPP-NMH47-7TTHJ-W3FW7-8HV%h%2C__OEM:NONSLP_Enterprise
D6RD9-D4N8T-RT9QX-YW6YT-FCW%h%WJ______Retail_Starter
3V6Q6-NQXCX-V8YXR-9QCYV-QPF%h%CT__Volume:MAK_EnterpriseN
3NFXW-2T27M-2BDW6-4GHRV-68X%h%RX______Retail_StarterN
VK7JG-NPHTM-C97JM-9MPGT-3V6%h%6T______Retail_Professional
2B87N-8KFHP-DKV6R-Y2C8J-PKC%h%KT______Retail_ProfessionalN
4CPRK-NM3K3-X6XXQ-RXX86-WXC%h%HW______Retail_CoreN
N2434-X9D7W-8PF6X-8DV9T-8TY%h%MD______Retail_CoreCountrySpecific
BT79Q-G7N6G-PGBYW-4YWX6-6F4%h%BT______Retail_CoreSingleLanguage
YTMG3-N6DKC-DKB77-7M9GH-8HV%h%X7______Retail_Core
XKCNC-J26Q9-KFHD2-FKTHY-KD7%h%2Y__OEM:NONSLP_PPIPro
YNMGQ-8RYV3-4PGQ3-C8XTP-7CF%h%BY______Retail_Education
84NGF-MHBT6-FXBX8-QWJK7-DRR%h%8H______Retail_EducationN
KCNVH-YKWX8-GJJB9-H9FDT-6F7%h%W2__Volume:MAK_EnterpriseS_VB
43TBQ-NH92J-XKTM7-KT3KK-P39%h%PB__OEM:NONSLP_EnterpriseS_RS5
NK96Y-D9CD8-W44CQ-R8YTK-DYJ%h%WX__OEM:NONSLP_EnterpriseS_RS1
FWN7H-PF93Q-4GGP8-M8RF3-MDW%h%WW__OEM:NONSLP_EnterpriseS_TH
RQFNW-9TPM3-JQ73T-QV4VQ-DV9%h%PT__Volume:MAK_EnterpriseSN_VB
M33WV-NHY3C-R7FPM-BQGPT-239%h%PG__Volume:MAK_EnterpriseSN_RS5
2DBW3-N2PJG-MVHW3-G7TDK-9HK%h%R4__Volume:MAK_EnterpriseSN_RS1
NTX6B-BRYC2-K6786-F6MVQ-M7V%h%2X__Volume:MAK_EnterpriseSN_TH
G3KNM-CHG6T-R36X3-9QDG6-8M8%h%K9______Retail_ProfessionalSingleLanguage
HNGCC-Y38KG-QVK8D-WMWRK-X86%h%VK______Retail_ProfessionalCountrySpecific
DXG7C-N36C4-C4HTG-X4T3X-2YV%h%77______Retail_ProfessionalWorkstation
WYPNQ-8C467-V2W6J-TX4WX-WT2%h%RQ______Retail_ProfessionalWorkstationN
8PTT6-RNW4C-6V7J2-C2D3X-MHB%h%PB______Retail_ProfessionalEducation
GJTYN-HDMQY-FRR76-HVGC7-QPF%h%8P______Retail_ProfessionalEducationN
C4NTJ-CX6Q2-VXDMR-XVKGM-F9D%h%JC__Volume:MAK_EnterpriseG
46PN6-R9BK9-CVHKB-HWQ9V-MBJ%h%Y8__Volume:MAK_EnterpriseGN
NJCF7-PW8QT-3324D-688JX-2YV%h%66______Retail_ServerRdsh
XQQYW-NFFMW-XJPBH-K8732-CKF%h%FD______OEM:DM_IoTEnterprise
QPM6N-7J2WJ-P88HH-P3YRH-YY7%h%4H__OEM:NONSLP_IoTEnterpriseS
K9VKN-3BGWV-Y624W-MCRMQ-BHD%h%CD______Retail_CloudEditionN
KY7PN-VR6RX-83W6Y-6DDYQ-T6R%h%4W______Retail_CloudEdition
V3WVW-N2PV2-CGWC3-34QGF-VMJ%h%2C______Retail_Cloud
NH9J3-68WK7-6FB93-4K3DF-DJ4%h%F6______Retail_CloudN
2HN6V-HGTM8-6C97C-RK67V-JQP%h%FD______Retail_CloudE
WC2BQ-8NRM3-FDDYY-2BFGV-KHK%h%QY_Volume:GVLK_ServerStandard%Cor%_RS1
CB7KF-BWN84-R7R2Y-793K2-8XD%h%DG_Volume:GVLK_ServerDatacenter%Cor%_RS1
JCKRF-N37P4-C2D82-9YXRT-4M6%h%3B_Volume:GVLK_ServerSolution_RS1
QN4C6-GBJD2-FB422-GHWJK-GJG%h%2R_Volume:GVLK_ServerCloudStorage_RS1
VP34G-4NPPG-79JTQ-864T4-R3M%h%QX_Volume:GVLK_ServerAzureCor_RS1
9JQNQ-V8HQ6-PKB8H-GGHRY-R62%h%H6______Retail_ServerAzureNano_RS1
VN8D3-PR82H-DB6BJ-J9P4M-92F%h%6J______Retail_ServerStorageStandard_RS1
48TQX-NVK3R-D8QR3-GTHHM-8FH%h%XC______Retail_ServerStorageWorkgroup_RS1
2HXDN-KRXHB-GPYC7-YCKFJ-7FV%h%DG_Volume:GVLK_ServerDatacenterACor_RS3
PTXN8-JFHJM-4WC78-MPCBR-9W4%h%KR_Volume:GVLK_ServerStandardACor_RS3
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
