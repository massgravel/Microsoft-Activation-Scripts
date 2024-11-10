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
title  Change Office Edition %masver%

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

cls
if not defined terminal mode 98, 30
title  Change Office Edition %masver%

echo:
echo Initializing...
echo:

if not exist %SysPath%\sppsvc.exe (
%eline%
echo [%SysPath%\sppsvc.exe] file is missing. Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
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

::  Check Windows Edition
::  This is just to ensure that SPP/WMI are functional

cls
set osedition=0
if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path %spp% where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL) get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM %spp% WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND LicenseDependsOn is NULL AND PartialProductKey IS NOT NULL').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
%chkedi% do if not errorlevel 1 (call set "osedition=%%a")

if %osedition%==0 (
%eline%
echo Failed to detect OS Edition. Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

::  Check installed Office 16.0 C2R

set o16c2r=
set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (set o16c2r=1&set o16c2r_reg=%_86%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (set o16c2r=1&set o16c2r_reg=%_68%\ClickToRun)

if not defined o16c2r_reg (
%eline%
echo Office C2R 2016 or later is not installed, which is required for this script.
echo Download and install Office from below URL and try again.
echo:
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto dk_done
)

call :ch_getinfo

::========================================================================================================================================

::  Check minimum required details

if %verchk% LSS 9029 (
%eline%
echo Installed Office version is %_version%.
echo Minimum required version is 16.0.9029.2167
echo Aborting...
echo:
call :dk_color %Blue% "Download and install latest Office from below URL and try again."
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto dk_done
)

for %%A in (
_oArch
_updch
_lang
_clversion
_version
_oIds
_c2rXml
_c2rExe
_c2rCexe
_masterxml
) do (
if not defined %%A (
%eline%
echo Failed to find %%A. Aborting...
echo:
call :dk_color %Blue% "Download and install Office from below URL and try again."
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto dk_done
)
)

if %winbuild% LSS 10240 if defined ltscfound (
%eline%
echo Installed Office appears to be from the Volume channel %ltsc19%%ltsc21%%ltsc24%,
echo which is not officially supported on your Windows build version %winbuild%.
echo Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

set unsupbuild=
if %winbuild% LSS 10240 if %winbuild% GEQ 9200 if %verchk% GTR 16026 set unsupbuild=1
if %winbuild% LSS 9200 if %verchk% GTR 12527 set unsupbuild=1

if defined unsupbuild (
%eline%
echo Unsupported Office %verchk% is installed on your Windows build version %winbuild%.
echo Aborting...
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

:oemenu

cls
set fixes=
if not defined terminal mode 76, 25
title  Change Office Edition %masver%
echo:
echo:
echo:
echo:
echo         ____________________________________________________________
echo:
echo                 [1] Change all editions
echo                 [2] Add edition
echo                 [3] Remove edition
echo:
echo                 [4] Add/Remove apps
echo                 ____________________________________________
echo:
echo                 [5] Change Office Update Channel
echo                 [0] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "           " %_Green% "Choose a menu option using your keyboard [1,2,3,4,5,0]"
choice /C:123450 /N
set _el=!errorlevel!
if !_el!==6  exit /b
if !_el!==5  goto :oe_changeupdchnl
if !_el!==4  goto :oe_editedition
if !_el!==3  goto :oe_removeedition
if !_el!==2  set change=0& goto :oe_edition
if !_el!==1  set change=1& goto :oe_edition
goto :oemenu

::========================================================================================================================================

:oe_edition

cls
call :oe_chkinternet
if not defined _int (
goto :oe_goback
)

cls
if not defined terminal mode 76, 25
if %change%==1 (
title  Change all editions %masver%
) else (
title  Add edition %masver%
)

echo:
echo:
echo:
echo:
echo                 O365/Mondo editions have the latest features.     
echo         ____________________________________________________________
echo:
echo                 [1] Office Suites     - Retail
echo                 [2] Office Suites     - Volume
echo                 [3] Office SingleApps - Retail
echo                 [4] Office SingleApps - Volume
echo                 ____________________________________________
echo:
echo                 [0] Go Back
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "            " %_Green% "Choose a menu option using your keyboard [1,2,3,4,0]"
choice /C:12340 /N
set _el=!errorlevel!
if !_el!==5  goto :oemenu
if !_el!==4  set list=SingleApps_Volume&goto :oe_editionchangepre
if !_el!==3  set list=SingleApps_Retail&goto :oe_editionchangepre
if !_el!==2  set list=Suites_Volume&goto :oe_editionchangepre
if !_el!==1  set list=Suites_Retail&goto :oe_editionchangepre
goto :oe_edition

::========================================================================================================================================

:oe_editionchangepre

cls
call :ch_getinfo
call :oe_tempcleanup
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':getlist\:.*';iex ($f[1])"

:oe_editionchange

cls
if not defined terminal (
mode 98, 45
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=44;$B.Height=100;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)

if not exist %SystemRoot%\Temp\%list%.txt (
%eline%
echo Failed to generate available editions list.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto :oe_goback
)

set inpt=
set counter=0
set verified=0
set _notfound=
set targetedition=

%line%
echo:
call :dk_color %Gray% "Installed Office editions: %_oIds%"
call :dk_color %Gray% "You can select one of the following Office Editions."
if %winbuild% LSS 10240 (
echo Unsupported products such as 2019/2021/2024 are excluded from this list.
) else (
for %%# in (2019 2021 2024) do (
find /i "%%#" "%SystemRoot%\Temp\%list%.txt" %nul1% || (
if defined _notfound (set _notfound=%%#, !_notfound!) else (set _notfound=%%#)
)
)
if defined _notfound call :dk_color %Gray% "Office !_notfound! is not in this list because old version [%_version%] of Office is installed."
)
%line%
echo:

for /f "usebackq delims=" %%A in (%SystemRoot%\Temp\%list%.txt) do (
set /a counter+=1
if !counter! LSS 10 (
echo [!counter!]  %%A
) else (
echo [!counter!] %%A
)
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  Go Back
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto :oe_editionchange
if "%inpt%"=="0" (call :oe_tempcleanup & goto :oe_edition)
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto :oe_editionchange

::========================================================================================================================================

::  Set app exclusions

:oe_excludeappspre

cls
set suites=
echo %list% | find /i "Suites" %nul1% && (
set suites=1
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':getappnames\:.*';iex ($f[1])"
if not exist %SystemRoot%\Temp\getAppIds.txt (
%eline%
echo Failed to generate available apps list.
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
goto :oe_goback
)
)

for %%# in (
Access
Excel
Lync
OneNote
Outlook
PowerPoint
Project
Publisher
Visio
Word
) do (
if defined suites (
find /i "%%#" "%SystemRoot%\Temp\getAppIds.txt" %nul1% && (set %%#_st=On) || (set %%#_st=)
) else (
set %%#_st=
)
)

if defined Lync_st set Lync_st=Off
set OneDrive_st=Off
if defined suites (set Teams_st=Off) else (set Teams_st=)

:oe_excludeapps

cls
if not defined terminal mode 98, 32

%line%
echo:
call :dk_color %Gray% "Target edition: %targetedition%"
call :dk_color %Gray% "You can exclude the below apps from installation."
%line%
if defined suites echo:
if defined Access_st     echo [A] Access           : %Access_st%
if defined Excel_st      echo [E] Excel            : %Excel_st%
if defined OneNote_st    echo [N] OneNote          : %OneNote_st%
if defined Outlook_st    echo [O] Outlook          : %Outlook_st%
if defined PowerPoint_st echo [P] PowerPoint       : %PowerPoint_st%
if defined Project_st    echo [J] Project          : %Project_st%
if defined Publisher_st  echo [R] Publisher        : %Publisher_st%
if defined Visio_st      echo [V] Visio            : %Visio_st%
if defined Word_st       echo [W] Word             : %Word_st%
echo:
if defined Lync_st       echo [L] SkypeForBusiness : %Lync_st%
if defined OneDrive_st   echo [D] OneDrive         : %OneDrive_st%
if defined Teams_st      echo [T] Teams            : %Teams_st%
%line%
echo:
echo [1] Continue
echo [0] Go Back
%line%
echo:
call :dk_color %_Green% "Choose a menu option using your keyboard:"
choice /C:AENOPJRVWLDT10 /N
set _el=!errorlevel!
if !_el!==14 goto :oemenu
if !_el!==13 call :excludelist & goto :oe_editionchangefinal
if !_el!==12 if defined Teams_st      (if "%Teams_st%"=="Off"      (set Teams_st=ON)      else (set Teams_st=Off))
if !_el!==11 if defined OneDrive_st   (if "%OneDrive_st%"=="Off"   (set OneDrive_st=ON)   else (set OneDrive_st=Off))
if !_el!==10 if defined Lync_st       (if "%Lync_st%"=="Off"       (set Lync_st=ON)       else (set Lync_st=Off))
if !_el!==9  if defined Word_st       (if "%Word_st%"=="Off"       (set Word_st=ON)       else (set Word_st=Off))
if !_el!==8  if defined Visio_st      (if "%Visio_st%"=="Off"      (set Visio_st=ON)      else (set Visio_st=Off))
if !_el!==7  if defined Publisher_st  (if "%Publisher_st%"=="Off"  (set Publisher_st=ON)  else (set Publisher_st=Off))
if !_el!==6  if defined Project_st    (if "%Project_st%"=="Off"    (set Project_st=ON)    else (set Project_st=Off))
if !_el!==5  if defined PowerPoint_st (if "%PowerPoint_st%"=="Off" (set PowerPoint_st=ON) else (set PowerPoint_st=Off))
if !_el!==4  if defined Outlook_st    (if "%Outlook_st%"=="Off"    (set Outlook_st=ON)    else (set Outlook_st=Off))
if !_el!==3  if defined OneNote_st    (if "%OneNote_st%"=="Off"    (set OneNote_st=ON)    else (set OneNote_st=Off))
if !_el!==2  if defined Excel_st      (if "%Excel_st%"=="Off"      (set Excel_st=ON)      else (set Excel_st=Off))
if !_el!==1  if defined Access_st     (if "%Access_st%"=="Off"     (set Access_st=ON)     else (set Access_st=Off))
goto :oe_excludeapps

:excludelist

set excludelist=
for %%# in (
access
excel
onenote
outlook
powerpoint
project
publisher
visio
word
lync
onedrive
teams
) do (
if /i "!%%#_st!"=="Off" if defined excludelist (set excludelist=!excludelist!,%%#) else (set excludelist=,%%#)
)
exit /b

::========================================================================================================================================

::  Final command to change/add edition

:oe_editionchangefinal

cls
if not defined terminal mode 105, 32

::  Check for Project and Visio with unsupported language

set projvis=
set langmatched=
echo: %Project_st% %Visio_st% | find /i "ON" %nul% && set projvis=1
echo: %targetedition% | findstr /i "Project Visio" %nul% && set projvis=1

if defined projvis (
for %%# in (
ar-sa
cs-cz
da-dk
de-de
el-gr
en-us
es-es
fi-fi
fr-fr
he-il
hu-hu
it-it
ja-jp
ko-kr
nb-no
nl-nl
pl-pl
pt-br
pt-pt
ro-ro
ru-ru
sk-sk
sl-si
sv-se
tr-tr
uk-ua
zh-cn
zh-tw
) do (
if /i "%_lang%"=="%%#" set langmatched=1
)
if not defined langmatched (
%eline%
echo %_lang% language is not available for Project/Visio apps.
echo:
call :dk_color %Blue% "Install Office in the supported language for Project/Visio from the below URL."
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto :oe_goback
)
)

::  Thanks to @abbodi1406 for first discovering OfficeClickToRun.exe uses
::  Thanks to @may for the suggestion to use it to change edition with CDN as a source
::  OfficeClickToRun.exe with productstoadd method is used here to add editions
::  It uses delta updates, meaning that since it's using same installed build, it will consume very less Internet

set "c2rcommand="%_c2rExe%" platform=%_oArch% culture=%_lang% productstoadd=%targetedition%.16_%_lang%_x-none cdnbaseurl.16=http://officecdn.microsoft.com/pr/%_updch% baseurl.16=http://officecdn.microsoft.com/pr/%_updch% version.16=%_version% mediatype.16=CDN sourcetype.16=CDN deliverymechanism=%_updch% %targetedition%.excludedapps.16=groove%excludelist% flt.useteamsaddon=disabled flt.usebingaddononinstall=disabled flt.usebingaddononupdate=disabled"

if %change%==1 (
set "c2rcommand=!c2rcommand! productstoremove=AllProducts"
)

echo:
echo Running the below command, please wait...
echo:
echo %c2rcommand%
%c2rcommand%
set errorcode=%errorlevel%
timeout /t 10 %nul%

echo:
if %errorcode% EQU 0 (
call :dk_color %Gray% "Now run the Office activation option from the main menu."
) else (
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

call :oe_tempcleanup
goto :oe_goback

::========================================================================================================================================

::  Edit Office edition

:oe_editedition

cls
title  Add/Remove Apps %masver%

call :oe_chkinternet
if not defined _int (
goto :oe_goback
)

set change=0
call :ch_getinfo
cls

if not defined terminal (
mode 98, 35
)

set inpt=
set counter=0
set verified=0
set targetedition=

%line%
echo:
call :dk_color %Gray% "You can edit [add/remove apps] one of the following Office editions."
%line%
echo:

for %%A in (%_oIds%) do (
set /a counter+=1
echo [!counter!] %%A
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  Go Back
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto :oe_editedition
if "%inpt%"=="0" goto :oemenu
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto :oe_editedition

::===============

cls
if not defined terminal mode 98, 32

echo %targetedition% | findstr /i "Access Excel OneNote Outlook PowerPoint Project Publisher Skype Visio Word" %nul% && (set list=SingleApps) || (set list=Suites)
goto :oe_excludeappspre

::========================================================================================================================================

::  Remove Office editions

:oe_removeedition

title  Remove Office editions %masver%

call :ch_getinfo

cls
if not defined terminal (
mode 98, 35
)

set counter=0
for %%A in (%_oIds%) do (set /a counter+=1)

if !counter! LEQ 1 (
echo:
echo Only "%_oIds%" product is installed.
echo This option is available only when multiple products are installed.
goto :oe_goback
)

::===============

set inpt=
set counter=0
set verified=0
set targetedition=

%line%
echo:
call :dk_color %Gray% "You can uninstall one of the following Office editions."
%line%
echo:

for %%A in (%_oIds%) do (
set /a counter+=1
echo [!counter!] %%A
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  Go Back
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto :oe_removeedition
if "%inpt%"=="0" goto :oemenu
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto :oe_removeedition

::===============

cls
if not defined terminal mode 105, 32

set _lang=
echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=10) || (set _tok=9)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs\%_actconfig%\%targetedition%.16" /f "-" /k ^| findstr /i ".*16\\.*-.*"') do (
if defined _lang (set "_lang=!_lang!_%%a") else (set "_lang=_%%a")
)

set "c2rcommand="%_c2rExe%" platform=%_oArch% productstoremove=%targetedition%.16%_lang%"

echo:
echo Running the below command, please wait...
echo:
echo %c2rcommand%
%c2rcommand%

if %errorlevel% NEQ 0 (
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Help - " %_Yellow% " %mas%troubleshoot"
)

goto :oe_goback

::========================================================================================================================================

::  Change Office update channel

:oe_changeupdchnl

title  Change Office update channel %masver%
call :ch_getinfo

cls
if not defined terminal (
mode 98, 33
)

call :oe_chkinternet
if not defined _int (
goto :oe_goback
)

if %winbuild% LSS 10240 (
echo %_oIds% | findstr "2019 2021 2024" %nul% && (
%eline%
echo Installed Office editions: %_oIds%
echo Unsupported Office edition is installed on your Windows build version %winbuild%.
goto :oe_goback
)
)

::===============

set inpt=
set counter=0
set verified=0
set targetFFN=
set targetchannel=

%line%
echo:
call :dk_color %Gray% "Installed update channel: %_AudienceData%, %_version%, Client: %_clversion%"
call :dk_color %Gray% "Unsupported update channels are excluded from this list."
%line%
echo:

for %%# in (
"5440FD1F-7ECB-4221-8110-145EFAA6372F_Insider Fast [Beta]  -    Insiders::DevMain"
"64256AFE-F5D9-4F86-8936-8840A6A4F5BE_Monthly Preview      -    Insiders::CC"
"492350F6-3A01-4F97-B9C0-C7C6DDF67D60_Monthly [Current]    -  Production::CC"
"55336B82-A18D-4DD6-B5F6-9E5095C314A6_Monthly Enterprise   -  Production::MEC"
"B8F9B850-328D-4355-9145-C59439A0C4CF_Semi Annual Preview  -    Insiders::FRDC"
"7FFBC6BF-BC32-4F92-8982-F9DD17FD3114_Semi Annual          -  Production::DC"
"EA4A4090-DE26-49D7-93C1-91BFF9E53FC3_DevMain Channel      -     Dogfood::DevMain"
"B61285DD-D9F7-41F2-9757-8F61CBA4E9C8_Microsoft Elite      -   Microsoft::DevMain"
"F2E724C1-748F-4B47-8FB8-8E0D210E9208_Perpetual2019 VL     -  Production::LTSC"
"1D2D2EA6-1680-4C56-AC58-A441C8C24FF9_Microsoft2019 VL     -   Microsoft::LTSC"
"5030841D-C919-4594-8D2D-84AE4F96E58E_Perpetual2021 VL     -  Production::LTSC2021"
"86752282-5841-4120-AC80-DB03AE6B5FDB_Microsoft2021 VL     -   Microsoft::LTSC2021"
"7983BAC0-E531-40CF-BE00-FD24FE66619C_Perpetual2024 VL     -  Production::LTSC2024"
"C02D8FE6-5242-4DA8-972F-82EE55E00671_Microsoft2024 VL     -   Microsoft::LTSC2024"
) do (
for /f "tokens=1-2 delims=_" %%A in ("%%~#") do (
set supported=
if %winbuild% LSS 10240 (echo %%B | findstr /i "LTSC DevMain" %nul% || set supported=1) else (set supported=1)
if %winbuild% GEQ 10240 (
if defined ltsc19 echo %%B | find /i "2019 VL" %nul% || set supported=
if defined ltsc21 echo %%B | find /i "2021 VL" %nul% || set supported=
if defined ltsc24 echo %%B | find /i "2024 VL" %nul% || set supported=
if not defined ltscfound echo %%B | find /i "LTSC" %nul% && set supported=
)
if defined supported (
set /a counter+=1
if !counter! LSS 10 (
echo [!counter!]  %%B
) else (
echo [!counter!] %%B
)
set targetFFN!counter!=%%A
set targetchannel!counter!=%%B
)
)
)

%line%
echo:
echo [R]  Learn about update channels
echo [0]  Go back
echo:
call :dk_color %_Green% "Enter an option number using your keyboard and press Enter to confirm:"
set /p inpt=
if "%inpt%"=="" goto :oe_changeupdchnl
if "%inpt%"=="0" goto :oemenu
if /i "%inpt%"=="R" start https://learn.microsoft.com/microsoft-365-apps/updates/overview-update-channels & goto :oe_changeupdchnl
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetFFN=!targetFFN%inpt%!
set targetchannel=!targetchannel%inpt%!
if %verified%==0 goto :oe_changeupdchnl

::=======================

cls
if not defined terminal mode 105, 32

::  Get build number for the target FFN, using build number with OfficeC2RClient.exe command to trigger updates provides accurate results

set build=
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':getbuild\:.*';iex ($f[1])" %nul6%') do (set build=%%a)
echo "%build%" | find /i "16." %nul% || set build=

::  Cleanup Office update related registries, thanks to @abbodi1406
::  https://techcommunity.microsoft.com/t5/office-365-blog/how-to-manage-office-365-proplus-channels-for-it-pros/ba-p/795813
::  https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels#considerations-when-changing-channels

echo:
for /f "tokens=1 delims=-" %%A in ("%targetchannel%") do (echo Target update channel: %%A)
echo:
echo Cleaning Office update registry keys...
echo Adding new update channel to registry keys...

%nul% reg add %o16c2r_reg%\Configuration /v CDNBaseUrl /t REG_SZ /d "https://officecdn.microsoft.com/pr/%targetFFN%" /f
%nul% reg add %o16c2r_reg%\Configuration /v UpdateChannel /t REG_SZ /d "https://officecdn.microsoft.com/pr/%targetFFN%" /f
%nul% reg add %o16c2r_reg%\Configuration /v UpdateChannelChanged /t REG_SZ /d "True" /f
%nul% reg delete %o16c2r_reg%\Configuration /v UnmanagedUpdateURL /f
%nul% reg delete %o16c2r_reg%\Configuration /v UpdateUrl /f
%nul% reg delete %o16c2r_reg%\Configuration /v UpdatePath /f
%nul% reg delete %o16c2r_reg%\Configuration /v UpdateToVersion /f
%nul% reg delete %o16c2r_reg%\Updates /v UpdateToVersion /f
%nul% reg delete HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate /f
%nul% reg delete HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate /f /reg:32
%nul% reg delete HKCU\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate /f
%nul% reg delete HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate /f
%nul% reg delete HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate /f /reg:32
%nul% reg delete HKCU\Software\Policies\Microsoft\cloud\office\16.0\Common\officeupdate /f

if not defined build (
if %winbuild% GEQ 9200 call :dk_color %Gray% "Failed to detect build number for the target FFN."
set "updcommand="%_c2rCexe%" /update user"
) else (
set "updcommand="%_c2rCexe%" /update user updatetoversion=%build%"
)
echo Running the below command to trigger updates...
echo:
echo %updcommand%
%updcommand%
echo:
echo Help - %mas%troubleshoot
goto :oe_goback

::========================================================================================================================================

:oe_goback

call :oe_tempcleanup

echo:
if defined fixes (
call :dk_color %White% "Follow ALL the ABOVE blue lines.   "
call :dk_color2 %Blue% "Press [1] to Open Support Webpage " %Gray% " Press [0] to Ignore"
choice /C:10 /N
if !errorlevel!==1 (for %%# in (%fixes%) do (start %%#))
)

if defined terminal (
call :dk_color %_Yellow% "Press [0] key to go back..."
choice /c 0 /n
) else (
call :dk_color %_Yellow% "Press any key to go back..."
pause %nul1%
)
goto :oemenu

::========================================================================================================================================

:oe_tempcleanup

del /f /q %SystemRoot%\Temp\SingleApps_Volume.txt %nul%
del /f /q %SystemRoot%\Temp\SingleApps_Retail.txt %nul%
del /f /q %SystemRoot%\Temp\Suites_Volume.txt %nul%
del /f /q %SystemRoot%\Temp\Suites_Retail.txt %nul%
del /f /q %SystemRoot%\Temp\getAppIds.txt %nul%
exit /b

::========================================================================================================================================

::  Fetch required info

:ch_getinfo

set _oRoot=
set _oArch=
set _updch=
set _oIds=
set _lang=
set _cfolder=
set _version=
set _clversion=
set _AudienceData=
set _actconfig=
set _c2rXml=
set _c2rExe=
set _c2rCexe=
set _masterxml=
set ltsc19=
set ltsc21=
set ltsc24=
set ltscfound=

for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ClientFolder" %nul6%') do (set "_cfolder=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v AudienceId" %nul6%') do (set "_updch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ClientCulture" %nul6%') do (set "_lang=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v ClientVersionToReport" %nul6%') do (set "_clversion=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v AudienceData" %nul6%') do (set "_AudienceData=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\ProductReleaseIDs /v ActiveConfiguration" %nul6%') do (set "_actconfig=%%b")

echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=9) || (set _tok=8)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs\%_actconfig%" /f ".16" /k %nul6% ^| findstr /i "Retail Volume"') do (
if defined _oIds (set "_oIds=!_oIds! %%a") else (set "_oIds=%%a")
)
set _oIds=%_oIds:.16=%

set verchk=0
for /f "tokens=3 delims=." %%a in ("%_version%") do set "verchk=%%a"

if exist "%_oRoot%\Licenses16\c2rpridslicensefiles_auto.xml" set "_c2rXml=%_oRoot%\Licenses16\c2rpridslicensefiles_auto.xml"

if exist "%ProgramData%\Microsoft\ClickToRun\ProductReleases\%_actconfig%\x-none.16\MasterDescriptor.x-none.xml" (
set "_masterxml=%ProgramData%\Microsoft\ClickToRun\ProductReleases\%_actconfig%\x-none.16\MasterDescriptor.x-none.xml"
)

if exist "%_cfolder%\OfficeClickToRun.exe" (
set "_c2rExe=%_cfolder%\OfficeClickToRun.exe"
)

if exist "%_cfolder%\OfficeC2RClient.exe" (
set "_c2rCexe=%_cfolder%\OfficeC2RClient.exe"
)

set "audidata4=%_AudienceData:~-4%"

if /i "%audidata4%"=="LTSC" set ltsc19=LTSC
echo %_clversion% %_version% | findstr "16.0.103 16.0.104 16.0.105" %nul% && set ltsc19=LTSC

if /i "%audidata4%"=="2021" set ltsc21=LTSC2021
echo %_clversion% %_version% | findstr "16.0.14332" %nul% && set ltsc21=LTSC2021

if /i "%audidata4%"=="2024" set ltsc24=LTSC2024
::  LTSC 2024 build is not fixed yet

if not "%ltsc19%%ltsc21%%ltsc24%"=="" set ltscfound=1

exit /b

::========================================================================================================================================

::  Check Internet connection

:oe_chkinternet

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not "%%#"=="" set _int=1)
)

if not defined _int (
%psc% "If([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet){Exit 0}Else{Exit 1}"
if !errorlevel!==0 (set _int=1)
)

if not defined _int (
%eline%
call :dk_color %Red% "Internet is not connected."
call :dk_color %Blue% "Internet is required for this operation."
)
exit /b

::========================================================================================================================================

::  Get available build number for a FFN

:getbuild:
$Tls12 = [Enum]::ToObject([System.Net.SecurityProtocolType], 3072)
[System.Net.ServicePointManager]::SecurityProtocol = $Tls12

$FFN = $env:targetFFN
$windowsBuild = [System.Environment]::OSVersion.Version.Build

$baseUrl = "https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData?audienceFFN=$FFN"
$url = if ($windowsBuild -lt 9200) { "$baseUrl&osver=Client|6.1" } elseif ($windowsBuild -lt 10240) { "$baseUrl&osver=Client|6.3" } else { $baseUrl }

$response = if ($windowsBuild -ge 9200) { irm -Uri $url -Method Get } else { (New-Object System.Net.WebClient).DownloadString($url) }

if ($windowsBuild -lt 9200) {
    if ($response -match '"AvailableBuild"\s*:\s*"([^"]+)"') { Write-Host $matches[1] }
} else {
    Write-Host $response.AvailableBuild
}
:getbuild:

::========================================================================================================================================

::  Get available edition list from c2rpridslicensefiles_auto.xml
::  and filter the list using MasterDescriptor.x-none.xml
::  and exclude unsupported products on Windows 7/8/8.1

:getlist:
$xmlPath1 = $env:_c2rXml
$xmlPath2 = $env:_masterxml
$outputDir = $env:SystemRoot + "\Temp\"
$buildNumber = [System.Environment]::OSVersion.Version.Build
$excludedKeywords = @("2019", "2021", "2024")
$productReleaseIds = @()

if (Test-Path $xmlPath1) {
    $xml1 = New-Object -TypeName System.Xml.XmlDocument
    $xml1.Load($xmlPath1)
    foreach ($node in $xml1.SelectNodes("//ProductReleaseId")) {
        $id = $node.GetAttribute("id")
        $exclude = $false
        if ($buildNumber -lt 10240) {
            foreach ($keyword in $excludedKeywords) {
                if ($id -match $keyword) { $exclude = $true; break }
            }
        }
        if ($id -ne "CommonLicenseFiles" -and -not $exclude) { $productReleaseIds += $id }
    }
}

$categories = @{
    "Suites_Retail" = @(); "Suites_Volume" = @()
    "SingleApps_Retail" = @(); "SingleApps_Volume" = @()
}

foreach ($id in $productReleaseIds) {
    $category = if ($id -match "Retail") { "Retail" } else { "Volume" }
    $categories["SingleApps_$category"] += $id
}

if (Test-Path $xmlPath2) {
    $xml2 = New-Object -TypeName System.Xml.XmlDocument
    $xml2.Load($xmlPath2)
    foreach ($sku in $xml2.SelectNodes("//SKU")) {
        $skuId = $sku.GetAttribute("ID")
        if ($productReleaseIds -contains $skuId) {
            $appIds = $sku.SelectNodes("Apps/App") | ForEach-Object { $_.GetAttribute("id") }
            if ($appIds -contains "Excel" -and $appIds -contains "Word") {
                $category = if ($skuId -match "Retail") { "Retail" } else { "Volume" }
                $categories["Suites_$category"] += $skuId
                $categories["SingleApps_$category"] = $categories["SingleApps_$category"] | Where-Object { $_ -ne $skuId }
            }
        }
    }
}

foreach ($section in $categories.Keys) {
    $filePath = Join-Path -Path $outputDir -ChildPath "$section.txt"
    $ids = $categories[$section]
    if ($ids.Count -gt 0) { $ids | Out-File -FilePath $filePath -Encoding ASCII }
}
:getlist:

::========================================================================================================================================

::  Get App list for a specific product ID using MasterDescriptor.x-none.xml

:getappnames:
$xmlPath = $env:_masterxml
$targetSkuId = $env:targetedition
$outputDir = $env:SystemRoot + "\Temp\"
$outputFile = Join-Path -Path $outputDir -ChildPath "getAppIds.txt"
$excludeIds = @("shared", "PowerPivot", "PowerView", "MondoOnly", "OSM", "OSMUX", "Groove", "DCF")

$xml = New-Object -TypeName System.Xml.XmlDocument
$xml.Load($xmlPath)

$appIdsList = @()
$skuNodes = $xml.SelectNodes("//SKU[@ID='$targetSkuId']")

foreach ($skuNode in $skuNodes) {
    foreach ($app in $skuNode.SelectNodes("Apps/App")) {
        $appId = $app.GetAttribute("id")
        if ($excludeIds -notcontains $appId) {
            $appIdsList += $appId
        }
    }
}

if ($appIdsList.Count -gt 0) {
    $appIdsList | Out-File -FilePath $outputFile -Encoding ASCII
}
:getappnames:

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
:: Leave empty line below
