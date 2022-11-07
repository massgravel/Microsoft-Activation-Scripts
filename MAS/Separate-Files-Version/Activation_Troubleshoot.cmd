@setlocal DisableDelayedExpansion
@echo off



::============================================================================
::
::   This script is a part of 'Microsoft Activation Scripts' (MAS) project.
::
::   Homepage: massgrave.dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================




::========================================================================================================================================

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="r1" set r1=1
if /i "%%#"=="r2" set r2=1
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

::  Set Path variable, it helps if it is misconfigured in the system

set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "PATH=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%PATH%"
)

::  Check LF line ending

pushd "%~dp0"
>nul findstr /rxc:".*" "%~nx0"
if not %errorlevel%==0 (
echo:
echo Error: This is not a correct file. It has LF line ending issue.
echo:
ping 127.0.0.1 -n 6 > nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title  Activation Troubleshoot

set _elev=
if /i "%~1"=="-el" set _elev=1

set winbuild=1
set "nul=>nul 2>&1"
set psc=powershell.exe
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 2>nul | find /i "0x0" 1>nul && (set _NCS=0)

call :_colorprep

set cbs_log=%SystemRoot%\logs\cbs\cbs.log
set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :_color %Red% "==== ERROR ====" &echo:"
set "line=_________________________________________________________________________________________________"
if %~z0 GEQ 200000 (set "_exitmsg=Go back") else (set "_exitmsg=Exit")

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected.
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
goto at_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto at_done
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%temp%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" 1>nul && (
if /i not "!_work!"=="!_ttemp!" (
%nceline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto at_done
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

>nul fltmc || (
if not defined _elev %nul% %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
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
if not defined desktop for /f "delims=" %%a in ('%psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"

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
echo:       _______________________________________________________________
echo:                                                   
call :_color2 %_White% "             [1] " %_Green% "Help"
echo:             ___________________________________________________
echo:                                                                      
echo:             [2] Dism RestoreHealth
echo:             [3] SFC Scannow
echo:
echo:             [4] Rebuild Licensing Tokens
echo:             [5] Rebuild ClipSVC Licences
echo:             [6] Clear Office vNext Licences
echo:             ___________________________________________________
echo:                                                                      
echo:             [7] Rebuild WMI Repository
echo:             [8] Fix: Issues Caused By Gaming Spoofers
echo:             [9] Fix: Issues Caused By KB971033 In Windows 7
echo:             [G] Fix: Office Is Not Genuine Banner
echo:             [E] Export Event Viewer Logs
echo:             ___________________________________________________
echo:
echo:             [0] %_exitmsg%
echo:       _______________________________________________________________
echo:          
call :_color2 %_White% "            " %_Green% "Enter a menu option in the Keyboard :"
choice /C:123456789GE0 /N
set _erl=%errorlevel%

if %_erl%==12 exit /b
if %_erl%==11 goto:exportevtlogs
if %_erl%==10 start https://massgrave.dev/office-license-is-not-genuine &goto at_menu
if %_erl%==9 goto:fixwindows7
if %_erl%==8 goto:fixspoofer
if %_erl%==7 goto:rewmi
if %_erl%==6 goto:clearvnext
if %_erl%==5 goto:reclipsvc
if %_erl%==4 goto:retokens
if %_erl%==3 goto:sfcscan
if %_erl%==2 goto:dism_rest
if %_erl%==1 start https://massgrave.dev/troubleshoot.html &goto at_menu
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
for %%a in (dns.msftncsi.com) do (
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
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
mode 110, 30
echo:

call :_stopservice TrustedInstaller
del /s /f /q "%SystemRoot%\logs\cbs\*.*"

set _time=
for /f %%a in ('%psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
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
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
echo:

call :_stopservice TrustedInstaller
del /s /f /q "%SystemRoot%\logs\cbs\*.*"

set _time=
for /f %%a in ('%psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
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

:clearvnext

cls
mode 98, 30
title  Clear Office vNext License

echo:
echo %line%
echo:    
echo      This options will clear Office vNext ^(subscription^) license
echo:
echo      You need to use this option when,
echo          - KMS option is not activating office due to existing subscription license
echo          - KMS option activated Office but Office activation page is not showing activated
echo:   
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
echo:
echo %line%
echo:
call :_color %Magenta% "Clearing Office vNext License"
echo:

setlocal DisableDelayedExpansion
set "_Local=%LocalAppData%"
setlocal EnableDelayedExpansion

attrib -R "!ProgramData!\Microsoft\Office\Licenses" %nul%
attrib -R "!_Local!\Microsoft\Office\Licenses" %nul%
rd /s /q "!ProgramData!\Microsoft\Office\Licenses\" %nul%
rd /s /q "!_Local!\Microsoft\Office\Licenses\" %nul%

if exist "!ProgramData!\Microsoft\Office\Licenses\" (
echo Failed To Delete - !ProgramData!\Microsoft\Office\Licenses\
) else (
echo Deleted Folder - !ProgramData!\Microsoft\Office\Licenses\
)

if exist "!_Local!\Microsoft\Office\Licenses\" (
echo Failed To Delete - !_Local!\Microsoft\Office\Licenses\
) else (
echo Deleted Folder - !_Local!\Microsoft\Office\Licenses\
)

echo:
for %%# in (
HKCU\Software\Microsoft\Office\16.0\Common\Licensing
HKCU\Software\Microsoft\Office\16.0\Registration
) do (
reg query %%# %nul% && (
reg delete %%# /f %nul% && (
echo Deleted Registry - %%#
) || (
echo Failed to Delete - %%#
)
) || (
echo Deleted Registry - %%#
)
)

goto :at_back

::========================================================================================================================================

:retokens

cls
mode con cols=115 lines=32
%nul% %psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"
title  Rebuild Licensing Tokens ^(SPP ^+ OSPP)

echo:
echo %line%
echo:   
echo      Notes:
echo:
echo       - It helps in troubleshooting activation issues.
echo:
call :_color2 %_White% "      - " %Magenta% "This option will,"
call :_color2 %_White% "        " %Magenta% "- Deactivate Windows and Office, you will need to reactivate"
call :_color2 %_White% "        " %Magenta% "- Uninstall Office licenses and keys"
call :_color2 %_White% "        " %Magenta% "- Clear SPP-OSPP data.dat, tokens.dat, cache.dat"
call :_color2 %_White% "        " %Magenta% "- Trigger the repair option for Office"
echo:
call :_color2 %_White% "      - " %Red% "Apply it only when it is necessary."
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu


cls
:cleanspptoken
echo:
echo %line%
echo:
call :_color %Magenta% "Rebuilding SPP Licensing Tokens"
echo:

call :scandat check

if not defined token (
call :_color %Red% "tokens.dat file not found."
) else (
echo tokens.dat file: [%token%]
)

echo:
echo Stopping sppsvc service...
call :_stopservice sppsvc

echo:
call :scandat delete
call :scandat check

if defined token (
echo:
call :_color %Red% "Failed to delete .dat files."
echo:
)

echo:
echo Reinstalling System Licenses [slmgr /rilc]...
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if %errorlevel% NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if %errorlevel% EQU 0 (
echo [Successful]
) else (
call :_color %Red% "[Failed]"
)

call :scandat check

echo:
if not defined token (
call :_color %Red% "Failed to rebuilt tokens.dat file."
) else (
call :_color %Green% "tokens.dat file was rebuilt successfully."
)

::========================================================================================================================================

::  Rebuild OSPP Tokens

echo:
echo %line%
echo:

sc qc osppsvc %nul% || (
echo:
call :_color %Magenta% "OSPP based Office is not installed"
call :_color %Magenta% "Skipping rebuilding OSPP tokens"
goto :repairoffice
)

call :_color %Magenta% "Rebuilding OSPP Licensing Tokens"
echo:

call :scandatospp check

if not defined token (
call :_color %Red% "tokens.dat file not found."
) else (
echo tokens.dat file: [%token%]
)

echo:
echo Stopping osppsvc service...
call :_stopservice osppsvc

echo:
call :scandatospp delete
call :scandatospp check

if defined token (
echo:
call :_color %Red% "Failed to delete .dat files."
echo:
)

echo:
echo Starting osppsvc service to generate tokens.dat
call :_startservice osppsvc
call :scandatospp check
if not defined token (
call :_stopservice osppsvc
call :_startservice osppsvc
timeout /t 3 %nul%
)

call :scandatospp check

echo:
if not defined token (
call :_color %Red% "Failed to rebuilt tokens.dat file."
) else (
call :_color %Green% "tokens.dat file was rebuilt successfully."
)

::========================================================================================================================================

:repairoffice

echo:
echo %line%
echo:
call :_color %Magenta% "Repairing Office Licenses"
echo:

for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b

if /i "%arch%"=="ARM64" (
echo:
echo ARM64 Windows Found.
echo You need to use repair option in Windows settings for Office.
echo:
start ms-settings:appsfeatures
goto :repairend
)

if /i "%arch%"=="x86" (
set arch=X86
) else (
set arch=X64
)

for %%# in (68 86) do (
for %%A in (msi14 msi15 msi16 c2r14 c2r15 c2r16) do (set %%A_%%#=&set %%Arepair%%#=)
)

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

%nul% reg query %_68%\14.0\Common\InstallRoot /v Path  && (set "msi14_68=Office 14.0 MSI x86/x64"  & set "msi14repair68=%systemdrive%\Program Files\Common Files\microsoft shared\OFFICE14\Office Setup Controller\Setup.exe")
%nul% reg query %_86%\14.0\Common\InstallRoot /v Path  && (set "msi14_86=Office 14.0 MSI x86"      & set "msi14repair86=%systemdrive%\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\Office Setup Controller\Setup.exe")
%nul% reg query %_68%\15.0\Common\InstallRoot /v Path  && (set "msi15_68=Office 15.0 MSI x86/x64"  & set "msi15repair68=%systemdrive%\Program Files\Common Files\microsoft shared\OFFICE15\Office Setup Controller\Setup.exe")
%nul% reg query %_86%\15.0\Common\InstallRoot /v Path  && (set "msi15_86=Office 15.0 MSI x86"      & set "msi15repair86=%systemdrive%\Program Files (x86)\Common Files\Microsoft Shared\OFFICE15\Office Setup Controller\Setup.exe")
%nul% reg query %_68%\16.0\Common\InstallRoot /v Path  && (set "msi16_68=Office 16.0 MSI x86/x64"  & set "msi16repair68=%systemdrive%\Program Files\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\Setup.exe")
%nul% reg query %_86%\16.0\Common\InstallRoot /v Path  && (set "msi16_86=Office 16.0 MSI x86"      & set "msi16repair86=%systemdrive%\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\Setup.exe")
%nul% reg query %_68%\14.0\CVH /f Click2run /k         && (set "c2r14_68=Office 14.0 C2R x86/x64"  & set "c2r14repair68=")
%nul% reg query %_86%\14.0\CVH /f Click2run /k         && (set "c2r14_86=Office 14.0 C2R x86"      & set "c2r14repair86=")
%nul% reg query %_68%\15.0\ClickToRun /v InstallPath   && (set "c2r15_68=Office 15.0 C2R x86/x64"  & set "c2r15repair68=%systemdrive%\Program Files\Microsoft Office 15\Client%arch%\integratedoffice.exe")
%nul% reg query %_86%\15.0\ClickToRun /v InstallPath   && (set "c2r15_86=Office 15.0 C2R x86"      & set "c2r15repair86=%systemdrive%\Program Files\Microsoft Office 15\Client%arch%\integratedoffice.exe")
%nul% reg query %_68%\ClickToRun /v InstallPath        && (set "c2r16_68=Office 16.0 C2R x86/x64"  & set "c2r16repair68=%systemdrive%\Program Files\Microsoft Office 15\Client%arch%\OfficeClickToRun.exe")
%nul% reg query %_86%\ClickToRun /v InstallPath        && (set "c2r16_86=Office 16.0 C2R x86"      & set "c2r16repair86=%systemdrive%\Program Files\Microsoft Office 15\Client%arch%\OfficeClickToRun.exe")

set uwp16=
if %winbuild% GEQ 10240 (
dir /b "%ProgramFiles%\WindowsApps\Microsoft.Office.Desktop*" %nul% && set uwp16=Office 16.0 UWP
dir /b "%ProgramW6432%\WindowsApps\Microsoft.Office.Desktop*" %nul% && set uwp16=Office 16.0 UWP
dir /b "%ProgramFiles(x86)%\WindowsApps\Microsoft.Office.Desktop*" %nul% && set uwp16=Office 16.0 UWP
%psc% "Get-AppxPackage -name "Microsoft.Office.Desktop"" | find /i "Office" 1>nul && set uwp16=Office 16.0 UWP
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
echo Multiple office versions found.
echo It's recommended to install only one version of office.
echo ________________________________________________________________
echo:
)

if %counter% EQU 0 (
echo:
echo Installed Office is not found.
goto :repairend
echo:
) else (
echo:
call :_color %_Yellow% "A Window will popup, in that Window you need to select [Quick] Repair Option..."
call :_color %_Yellow% "Press any key to continue..."
echo:
pause >nul
)

if defined uwp16 (
echo:
echo Note: Skipping repair for Office 16.0 UWP. 
echo       You need to use reset option in Windows settings for it.
echo ________________________________________________________________
echo:
start ms-settings:appsfeatures
)

set c2r14=
if defined c2r14_68 set c2r14=1
if defined c2r14_86 set c2r14=1

if defined c2r14 (
echo:
echo Note: Skipping repair for Office 14.0 C2R 
echo       You need to use Repair option in Windows settings for it.
echo ________________________________________________________________
echo:
start appwiz.cpl
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
call :_color %Green% "Finished"
goto :at_back

::========================================================================================================================================

:reclipsvc

cls
mode 98, 30
title  Rebuild ClipSVC Licences

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
echo       - Rebuilding ClipSVC Licences helps in troubleshooting HWID-KMS38 activation issues.
echo:
echo       - Do not run this option unless you are having issues in HWID-KMS38 activation.
echo:
echo       - System restart is recommended after applying it.
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
echo:

echo Stopping ClipSVC service...
call :_stopservice ClipSVC
timeout /t 2 %nul%

echo:
echo Applying the command to Clean ClipSVC Licences...
echo rundll32 clipc.dll,ClipCleanUpState

rundll32 clipc.dll,ClipCleanUpState

if %winbuild% LEQ 10240 (
call :_color %Green% "[Successful]"
) else (
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat" (
call :_color %Red% "[Failed]"
) else (
call :_color %Green% "[Successful]"
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
call :_color %Red% "[Failed]"
echo Restart the system, that will delete this registry key automatically.
) || (
call :_color %Green% "[Successful]"
)

::   Clear HWID token related registry to fix activation incase if there is any corruption

echo:
echo Deleting a IdentityCRL Registry Key...
echo [%_ident%]
reg delete "%_ident%" /f %nul%
reg query "%_ident%" %nul% && (
call :_color %Red% "[Failed]"
) || (
call :_color %Green% "[Successful]"
)

echo:
echo Restarting [ClipSVC wlidsvc LicenseManager sppsvc] services...
for %%# in (ClipSVC wlidsvc LicenseManager sppsvc) do (net stop %%# /y %nul% & net start %%# /y %nul%)
goto :at_back

::========================================================================================================================================

:fixspoofer

cls
mode con cols=115 lines=32
%nul% %psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"
title  Fix: Issues Caused By Gaming Spoofers

%psc% $ExecutionContext.SessionState.LanguageMode 2>nul | find /i "Full" 1>nul || (
%eline%
echo Powershell is not responding properly. Aborting."
goto :at_back
)

echo:
echo %line%
echo:
echo      Notes:
echo:
echo       - Gaming unban/spoofers/cleaners often cause Windows activation issues.
echo:
call :_color2 %_White% "      - " %Red% "Apply this fix ONLY if you have used these things."
echo:
echo       - This option will fix files and registry permissions and rebuild licensing tokens.
echo:
echo       - System restart is recommended after applying it.
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
echo:
echo Fixing registry and files permissions...
call :fixpermissions %nul%
goto :cleanspptoken

:fixpermissions

::  Thanks to skidaim for the fix

takeown /F %windir%\System32\sppsvc.exe
icacls %windir%\System32 /grant administrators:F /T
icacls %windir%\System32\spp /grant administrators:F /T

::  I know it's bad but people have messed up system32 permissions, that's why I don't recommend to run this unless users have messed up systems

%psc% $acl = Get-Acl 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'; $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc','FullControl','ContainerInherit, ObjectInherit','None','Allow'); $acl.SetAccessRule($rule); Set-Acl -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform' -AclObject $acl
%psc% $acl = Get-Acl 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SPP'; $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc','FullControl','ContainerInherit, ObjectInherit','None','Allow'); $acl.SetAccessRule($rule); Set-Acl -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SPP' -AclObject $acl
%psc% $acl = Get-Acl 'HKLM:\SYSTEM\CurrentControlSet\Services\SPPSVC'; $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc','FullControl','ContainerInherit, ObjectInherit','None','Allow'); $acl.SetAccessRule($rule); Set-Acl -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\SPPSVC' -AclObject $acl
%psc% $acl = Get-Acl 'HKLM:\SYSTEM\WPA'; $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc','FullControl','ContainerInherit, ObjectInherit','None','Allow'); $acl.SetAccessRule($rule); Set-Acl -Path 'HKLM:\SYSTEM\WPA' -AclObject $acl
%psc% $acl = Get-Acl '%windir%\System32'; $rule = New-Object System.Security.AccessControl.FileSystemAccessRule ('NT Service\sppsvc','FullControl','ContainerInherit, ObjectInherit','None','Allow'); $acl.SetAccessRule($rule); Set-Acl -Path '%windir%\System32' -AclObject $acl
%psc% $acl = Get-Acl '%windir%\System32\spp'; $rule = New-Object System.Security.AccessControl.FileSystemAccessRule ('NT Service\sppsvc','FullControl','ContainerInherit, ObjectInherit','None','Allow'); $acl.SetAccessRule($rule); Set-Acl -Path '%windir%\System32\spp' -AclObject $acl
exit /b

::========================================================================================================================================

:fixwindows7

cls
mode 98, 30
title  Fix: Issues Caused By KB971033 In Windows 7

if %winbuild% GEQ 9200 (
%eline%
echo Unsupported OS version Detected.
echo This option is supported only for Windows 7 and it's Server equivalent.
goto :at_back
)

echo:
echo %line%
echo:
echo      Notes:
echo:
echo       - This option fixes issues caused by Update KB971033 in Windows 7.
echo         https://support.microsoft.com/en-us/help/4487266
echo:
echo %line%
echo:
choice /C:01 /N /M ">    [1] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

cls
echo:

echo Checking Update KB971033...
dism /online /get-packages | find /i "Microsoft-Windows-Security-WindowsActivationTechnologies-package~31bf3856ad364e35~amd64~~7.1.7600.16395" 1>nul && (
echo [Found]
echo Uninstalling it...
) || (
echo [Not Found]
)

wusa /uninstall /quiet /norestart /kb:971033

echo:
echo Applying Fixes...
echo:

net stop sppuinotify /y
sc config sppuinotify start= disabled
net stop sppsvc /y
del %windir%\system32\7B296FB0-376B-497e-B012-9C450E1B7327-5P-0.C7483456-A289-439d-8115-601632D005A0 /ah
del %windir%\system32\7B296FB0-376B-497e-B012-9C450E1B7327-5P-1.C7483456-A289-439d-8115-601632D005A0 /ah
del %windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat
del %windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\cache\cache.dat
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
sc config sppuinotify start= demand

goto :at_back

::========================================================================================================================================

:rewmi

cls
mode 98, 30
title  Rebuild WMI Repository

::  https://techcommunity.microsoft.com/t5/ask-the-performance-team/wmi-repository-corruption-or-not/ba-p/375484

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo WMI rebuild is not recommended on Windows Server. Aborting...
goto :at_back
)

echo:
echo Initializing...

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" set _wmic=1

set error=
if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" 2>nul | find /i "computersystem" 1>nul
if %errorlevel% NEQ 0 set error=1
winmgmt /verifyrepository %nul%
if %errorlevel% NEQ 0 set error=1

cls
echo:
echo %line%
echo:
if defined error (
echo      WMI Status - [Not Responding] %_wmic%
) else (
call :_color %_Green% "     WMI Status - [Working]"
)
echo:
echo      Notes:
echo:
call :_color2 %_White% "      - " %Magenta% "WMI rebuild can cause some 3rd party apps to not work until reinstall."
echo:       
call :_color2 %_White% "      - " %Red% "Apply this fix ONLY if WMI is not working."
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] Continue [0] Go back : "
if %errorlevel%==1 goto at_menu

::  Below fixes are taken from https://kb.acronis.com/content/62731

cls
echo:

sc query Winmgmt %nul% || (
%eline%
echo Winmgmt service is not installed. Aborting...
goto :at_back
)

echo Disabling Winmgmt service...
sc config Winmgmt start= disabled %nul%
if %errorlevel% EQU 0 (
call :_color %Green% "[Successful]"
) else (
call :_color %Red% "[Failed] Aborting..."
goto :wmifixend
)

echo:
echo Stopping Winmgmt service...
call :_stopservice Winmgmt
call :_stopservice Winmgmt
sc query Winmgmt | find /i "1  STOPPED" %nul% && (
call :_color %Green% "[Successful]"
) || (
call :_color %Red% "[Failed] Aborting..."
goto :wmifixend
)

echo:
echo Deleting WMI repository...
if exist "%windir%\System32\wbem\repository\" rmdir /s /q "%windir%\System32\wbem\repository\" %nul%
if exist "%windir%\System32\wbem\repository\" (
call :_color %Red% "[Failed]"
) else (
call :_color %Green% "[Successful]"
)

echo:
echo Enabling Winmgmt service...
sc config Winmgmt start= auto %nul%
if %errorlevel% EQU 0 (
call :_color %Green% "[Successful]"
) else (
call :_color %Red% "[Failed]"
)

echo:
echo Checking WMI...
if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" 2>nul | find /i "computersystem" 1>nul
if %errorlevel% NEQ 0 (
call :_color %Red% "[Not Responding]"
) else (
call :_color %Green% "[Working]"
)

goto :at_back

:wmifixend

echo:
echo Enabling Winmgmt service...
sc config Winmgmt start= auto %nul%
if %errorlevel% EQU 0 (
call :_color %Green% "[Successful]"
) else (
call :_color %Red% "[Failed]"
)

goto :at_back

::========================================================================================================================================

:exportevtlogs

cls
mode con cols=125 lines=32
%nul% %psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=500;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"
title  Export Event Viewer Logs

set tdir=%SystemRoot%\Temp\_EventLogs
if exist %tdir%\. rd /s /q %tdir%\ %nul%
if exist %tdir%\ (
%eline%
echo Failed to delete below folder. Aborting...
echo %tdir%\
goto :at_back
)

md %tdir%\

echo:
echo Creating archive file of Event logs...

set _time=
for /f %%a in ('%psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
%nul% robocopy %SystemRoot%\System32\winevt\Logs\ %tdir%\

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
pushd "%tdir%\"
for /f "tokens=* delims=" %%D in ('dir /a:-D/b/s "%tdir%\"') do (
 echo/"%%~fD"  /inf=no;>>%ddf%
)
makecab /F %ddf% /D DiskDirectory1="" /D CabinetNameTemplate=%tdir%\Logs.cab
del /q /f %ddf%
popd

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%
copy /y /b "%tdir%\Logs.cab" "!desktop!\AT_Logs\EventLogs_%_time%.cab" %nul%
if exist %tdir%\. rd /s /q %tdir%\ %nul%

echo:
if exist "!desktop!\AT_Logs\EventLogs_%_time%.cab" (
call :_color %Green% "[Successful]"
echo EventLogs_%_time%.cab created inside AT_Logs folder on the dekstop.
) else (
call :_color %Red% "[Failed]"
)

goto :at_back

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
echo Press any key to %_exitmsg%...
pause >nul
exit /b

::========================================================================================================================================

:_stopservice

for %%# in (%1) do (
sc query %%# | find /i "STOPPED" %nul% || net stop %%# /y %nul%
sc query %%# | find /i "STOPPED" %nul% || sc stop %%# %nul%
)
exit /b

:_startservice

for %%# in (%1) do (
sc query %%# | find /i "RUNNING" %nul% || net start %%# /y %nul%
sc query %%# | find /i "RUNNING" %nul% || sc start %%# %nul%
)
exit /b

::========================================================================================================================================

:scandat

set token=
for %%# in (
%Systemdrive%\Windows\System32\spp\store_test\2.0\
%Systemdrive%\Windows\System32\spp\store\
%Systemdrive%\Windows\System32\spp\store\2.0\
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

setlocal
set "TMP=%SystemRoot%\Temp"
set "TEMP=%SystemRoot%\Temp"
%psc% "$f=[io.file]::ReadAllText('!_batp!') -split ':regown\:.*';iex ($f[1]);"
endlocal
exit /b

::  Below code takes ownership of a volatile registry key and deletes it
::  HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState

::  Thanks to Remko Weijnen for the code and thanks to abbodi1406 for the help
::  remkoweijnen.nl/blog/2012/01/16/take-ownership-of-a-registry-key-in-powershell/

:regown:
$definition = @"
using System;
using System.Runtime.InteropServices;
namespace Win32Api
{
    public class NtDll
    {
        [DllImport("ntdll.dll", EntryPoint="RtlAdjustPrivilege")]
        public static extern int RtlAdjustPrivilege(int Privilege, bool Enable, bool CurrentThread, ref bool Enabled);
    }
}
"@

Add-Type -TypeDefinition $definition -PassThru | Out-Null
[Win32Api.NtDll]::RtlAdjustPrivilege(9, $true, $false, [ref]$false) | Out-Null

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