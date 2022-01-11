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
title Extract $OEM$ Folder

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
set   "Green="42;97m""
set "Magenta="45;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set   "Green="DarkGreen" "white""
set "Magenta="Darkmagenta" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :ex_color %Red% "==== ERROR ====" &echo:"

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
goto done2
)

if not exist "%_psc%" (
%nceline%
echo Powershell is not installed in the system.
echo Aborting...
goto done2
)

::========================================================================================================================================

::  Check desktop location

set desktop=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "desktop=%%b"
if not defined desktop for /f "delims=" %%a in ('%_psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"

if not defined desktop (
%eline%
echo Desktop location was not detected, aborting...
goto done2
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"
set "_pdesk=%desktop:'=''%"

set "_ttemp=%temp%"

set "_dir=%desktop%\$OEM$\$$\Setup\Scripts"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" 1>nul && (
%eline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto done2
)

::========================================================================================================================================

mode 66, 25

if exist "!desktop!\$OEM$\" (
echo _____________________________________________________
%eline%
echo $OEM$ folder already exists on the Desktop.
echo _____________________________________________________
goto done2
)

set HWID_Activation.cmd=Activators\HWID-KMS38_Activation\HWID_Activation.cmd
set KMS38_Activation.cmd=Activators\HWID-KMS38_Activation\KMS38_Activation.cmd
set ClipUp.exe=Activators\HWID-KMS38_Activation\BIN\ClipUp.exe
set gatherosstate.exe=Activators\HWID-KMS38_Activation\BIN\gatherosstate.exe
set slc.dll=Activators\HWID-KMS38_Activation\BIN\slc.dll
set arm64_gatherosstate.exe=Activators\HWID-KMS38_Activation\BIN\arm64_gatherosstate.exe
set arm64_slc.dll=Activators\HWID-KMS38_Activation\BIN\arm64_slc.dll

set Activate.cmd=Activators\Online_KMS_Activation\Activate.cmd
set cleanosppx64.exe=Activators\Online_KMS_Activation\BIN\cleanosppx64.exe
set cleanosppx86.exe=Activators\Online_KMS_Activation\BIN\cleanosppx86.exe

pushd "!_work!"
pushd ..
pushd ..

set _nofile=
for %%# in (
%HWID_Activation.cmd%
%KMS38_Activation.cmd%
%ClipUp.exe%
%gatherosstate.exe%
%slc.dll%
%arm64_gatherosstate.exe%
%arm64_slc.dll%
%Activate.cmd%
%cleanosppx64.exe%
%cleanosppx86.exe%
) do (
if not exist "%%#" set _nofile=1
)

popd
popd
popd

if defined _nofile (
echo _____________________________________________________
%eline%
echo Some files are missing in the 'Activators' folder.
echo _____________________________________________________
goto done2
)

::========================================================================================================================================

:Menu

cls
echo:
echo:
echo:               Extract $OEM$ folder on the desktop           
echo:                  For more details use Read me
echo:                    Here ^> means Fallback to 
echo:     ________________________________________________________
echo:                                                            
echo:        [1] HWID
echo:        [2] KMS38
echo:        [3] Online KMS
echo:
echo:        [4] HWID   ^>  KMS38      
echo:        [5] HWID   ^>  Online KMS 
echo:        [6] KMS38  ^>  Online KMS 
echo:        [7] HWID   ^>  KMS38       ^> Online KMS
echo:
echo:        [8] Exit                                            
echo:     ________________________________________________________
echo:  
call :ex_color2 %_White% "      " %_Green% "Enter a menu option in the Keyboard [1,2,3,4,5,6,7,8]"
choice /C:12345678 /N
set _erl=%errorlevel%

if %_erl%==8 exit /b
if %_erl%==7 goto:hwid_kms38_kms
if %_erl%==6 goto:kms38_kms
if %_erl%==5 goto:hwid_kms
if %_erl%==4 goto:hwid_kms38
if %_erl%==3 goto:kms
if %_erl%==2 goto:kms38
if %_erl%==1 goto:hwid
goto :Menu

::========================================================================================================================================

:hwid

cls
call :prep
call :hwidprep
call :pop_d
call :export hwid_setup
call :hwidprep2

if defined _error goto errorfound
set "_oem=HWID"
goto done

:hwid_setup:
@echo off

reg query HKU\S-1-5-19 1>nul 2>nul || exit /b

start /b /wait cmd /c "%~dp0HWID_Activation.cmd" /a

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_setup:

::========================================================================================================================================

:kms38

cls
call :prep
call :kms38prep
call :pop_d
call :export kms38_setup
call :kms38prep2

if defined _error goto errorfound
set "_oem=KMS38"
goto done

:kms38_setup:
@echo off

reg query HKU\S-1-5-19 1>nul 2>nul || exit /b

start /b /wait cmd /c "%~dp0KMS38_Activation.cmd" /a

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_setup:

::========================================================================================================================================

:kms

cls
call :prep
call :kmsprep
call :pop_d
call :export kms_setup
call :kmsprep2

if defined _kerror goto errorfound
set "_oem=Online KMS"
goto done

:kms_setup:
@echo off

============================================================================

:: Change value from 1 to 0 to disable KMS Renewal And Activation Task
set Renewal_And_Activation_Task=1

:: Change value from 1 to 0 to disable KMS activation desktop context menu 
set Desktop_context_menu=1

============================================================================

reg query HKU\S-1-5-19 1>nul 2>nul || exit /b

start /b /wait cmd /c "%~dp0Activate.cmd" /wo
if %Renewal_And_Activation_Task% EQU 1 start /b /wait cmd /c "%~dp0Activate.cmd" /rat
if %Desktop_context_menu% EQU 1 start /b /wait cmd /c "%~dp0Activate.cmd" /dcm

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms_setup:

::========================================================================================================================================

:hwid_kms38

cls
call :prep
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
call :kms38prep
call :pop_d
call :export hwid_kms38_setup
call :kms38prep2
If not exist "!_dir!\HWID_Activation.cmd" set _error=1

if defined _error goto errorfound
set "_oem=HWID [Fallback to] KMS38"
goto done

:hwid_kms38_setup:
@echo off

reg query HKU\S-1-5-19 1>nul 2>nul || exit /b

set KMS38=
set osSKU=0
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "winos=%%b"
for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do set "osSKU=%%a"

:: Check Windows Server version
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" set KMS38=1

:: 171=EnterpriseG, 172=EnterpriseGN, 183=CloudE
for %%# in (171 172 183) do (if %osSKU%==%%# set KMS38=1)

:: Check Enterprise LTSC 2019/2021 N
if %osSKU%==126 (echo %winos% | findstr /i "2019 2021" 1>nul && set KMS38=1)

if defined KMS38 (
start /b /wait cmd /c "%~dp0KMS38_Activation.cmd" /a
) else (
start /b /wait cmd /c "%~dp0HWID_Activation.cmd" /a
)

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_kms38_setup:

::========================================================================================================================================

:hwid_kms

cls
call :prep
call :hwidprep
call :kmsprep
call :pop_d
call :export hwid_kms_setup
call :hwidprep2
call :kmsprep2

if defined _error goto errorfound
if defined _kerror goto errorfound
set "_oem=HWID [Fallback to] Online KMS"
goto done

:hwid_kms_setup:
@echo off

============================================================================

:: Change value from 1 to 0 to disable KMS Renewal And Activation Task
set Renewal_And_Activation_Task=1

:: Change value from 1 to 0 to disable KMS activation desktop context menu 
set Desktop_context_menu=1

============================================================================

reg query HKU\S-1-5-19 1>nul 2>nul || exit /b

set KMS=
set osSKU=0
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "winos=%%b"
for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do set "osSKU=%%a"

:: Check Windows 10 install
if %winbuild% LSS 10240 set KMS=1

:: Check Windows Server version
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" set KMS=1

:: 171=EnterpriseG, 172=EnterpriseGN, 183=CloudE
for %%# in (171 172 183) do (if %osSKU%==%%# set KMS=1)

:: Check Enterprise LTSC 2019/2021 N
if %osSKU%==126 (echo %winos% | findstr /i "2019 2021" 1>nul && set KMS=1)

if defined KMS (
start /b /wait cmd /c "%~dp0Activate.cmd" /wo
) else (
start /b /wait cmd /c "%~dp0HWID_Activation.cmd" /a
start /b /wait cmd /c "%~dp0Activate.cmd" /o
)

if %Renewal_And_Activation_Task% EQU 1 start /b /wait cmd /c "%~dp0Activate.cmd" /rat
if %Desktop_context_menu% EQU 1 start /b /wait cmd /c "%~dp0Activate.cmd" /dcm

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_kms_setup:

::========================================================================================================================================

:kms38_kms

cls
call :prep
call :kms38prep
call :kmsprep
call :pop_d
call :export kms38_kms_setup
call :kms38prep2
call :kmsprep2

if defined _error goto errorfound
if defined _kerror goto errorfound
set "_oem=KMS38 [Fallback to] Online KMS"
goto done

:kms38_kms_setup:
@echo off

============================================================================

:: Change value from 1 to 0 to disable KMS Renewal And Activation Task
set Renewal_And_Activation_Task=1

:: Change value from 1 to 0 to disable KMS activation desktop context menu 
set Desktop_context_menu=1

============================================================================

reg query HKU\S-1-5-19 1>nul 2>nul || exit /b

start /b /wait cmd /c "%~dp0KMS38_Activation.cmd" /a

start /b /wait cmd /c "%~dp0Activate.cmd" /wo
if %Renewal_And_Activation_Task% EQU 1 start /b /wait cmd /c "%~dp0Activate.cmd" /rat
if %Desktop_context_menu% EQU 1 start /b /wait cmd /c "%~dp0Activate.cmd" /dcm

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_kms_setup:

::========================================================================================================================================

:hwid_kms38_kms

cls
call :prep
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
call :kms38prep
call :kmsprep
call :pop_d
call :export hwid_kms38_kms_setup

call :kms38prep2
call :kmsprep2
If not exist "!_dir!\HWID_Activation.cmd" set _error=1

if defined _error goto errorfound
if defined _kerror goto errorfound
set "_oem=HWID [Fallback to] KMS38 [Fallback to] Online KMS"
goto done

:hwid_kms38_kms_setup:
@echo off

============================================================================

:: Change value from 1 to 0 to disable KMS Renewal And Activation Task
set Renewal_And_Activation_Task=1

:: Change value from 1 to 0 to disable KMS activation desktop context menu 
set Desktop_context_menu=1

============================================================================

reg query HKU\S-1-5-19 1>nul 2>nul || exit /b

set KMS=
set KMS38=
set osSKU=0
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do set "winos=%%b"
for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do set "osSKU=%%a"

:: Check Windows 10 install
if %winbuild% LSS 10240 set KMS=1

:: Check Windows Server version
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" set KMS38=1

:: 171=EnterpriseG, 172=EnterpriseGN, 183=CloudE
for %%# in (171 172 183) do (if %osSKU%==%%# set KMS38=1)

:: Check Enterprise LTSC 2019/2021 N
if %osSKU%==126 (echo %winos% | findstr /i "2019 2021" 1>nul && set KMS38=1)

if defined KMS38 set KMS=1

if defined KMS (
if defined KMS38 (start /b /wait cmd /c "%~dp0KMS38_Activation.cmd" /a)
start /b /wait cmd /c "%~dp0Activate.cmd" /wo
) else (
start /b /wait cmd /c "%~dp0HWID_Activation.cmd" /a
start /b /wait cmd /c "%~dp0Activate.cmd" /o
)

if %Renewal_And_Activation_Task% EQU 1 start /b /wait cmd /c "%~dp0Activate.cmd" /rat
if %Desktop_context_menu% EQU 1 start /b /wait cmd /c "%~dp0Activate.cmd" /dcm

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_kms38_kms_setup:

::========================================================================================================================================

:errorfound

%eline%
echo $OEM$ Folder was not created successfully...
goto :done2

:done

echo _______________________________________________________
echo:
call :ex_color %Magenta% "%_oem%"
call :ex_color %Green% "$OEM$ folder is successfully created on the Desktop."
echo _______________________________________________________

:done2

echo:
call :ex_color %_Yellow% "Press any key to exit..."
pause >nul
exit /b

::========================================================================================================================================

:prep

pushd "!desktop!"
md "!desktop!\$OEM$\$$\Setup\Scripts\BIN"
pushd "!_work!"
pushd ..
pushd ..
exit /b

:hwidprep

copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
copy /y /b "%gatherosstate.exe%" "!_dir!\BIN\gatherosstate.exe" %nul%
copy /y /b "%slc.dll%" "!_dir!\BIN\slc.dll" %nul%
copy /y /b "%arm64_gatherosstate.exe%" "!_dir!\BIN\arm64_gatherosstate.exe" %nul%
copy /y /b "%arm64_slc.dll%" "!_dir!\BIN\arm64_slc.dll" %nul%
exit /b

:kms38prep

copy /y /b "%KMS38_Activation.cmd%" "!_dir!\KMS38_Activation.cmd" %nul%
copy /y /b "%ClipUp.exe%" "!_dir!\BIN\ClipUp.exe" %nul%
copy /y /b "%gatherosstate.exe%" "!_dir!\BIN\gatherosstate.exe" %nul%
copy /y /b "%slc.dll%" "!_dir!\BIN\slc.dll" %nul%
copy /y /b "%arm64_gatherosstate.exe%" "!_dir!\BIN\arm64_gatherosstate.exe" %nul%
copy /y /b "%arm64_slc.dll%" "!_dir!\BIN\arm64_slc.dll" %nul%
exit /b

:kmsprep

copy /y /b "%Activate.cmd%" "!_dir!\Activate.cmd" %nul%
copy /y /b "%cleanosppx64.exe%" "!_dir!\BIN\cleanosppx64.exe" %nul%
copy /y /b "%cleanosppx86.exe%" "!_dir!\BIN\cleanosppx86.exe" %nul%
exit /b

:hwidprep2

set _error=
pushd "!_dir!\"

for %%# in (
HWID_Activation.cmd
BIN\gatherosstate.exe
BIN\slc.dll
BIN\arm64_gatherosstate.exe
BIN\arm64_slc.dll
SetupComplete.cmd
) do (
if not exist "%%#" set _error=1
)
popd
exit /b

:kms38prep2

set _error=
pushd "!_dir!\"

for %%# in (
KMS38_Activation.cmd
BIN\ClipUp.exe
BIN\gatherosstate.exe
BIN\slc.dll
BIN\arm64_gatherosstate.exe
BIN\arm64_slc.dll
SetupComplete.cmd
) do (
if not exist "%%#" set _error=1
)
popd
exit /b

:kmsprep2

set _kerror=
pushd "!_dir!\"

for %%# in (
Activate.cmd
BIN\cleanosppx64.exe
BIN\cleanosppx86.exe
SetupComplete.cmd
) do (
if not exist "%%#" set _kerror=1
)
popd
exit /b

:pop_d

popd
popd
popd
popd
exit /b

::========================================================================================================================================

::  Extract the text from batch script without character and file encoding issue

:export

%nul% %_psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('!_pdesk!\$OEM$\$$\Setup\Scripts\SetupComplete.cmd',$f[1].Trim(),[System.Text.Encoding]::ASCII);"
exit /b

::========================================================================================================================================

:ex_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
if not exist %_psc% (echo %~3) else (%_psc% write-host -back '%1' -fore '%2' '%3')
)
exit /b

:ex_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
if not exist %_psc% (echo %~3%~6) else (%_psc% write-host -back '%1' -fore '%2' '%3' -NoNewline; write-host -back '%4' -fore '%5' '%6')
)
exit /b

::========================================================================================================================================