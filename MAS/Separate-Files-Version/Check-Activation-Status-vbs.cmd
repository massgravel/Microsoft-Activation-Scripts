@setlocal DisableDelayedExpansion
@echo off
@cls

:: Check-Activation-Status
:: Written by @abbodi1406
:: forums.mydigitallife.net/posts/838808

set _args=
set _args=%*
for %%A in (%_args%) do (
    if /i "%%A"=="-wow" set _rel1=1
    if /i "%%A"=="-arm" set _rel2=1
)

set "_cmdf=%~f0"

:: Handle WOW64 and ARM scenarios
if exist "%SystemRoot%\Sysnative\cmd.exe" if not defined _rel1 (
    setlocal EnableDelayedExpansion
    start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" -wow"
    exit /b
)
if exist "%SystemRoot%\SysArm32\cmd.exe" if /i %PROCESSOR_ARCHITECTURE%==AMD64 if not defined _rel2 (
    setlocal EnableDelayedExpansion
    start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" -arm"
    exit /b
)

:: Set console appearance
color 07
title Check Activation Status [vbs]

:: Update system paths for running commands
set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
    set "SysPath=%SystemRoot%\Sysnative"
    set "Path=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)

:: Check for LF line ending issues
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

:: Detect Office Ohook installation
set ohook=
for %%# in (15 16) do (
    for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
        if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" set ohook=1
    )
)

for %%# in (System SystemX86) do (
    for %%G in ("Office 15" "Office") do (
        for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
            if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" set ohook=1
        )
    )
)

:: Determine system architecture
set "_bit=64"
set "_wow=1"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "_wow=0"&set "_bit=32"

:: Initialize variables
set "_utemp=%TEMP%"
set "line2=************************************************************"
set "line3=____________________________________________________________"
set _sO16vbs=0
set _sO15vbs=0

:: Check for Office 2013/2016/2019/2021 installations
for %%v in (15 16) do (
    for %%P in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
        if exist "%%~P\Microsoft Office\Office%%v\ospp.vbs" (
            if %%v==15 set _sO15vbs=1
            if %%v==16 set _sO16vbs=1
        )
    )
)

:: Display Windows activation status
setlocal EnableDelayedExpansion
echo %line2%
echo ***                   Windows Status                     ***
echo %line2%
pushd "!_utemp!"
copy /y %SystemRoot%\System32\slmgr.vbs . >nul 2>&1
net start sppsvc /y >nul 2>&1
cscript //nologo slmgr.vbs /dli || (
    echo Error executing slmgr.vbs
    del /f /q slmgr.vbs
    popd
    goto :casVend
)
cscript //nologo slmgr.vbs /xpr
del /f /q slmgr.vbs >nul 2>&1
popd
echo %line3%

:: Display Office activation status if Ohook is installed
if defined ohook (
    echo.
    echo.
    echo %line2%
    echo ***            Office Ohook Activation Status            ***
    echo %line2%
    echo.
    powershell "write-host -back 'Black' -fore 'Yellow' 'Ohook for permanent Office activation is installed.'; write-host -back 'Black' -fore 'Yellow' 'You can ignore below Office activation status.'"
    echo.
)

:: Check Office 2016 and 2019 activation status
:casVo16
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do set "office=%%b"
if exist "!office!\ospp.vbs" (
    echo.
    echo %line2%
    if %_sO15vbs% EQU 0 (
        echo ***              Office 2016 %_bit%-bit Status               ***
    ) else (
        echo ***               Office 2013/2016 Status                ***
    )
    echo %line2%
    cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVo13
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do set "office=%%b"
if exist "!office!\ospp.vbs" (
    echo.
    echo %line2%
    if %_sO15vbs% EQU 0 (
        echo ***              Office 2016 32-bit Status               ***
    ) else (
        echo ***               Office 2013/2016 Status                ***
    )
    echo %line2%
    cscript //nologo "!office!\ospp.vbs" /dstatus
)

:: Check Office 2013 activation status
:casVo13
if %_sO16vbs% EQU 1 goto :casVo10
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do set "office=%%b"
if exist "!office!\ospp.vbs" (
    echo.
    echo %line2%
    echo ***              Office 2013 %_bit%-bit Status               ***
    echo %line2%
    cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVo10
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do set "office=%%b"
if exist "!office!\ospp.vbs" (
    echo.
    echo %line2%
    echo ***              Office 2013 32-bit Status               ***
    echo %line2%
    cscript //nologo "!office!\ospp.vbs" /dstatus
)

:: Check Office 2010 activation status
:casVo10
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do set "office=%%b"
if exist "!office!\ospp.vbs" (
    echo.
    echo %line2%
    echo ***              Office 2010 %_bit%-bit Status               ***
    echo %line2%
    cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVc16
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do set "office=%%b"
if exist "!office!\ospp.vbs" (
    echo.
    echo %line2%
    echo ***              Office 2010 32-bit Status               ***
    echo %line2%
    cscript //nologo "!office!\ospp.vbs" /dstatus
)

:: Check Office Click-to-Run activation status
:casVc16
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath >nul 2>&1 || (
    reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath >nul 2>&1 && set "install32=1"
)
if defined install32 (
    for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do set "office=%%b"
) else (
    for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do set "office=%%b"
)
if exist "!office!\ospp.vbs" (
    echo.
    echo %line2%
    echo ***         Office Click-to-Run %_bit%-bit Status         ***
    echo %line2%
    cscript //nologo "!office!\ospp.vbs" /dstatus
)

:: End script
:casVend
echo.
echo %line3%
echo ***                Script Ended                          ***
echo %line3%
exit /b
