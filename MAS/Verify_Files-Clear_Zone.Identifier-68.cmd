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
mode con cols=65 lines=12
title Verify Files ^& Clear Zone.Identifier

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
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set   "Green="DarkGreen" "white""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :color %Red% "==== ERROR ====" &echo:"

::========================================================================================================================================

::  Self verification (made sure that script won't crash, if it's in Unix-LF format) 

set "_hash="
for /f "skip=1 tokens=* delims=" %%G in ('certutil -hashfile "%~f0" SHA1^|findstr /i /v CertUtil') do set "_hash=%%G"
set "_hash=%_hash: =%"
set "_hash=%_hash:~-2%"
set "_fina=%~n0"
set "_fina=%_fina:~-2%"

if /i not "%_hash%"=="%_fina%" (
%nceline%
echo File SHA-1 verification failed.
echo Make sure that file is not modified / renamed.
echo:
echo Press any key to exit...
pause >nul
exit /b
)

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
goto done
)

if not exist %_psc% (
%nceline%
echo Powershell is not installed in the system.
echo Aborting...
goto done
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

for %%# in (All-In-One-Version,Separate-Files-Version) do (
if not exist "!_work!\%%#" (
%eline%
echo [%%~#] folder not found in the current directory.
goto done
)
)

::========================================================================================================================================

set fileM=0
set hashM=0

for %%# in (
31e13b31812ea4fb3073c0ef4a0527490be5d9da+Separate-Files-Version\Extras\Change_W10_11_Edition.cmd
db4c68bba8a9c9cccfe76d0f1753a2cd922b94f2+Separate-Files-Version\Extras\Extract_OEM_Folder\Extract_OEM_Folder.cmd
d047f7b3bb205e8eb9412af11402e32a31b7906f+Separate-Files-Version\Extras\Extract_OEM_Folder\ReadMe.html
44a364ac2d6fad784aef03361fd64460fbe7357a+Separate-Files-Version\Extras\Install_W10_11_HWID_Key.cmd
89054be4d565ee7f9defa4159b1997d1bdf96d56+Separate-Files-Version\Extras\_Homepage.html
236916e59019d183a55ace4f892016d5cd2194bd+ReadMe.html
9cc32357cb46a078779e51c14402ec594acf611b+Separate-Files-Version\Activators\Online_KMS_Activation\Activate.cmd
fabb5a0fc1e6a372219711152291339af36ed0b5+Separate-Files-Version\Activators\HWID-KMS38_Activation\BIN\gatherosstate.exe
d30a0e4e5911d3ca705617d17225372731c770e2+Separate-Files-Version\Activators\Online_KMS_Activation\BIN\cleanosppx64.exe
da1afd97d92dd6026e7095ee7442a2144f78ed0b+Separate-Files-Version\Activators\HWID-KMS38_Activation\BIN\slc.dll
286f3bb552b6368a347ca74cb7407026624c4eb3+Separate-Files-Version\Activators\HWID-KMS38_Activation\BIN\_Info.html
39ed8659e7ca16aaccb86def94ce6cec4c847dd6+Separate-Files-Version\Activators\Online_KMS_Activation\BIN\cleanosppx86.exe
836ae2f742e8dbf54762f4ecc2468c68eecff6d9+Separate-Files-Version\Activators\Online_KMS_Activation\BIN\_Info.html
6cd44e7186b396016bd97802a7e28d659ac94e78+Separate-Files-Version\Activators\HWID-KMS38_Activation\HWID_Activation.cmd
81d25225805b80a5d32906f32b5aa67d00b24b0c+Separate-Files-Version\Activators\Online_KMS_Activation\ReadMe.html
5bf7ebbb3c4de976476925053b3a8e6dc689cff5+Separate-Files-Version\Extras\Activation_Troubleshoot.cmd
f4d1fa0d085bc17561416946ccbdaf419570b8f9+Separate-Files-Version\Activators\HWID-KMS38_Activation\KMS38_Activation.cmd
06ae500b740d90148a951bd7b40ddc8f9ec0a109+Separate-Files-Version\Activators\HWID-KMS38_Activation\ReadMe_HWID.html
1f90667b15471d9a74ee3a2839a8b795b623fc86+Separate-Files-Version\Activators\HWID-KMS38_Activation\ReadMe_KMS38.html
4d11828cac7728e25f6e2d1e76553d779d4a33ff+All-In-One-Version\MAS_1.5_AIO_CRC32_21D20776.cmd
f19d8a19f6a684e87e2421d185d83af3f5c24a70+Separate-Files-Version\Activators\Activations_Summary.html
c00cd43aa95e8221b8ee6a9e758eb7b128139997+Separate-Files-Version\Activators\Check-Activation-Status-vbs.cmd
27ead0b8d2b8346e55ab54bb682dc3c5afd1ed59+Separate-Files-Version\Activators\Check-Activation-Status-wmi.cmd
023d88e8e0a125f5d85ee2d999b512c4886aab29+Separate-Files-Version\Activators\HWID-KMS38_Activation\BIN\arm64_slc.dll
7e449ae5549a0d93cf65f4a1bb2aa7d1dc090d2d+Separate-Files-Version\Activators\HWID-KMS38_Activation\BIN\arm64_gatherosstate.exe
48d928b1bec25a56fe896c430c2c034b7866aa7a+Separate-Files-Version\Activators\HWID-KMS38_Activation\BIN\ClipUp.exe
) do for /f "tokens=1,2* delims=+" %%A in ("%%#") do (
if not exist "!_work!\%%B" (
set fileM=1
set hashM=1
) else (
set "_hash="
for /f "skip=1 tokens=* delims=" %%G in ('certutil -hashfile "!_work!\%%B" SHA1^|findstr /i /v CertUtil') do (
set "_hash=%%G"
set "_hash=!_hash: =!"
if /i not "%%A"=="!_hash!" set hashM=1
)
)
)

::========================================================================================================================================

cls
echo:

set n=0
set mn=0
set cn=27

for /f %%a in ('2^>nul dir "!_work!\" /a-d/b/-o/-p/s^|find /v /c ""') do set n=%%a

if %fileM%==0 (
echo Checking Files                          [Passed]
) else (
call :color %Red% "Checking Files                          [Files Are Missing]"
)

if %n% EQU %cn% echo Checking Total Number Of Files          [Passed] [%cn%]

if %n% GTR %cn% (
set /a "mn=%n%-%cn%"
call :color %Red% "Checking Total Number Of Files          [!mn! - Extra Files Found]"
)

if %n% LSS %cn% (
set /a "mn=%cn%-%n%"
call :color %Red% "Checking Total Number Of Files          [!mn! - Less Files Found]"
)

if %hashM%==0 (
echo Verifying Files SHA-1 Hash              [Passed]
) else (
call :color %Red% "Verifying Files SHA-1 Hash              [Mismatch Found]"
)

::========================================================================================================================================

::  Clear NTFS alternate data streams (Zone.Identifier)
::  winitor.com/pdf/NtfsAlternateDataStreams.pdf
::  docs.microsoft.com/en-us/archive/blogs/askcore/alternate-data-streams-in-ntfs

set zone=0
pushd "!_work!\"
dir /s /r | find ":$DATA" 1>nul && set zone=1

if %zone%==0 (
echo Clearing Zone.Identifier From Files     [Already clean]
) else (
if %winbuild% LSS 9200 (
%nul% %_psc% "iex(([io.file]::ReadAllText('!_batp!')-split':unblock\:.*')[1])"
) else (
%nul% %_psc% "& {gci -recurse | unblock-file}"
)
dir /s /r | find ":$DATA" 1>nul
if [!errorlevel!]==[0] (
call :color %Red% "Clearing Zone.Identifier From Files     [Failed]
) else (
echo Clearing Zone.Identifier From Files     [Passed]
)
)
popd

::========================================================================================================================================

:done

echo:
if not exist "%_psc%" (
echo Press any key to exit...
) else (
call :color %_Yellow% "Press any key to exit..."
)
pause >nul
exit /b

::========================================================================================================================================

:color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
%_psc% write-host -back '%1' -fore '%2' '%3'
)
exit /b

::========================================================================================================================================

::  andyarismendi.blogspot.com/2012/02/unblocking-files-with-powershell.html
::  github.com/ellisgeek/Scripts_Windows/blob/master/Powershell/Helper%20Functions/Unblock-File.ps1
::  Written by @Andy Arismendi

::  This code to unblock files is used when PowerShell 2.0 is installed (Windows 7 and equivalent).
::  With PowerShell 3.0 (Windows 8 and equivalent) and above, script uses one liner cmdlet 'Unblock-File'

:unblock:

function Unblock-File {
#Requires -Version 2.0
    [cmdletbinding(DefaultParameterSetName = "ByName",
                   SupportsShouldProcess = $True)]
    param (
        [parameter(Mandatory = $true,
                   ParameterSetName = "ByName",
                   Position = 0)]
        [string]
        $Path,
        [parameter(Mandatory = $true,
                   ParameterSetName = "ByInput",
                   ValueFromPipeline = $true)]
        $InputObject
    )
    begin {
        Add-Type -Namespace Win32 -Name PInvoke -MemberDefinition @"
        //  msdn.microsoft.com/en-us/library/windows/desktop/aa363915(v=vs.85).aspx
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeleteFile(string name);
        public static int Win32DeleteFile(string filePath) {
            bool is_gone = DeleteFile(filePath); return Marshal.GetLastWin32Error();}
 
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetFileAttributes(string lpFileName);
        public static bool Win32FileExists(string filePath) {return GetFileAttributes(filePath) != -1;}
"@
    }
    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName'  {
                $input_paths = Resolve-Path -Path $Path | ? { [IO.File]::Exists($_.Path) } | `
                Select -Exp Path
            }
            'ByInput' {
                if ($InputObject -is [System.IO.FileInfo]) {
                    $input_paths = $InputObject.FullName
                }
            }
        }
        $input_paths | % {
            if ([Win32.PInvoke]::Win32FileExists($_ + ':Zone.Identifier')) {
                if ($PSCmdlet.ShouldProcess($_)) {
                    $result_code = [Win32.PInvoke]::Win32DeleteFile($_ + ':Zone.Identifier')
                    if ([Win32.PInvoke]::Win32FileExists($_ + ':Zone.Identifier')) {
                        Write-Error ("Failed to unblock '{0}' the Win32 return code is '{1}'." -f `
                                     $_, $result_code)
                    }
                }
            }
        }
    }
}
gci -recurse | Unblock-File

:unblock:

::========================================================================================================================================