@setlocal DisableDelayedExpansion
@echo off





:: =======================================================================================================
::
::   This script is a part of 'Microsoft Activation Scripts' project.
::
::   Homepages-
::   NsaneForums: (Login Required) https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/
::   GitHub: https://github.com/massgravel/Microsoft-Activation-Scripts
::   GitLab: https://gitlab.com/massgrave/microsoft-activation-scripts
::
::   Maintained by @WindowsAddict
::
:: =======================================================================================================













::========================================================================================================================================

title Protect / Unprotect KMS38 Activation
set _elev=
if /i "%~1"=="-el" set _elev=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set "_psc=powershell -nop -ep bypass -c"
set "nul=1>nul 2>nul"
set "Red="white" "DarkRed""
set "Green="white" "DarkGreen""
set "Magenta="white" "darkmagenta""
set "Gray="white" "darkgray""
set "Black="white" "Black""
set "ELine=echo: &call :PU_color "==== ERROR ====" %Red% &echo:"
set line=__________________________________________________________________________________________________

::========================================================================================================================================

for %%i in (powershell.exe) do if "%%~$path:i"=="" (
echo: &echo ==== ERROR ==== &echo:
echo Powershell is not installed in the system.
echo Aborting...
goto PU_Done
)

::========================================================================================================================================

if %winbuild% LSS 14393 (
%ELine%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 10 / Server 1607 [14393] and later builds.
goto PU_Done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop
::  Thanks to @hearywarlot [ https://forums.mydigitallife.net/threads/.74332/ ] for the VBS method.
::  Thanks to @abbodi1406 for the powershell method and solving special characters issue in file path name.

%nul% reg query HKU\S-1-5-19 && (
  goto :Passed
  ) || (
  if defined _elev goto :E_Admin
)

set "_batf=%~f0"
set "_vbsf=%temp%\admin.vbs"
set _PSarg="""%~f0""" -el

setlocal EnableDelayedExpansion

(
echo Set strArg=WScript.Arguments.Named
echo Set strRdlproc = CreateObject^("WScript.Shell"^).Exec^("rundll32 kernel32,Sleep"^)
echo With GetObject^("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" ^& strRdlproc.ProcessId ^& "'"^)
echo With GetObject^("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" ^& .ParentProcessId ^& "'"^)
echo If InStr ^(.CommandLine, WScript.ScriptName^) ^<^> 0 Then
echo strLine = Mid^(.CommandLine, InStr^(.CommandLine , "/File:"^) + Len^(strArg^("File"^)^) + 8^)
echo End If
echo End With
echo .Terminate
echo End With
echo CreateObject^("Shell.Application"^).ShellExecute "cmd.exe", "/c " ^& chr^(34^) ^& chr^(34^) ^& strArg^("File"^) ^& chr^(34^) ^& strLine ^& chr^(34^), "", "runas", 1
)>"!_vbsf!"

(%nul% cscript //NoLogo "!_vbsf!" /File:"!_batf!" -el) && (
del /f /q "!_vbsf!"
exit /b
) || (
del /f /q "!_vbsf!"
%nul% %_psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && (
exit /b
) || (
goto :E_Admin
)
)
exit /b

:E_Admin
%ELine%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto PU_Done

:Passed

::========================================================================================================================================

mode con: cols=98 lines=30

set slp=SoftwareLicensingProduct
set sls=SoftwareLicensingService
set wApp=55c92734-d682-4d71-983e-d6ec3f16059f
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"

::  Fix for the special characters limitation in path name
::  Written by @abbodi1406

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

setlocal EnableDelayedExpansion

::  Check Windows OS name

set winos=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do if not errorlevel 1 set "winos=%%b"
if not defined winos for /f "tokens=2* delims== " %%a in ('"wmic os get caption /value" 2^>nul') do if not errorlevel 1 set "winos=%%b"

::========================================================================================================================================

cls
echo:
echo %line%
echo:
echo    [1] Protect KMS38 Activation from being overwritten by 180 days KMS Activators
echo:   
echo    [2] Undo changes
echo:   
echo    [3] Exit
echo:
echo %line%
echo:
choice /C:123 /N /M ">  Enter Your Choice [1,2,3] : "

if errorlevel 3 exit /b
if errorlevel 2 goto Undo
if errorlevel 1 goto Protect

::========================================================================================================================================

:Protect

cls

::  Check KMS client setup key

set _gvlk=
wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get ProductKeyChannel 2>nul | findstr /i GVLK 1>nul && (set _gvlk=1)

if not defined _gvlk (
%ELine%
echo System is not activated with KMS38. ^(KMS Key is not installed^)  Aborting...
goto PU_Done
)

::  Check Activation Grace Period

set gpr=
for /f "tokens=2 delims==" %%# in ('"wmic path %slp% where (ApplicationID='%wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" ') do set "gpr=%%#"

if "%gpr%" LEQ "259200" (
%ELine%
echo System is not activated with KMS38.  Aborting...
goto PU_Done
)

::  Check SKU value

set SKU=
for /f "tokens=2 delims==" %%a IN ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" 2^>nul') do if not errorlevel 1 (set osSKU=%%a)
if not defined SKU for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do if not errorlevel 1 (set osSKU=%%a)

if "%osSKU%"=="" (
%ELine%
echo SKU value was not detected properly. Aborting...
goto PU_Done
)

::  Check Windows Edition with SKU value for better accuracy

set osedition=
call :K38_CheckEdition %nul%

if "%osedition%"=="" (
%ELine%
echo OS Edition was not detected properly. Aborting...
goto PU_Done
)

::  Check Activation ID

set app=
for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%wApp%' and LicenseFamily='%osedition%' and Description like '%%KMSCLIENT%%') get ID /VALUE" 2^>nul') do set "app=%%a"

if "%app%"=="" (
%ELine%
echo Activation ID was not detected properly. Aborting...
goto PU_Done
)

::========================================================================================================================================

wmic path %sls% where __CLASS='%sls%' call ClearKeyManagementServiceMachine %nul%
wmic path %sls% where __CLASS='%sls%' call ClearKeyManagementServicePort %nul%

reg query "HKLM\%SPPk%\%wApp%" %nul% && (
%nul% call :reg_takeownership "HKLM\%SPPk%\%wApp%" FullControl Allow S-1-5-32-544
reg delete "HKLM\%SPPk%\%wApp%" /f %nul%
reg delete "HKU\S-1-5-20\%SPPk%\%wApp%" /f %nul%
)

reg query "HKLM\%SPPk%\%wApp%" %nul% && (
%ELine%
echo Registry Key was not cleared successfully. Aborting...
goto PU_Done
)

::  Set specific KMS host to Local Host
::  Thanks to @abbodi1406

set setkms_error=

wmic path %slp% where ID='%app%' call ClearKeyManagementServiceMachine %nul% || (set setkms_error=1)
wmic path %slp% where ID='%app%' call ClearKeyManagementServicePort %nul% || (set setkms_error=1)
wmic path %slp% where ID='%app%' call SetKeyManagementServiceMachine MachineName="127.0.0.2" %nul% || (set setkms_error=1)
wmic path %slp% where ID='%app%' call SetKeyManagementServicePort 1688 %nul% || (set setkms_error=1)

if defined setkms_error (
reg delete "HKLM\%SPPk%\%wApp%" /f %nul%
reg delete "HKU\S-1-5-20\%SPPk%\%wApp%" /f %nul%
%ELine%
echo Specific KMS host to Local Host was not properly applied. Aborting...
goto PU_Done
)

%nul% call :reg_takeownership "HKLM\%SPPk%\%wApp%" "SetValue, Delete" Deny S-1-5-32-544

reg delete "HKLM\%SPPk%\%wApp%" /f %nul%
reg query "HKLM\%SPPk%\%wApp%" %nul% || (
%ELine%
echo Registry Key was not protected properly. Aborting...
goto PU_Done
)

::  Check Activation Grace Period

set gpr=
for /f "tokens=2 delims==" %%# in ('"wmic path %slp% where (ApplicationID='%wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" ') do set "gpr=%%#"

if "%gpr%" LEQ "259200" (
%ELine%
echo System is not activated with KMS38.
goto PU_Done
)

cls
echo:
echo %line%
echo:
echo Specific KMS host set to 127.0.0.2 ^(Local Host^) successfully.
echo:
echo Registry item locked successfully.
echo HKLM\%SPPk%\%wApp%
echo:
call :PU_color "%winos% - KMS38 Activation is now protected." %Green%
echo:
echo Now you need to activate Office with KMS ^(If required^)
echo:
echo %line%

goto PU_Done

::========================================================================================================================================

:Undo

cls
set exist_=
reg query "HKLM\%SPPk%\%wApp%" %nul% && (set exist_=1)

if defined exist_ (
%nul% call :reg_takeownership "HKLM\%SPPk%\%wApp%" FullControl Allow S-1-5-32-544
reg delete "HKLM\%SPPk%\%wApp%" /f %nul%
reg delete "HKU\S-1-5-20\%SPPk%\%wApp%" /f %nul%
)

reg query "HKLM\%SPPk%\%wApp%" %nul% && (
%ELine%
echo Registry Key was not cleared successfully. Aborting...
goto PU_Done
)

echo:
echo %line%
echo:
if defined exist_ (
echo Registry item deleted successfully.
echo HKLM\%SPPk%\%wApp%
echo:
call :PU_color "KMS38 Activation is now unprotected [set to default]" %Green%
) else (
call :PU_color "It is already unprotected [set to default]" %Green%
)
echo:
echo %line%

::========================================================================================================================================

:PU_Done
echo:
echo Press any key to exit...
pause >nul
exit /b

::========================================================================================================================================

:PU_color

%_psc% write-host '%1' -fore '%2' -back '%3'
exit /b

::========================================================================================================================================

::  Check Windows Edition with SKU value for better accuracy

:K38_CheckEdition

for %%# in (
4-Enterprise
7-ServerStandard
8-ServerDatacenter
27-EnterpriseN
48-Professional
49-ProfessionalN
50-ServerSolution
98-CoreN
99-CoreCountrySpecific
100-CoreSingleLanguage
101-Core
110-ServerCloudStorage
120-ServerARM64
121-Education
122-EducationN
125-EnterpriseS
126-EnterpriseSN
145-ServerDatacenterACor
146-ServerStandardACor
161-ProfessionalWorkstation
162-ProfessionalWorkstationN
164-ProfessionalEducation
165-ProfessionalEducationN
168-ServerAzureCor
171-EnterpriseG
172-EnterpriseGN
175-ServerRdsh
183-CloudE
) do for /f "tokens=1,2 delims=-" %%A in ("%%#") do (
if %osSKU%==%%A set "osedition=%%B"
)
exit /b

::========================================================================================================================================

::  Reg_takeownership snippet
::  Written by @AveYo aka @BAU
::  pastebin.com/XTPt0JSC

:reg_takeownership
set "pargs=$regkey='%~1'; $p='%~2'; $a='%~3'; $u='%~4'; $o='%~5';"
%_psc% "%pargs%; $f=[io.file]::ReadAllText('!_batp!') -split ':ps_reg_own\:.*';iex ($f[1]);" & exit/b:ps_reg_own:
$dll0='[DllImport("ntdll.dll")]public static extern IntPtr RtlAdjustPrivilege(int a,bool b,bool c,ref bool d);';
$nt=Add-Type -Member $dll0 -Name Nt -PassThru; foreach($i in @(9,17,18)){$null=$nt::RtlAdjustPrivilege($i,1,0,[ref]0)}
$root=$true; if($o -eq ''){$o=$u}; $rk=$regkey -split '\\',2; $key=$rk[1];
switch -regex ($rk[0]){ '[mM]'{$HK='LocalMachine'};'[uU]'{$HK='CurrentUser'};default{$HK='ClassesRoot'}; }
$usr=0,0,0; $sec=0,0,0; $rule=0,0,0; $perm='FullControl',$p,$p; $access='Allow',$a,$a; $s=$o,$u,'S-1-5-32-544';
for($i=0;$i -le 2;$i++){ $usr[$i]=[System.Security.Principal.SecurityIdentifier]$s[$i];
$rule[$i]=[System.Security.AccessControl.RegistryAccessRule]::new($usr[$i], $perm[$i], 3, 0, $access[$i]);
$sec[$i]=[System.Security.AccessControl.RegistrySecurity]::new(); }
function Reg_TakeOwnership { param($hive, $key, $root=$false);
$reg=[Microsoft.Win32.Registry]::$hive.OpenSubKey($key,'ReadWriteSubTree','TakeOwnership'); $sec[2].SetOwner($usr[2]);
$reg.SetAccessControl($sec[2]); if($root){ $reg=$reg.OpenSubKey('','ReadWriteSubTree','ChangePermissions');
$acl=$reg.GetAccessControl(); $acl.SetAccessRuleProtection($false,$false); $acl.ResetAccessRule($rule[1]);
$reg.SetAccessControl($acl); } $sec[0].SetOwner($usr[0]); $reg.SetAccessControl($sec[0]); }
Reg_TakeOwnership $HK $key $true; if($root){ $r=[Microsoft.Win32.Registry]::$HK.OpenSubKey($key);
foreach($sk in $r.GetSubKeyNames()){try{ Reg_TakeOwnership $HK "$($key+'\\'+$sk)" $false}catch{} }}
Get-Acl "$($rk[0]+':\\'+$rk[1])" | fl:ps_reg_own:

::========================================================================================================================================