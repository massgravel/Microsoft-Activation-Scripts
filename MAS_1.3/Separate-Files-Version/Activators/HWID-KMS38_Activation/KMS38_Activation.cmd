@setlocal DisableDelayedExpansion
@echo off


:: For unattended mode, run the script with /u parameter.



::=========================================================================================================
:    Credits:
::=========================================================================================================
::
::   @mspaintmsi   Original co-authors of HWID/KMS38 Activation without KMS or predecessor install/upgrade.
::      and        Created various methods for HWID/KMS38 Activation
::   *Anonymous    https://www.nsaneforums.com/topic/316668--/?do=findComment&comment=1497887
::                 https://github.com/massgravel/MASSGRAVE
::                 https://gitlab.com/massgrave/massgrave
::
::   @vyvojar      Original slshim (slc.dll)
::                 https://github.com/vyvojar/slshim/releases
::
::---------------------------------------------------------------------------------------------------------
::
::   HWID/KMS38 methods Suggestions and improvements:-
::  
::   @sponpa       New ideas for the HWID/KM38 Generation
::                 https://www.nsaneforums.com/topic/316668--/page/21/?tab=comments#comment-1431257
::
::   @leitek8      Improvements for the slc.dll
::                 https://www.nsaneforums.com/topic/316668--/page/22/?tab=comments#comment-1438005
::
::---------------------------------------------------------------------------------------------------------
::
::   Kind Help:-
::
::   Thanks for having my back and answering all of my queries. (In no particular order)
::   
::   @AveYo aka @BAU, @sponpa, @mspaintmsi @RPO, @leitek8, @mxman2k, @Yen, @abbodi1406
::
::   @BorrowedWifi for providing support in fixing English grammar errors in the Read Me.
::   @Chibi ANUBIS and @smashed for testing scripts for ARM64 system.
::
::   Special thanks to @abbodi1406 for providing the great help.
::
::---------------------------------------------------------------------------------------------------------
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
::   P.S. I (@WindowsAddict) did not help in the development of HWID/KMS38 Activation in any way, I only 
::   manage batch script tool which is based on the above mentioned original co-authors activation methods.
::
::=========================================================================================================









::========================================================================================================================================

cls
title  KMS38 Activation
set Unattended=
set _args=
set _elev=
set "_arg1=%~1"
if not defined _arg1 goto :K38_NoProgArgs
set "_args=%~1"
set "_arg2=%~2"
if defined _arg2 set "_args=%~1 %~2"
for %%A in (%_args%) do (
if /i "%%A"=="-el" set _elev=1
if /i "%%A"=="/u" set Unattended=1)
:K38_NoProgArgs
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set "_psc=powershell -nop -ep bypass -c"
set "nul=1>nul 2>nul"
set "Red="white" "DarkRed""
set "Green="white" "DarkGreen""
set "Magenta="white" "darkmagenta""
set "Gray="white" "darkgray""
set "Black="white" "Black""
set "ELine=echo: &call :K38_color "==== ERROR ====" %Red% &echo:"
set slp=SoftwareLicensingProduct
set sls=SoftwareLicensingService
set wApp=55c92734-d682-4d71-983e-d6ec3f16059f

::========================================================================================================================================

for %%i in (powershell.exe) do if "%%~$path:i"=="" (
echo: &echo ==== ERROR ==== &echo:
echo Powershell is not installed in the system.
echo Aborting...
goto K38_Done
)

::========================================================================================================================================

if %winbuild% LSS 14393 (
%ELine%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 10 / Server - 1607 [14393] and later builds.
goto K38_Done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop
::  Thanks to @hearywarlot [ https://forums.mydigitallife.net/threads/.74332/ ] for the VBS method.
::  Thanks to @abbodi1406 for the powershell method and solving special characters issue in file path name.

%nul% reg query HKU\S-1-5-19 && (
  goto :K38_Passed
  ) || (
  if defined _elev goto :K38_E_Admin
)

set "_batf=%~f0"
set "_vbsf=%temp%\admin.vbs"
set _PSarg="""%~f0""" -el
if defined _args set _PSarg="""%~f0""" -el """%_args%"""

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

(%nul% cscript //NoLogo "!_vbsf!" /File:"!_batf!" -el "!_args!") && (
del /f /q "!_vbsf!"
exit /b
) || (
del /f /q "!_vbsf!"
%nul% %_psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && (
exit /b
) || (
goto :K38_E_Admin
)
)
exit /b

:K38_E_Admin
%ELine%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto K38_Done

:K38_Passed

::========================================================================================================================================

::  Fix for the special characters limitation in path name
::  Written by @abbodi1406

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

mode con: cols=102 lines=30

::  Check Windows OS name

set winos=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do if not errorlevel 1 set "winos=%%b"
if not defined winos for /f "tokens=2* delims== " %%a in ('"wmic os get caption /value" 2^>nul') do if not errorlevel 1 set "winos=%%b"

call :K38_CheckPermAct
if defined PermAct (

echo ___________________________________________________________________________________________
echo:
call :K38_color1 "     " %Black% &call :K38_color "Checking: %winos% is Permanently Activated." %Green%
call :K38_color1 "     " %Black% &call :K38_color "Activation is not required." %Gray%
echo ___________________________________________________________________________________________
echo:
if defined Unattended goto K38_Done

echo      Press [1] or [2] button in Keyboard :
echo ___________________________________________
echo:
choice /C:12 /N /M ">    [1] Activate [2] Exit : "

if errorlevel 2 exit /b
if errorlevel 1 Goto K38_Continue
)

:K38_Continue
cls

::========================================================================================================================================

cd /d "!_work!"
pushd "!_work!"

if not exist "!_work!\BIN\" (
%ELine%
echo 'BIN' Folder does not exist in current directory.
echo It's supposed to have files required for the Activation.
goto K38_Done
)

::========================================================================================================================================

echo %winos%| findstr /I Evaluation >nul && set Eval=1||set Eval=
if defined Eval (
%ELine%
echo [%winos% ^| %winbuild%] KMS38 Activation is Not Supported.
echo %winos%| findstr /I Server >nul && (
echo Server Evaluation cannot be activated. Convert it to full Server OS.
) || (
echo Evaluation Editions cannot be activated. Install full Windows OS.
)
goto K38_Done
)

::========================================================================================================================================

::  Check SKU value

set SKU=
for /f "tokens=2 delims==" %%a IN ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" 2^>nul') do if not errorlevel 1 (set osSKU=%%a)
if not defined SKU for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do if not errorlevel 1 (set osSKU=%%a)

if "%osSKU%"=="" (
%ELine%
echo SKU value was not detected properly. Aborting...
goto K38_Done
)

::  Check Windows Edition with SKU value for better accuracy

set osedition=
call :K38_CheckEdition %nul%

if "%osedition%"=="" (
%ELine%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%] KMS38 Activation is Not Supported.
goto K38_Done
)

::  Check Activation ID

set app=
for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%wApp%' and LicenseFamily='%osedition%' and Description like '%%KMSCLIENT%%') get ID /VALUE" 2^>nul') do set "app=%%a"

::  Check Windows Architecture 

set arch=
reg Query "HKLM\Hardware\Description\System\CentralProcessor\0" | find /i "x86" > nul && set arch=x86|| set arch=x64
wmic os get osarchitecture | find /i "ARM" > nul && set arch=ARM64|| echo %PROCESSOR_ARCHITECTURE% | find /i "ARM" > nul && set arch=ARM64

if "%arch%"=="ARM64" call :K38_check ARM64_gatherosstate.exe ARM64_slc.dll
if not "%arch%"=="ARM64" call :K38_check gatherosstate.exe slc.dll
if defined _miss goto K38_Done

::========================================================================================================================================

cls
if "%app%"=="" (
%ELine%
echo [%winos% ^| %winbuild%] KMS38 Activation is Not Supported.
goto K38_Done
)

set key=
call :%app% %nul%

if "%key%"=="" (
%ELine%
echo [%winos% ^| %winbuild% ^| %app%] KMS38 Activation is Not Supported.
goto K38_Done
)

:: clipup.exe does not exist in server cor and acor editions.

set A_Cor=
echo %osedition%| findstr /I /B Server >nul && if not exist "%systemroot%\System32\clipup.exe" set A_Cor=1

if defined A_Cor (
call :K38_check clipup.exe
if defined _miss goto K38_Done
)

::========================================================================================================================================

cls
echo Checking OS Info                        [%winos% ^| %winbuild% ^| %arch%]

echo:
set _1=ClipSVC
set _3=sppsvc

for %%# in (%_1% %_3%) do call :K38_ServiceCheck %%#

set "CLecho=Checking %_1%                        [Service Status -%Cl_state%] [Startup Type -%Cl_start_type%]"
set "specho=Checking %_3%                         [Service Status -%sp_state%] [Startup Type -%sp_start_type%]"

if not "%Cl_start_type%"=="Demand"       (call :K38_color "%CLecho%" %Red% & set Clst_e=1) else (echo %CLecho%)
if not "%sp_start_type%"=="Delayed-Auto" (call :K38_color "%specho%" %Red% & set spst_e=1) else (echo %specho%)

echo:
if defined Clst_e (sc config %_1% start= Demand %nul%       && set Clst_s=%_1%-Demand || set Clst_u=%_1%-Demand )
if defined spst_e (sc config %_3% start= Delayed-Auto %nul% && set spst_s=%_3%-Delayed-Auto || set spst_u=%_3%-Delayed-Auto )

for %%# in (Clst_s,spst_s) do if defined %%# set st_s=1
if defined st_s (echo Changing services Startup Type to       [ %Clst_s%%spst_s%] [Successful])

for %%# in (Clst_u,spst_u) do if defined %%# set st_u=1
if defined st_u (call :K38_color "Error in changing Startup Type to       [ %Clst_u%%spst_u%]" %Red%)

if not "%Cl_state%"=="Running" (%_psc% start-service %_1% %nul% && set Cl_s=%_1% || set Cl_u=%_1% )
if not "%sp_state%"=="Running" (%_psc% start-service %_3% %nul% && set sp_s=%_3% || set sp_u=%_3% )

for %%# in (Cl_s,sp_s) do if defined %%# set s_s=1
if defined s_s (echo Starting services                       [ %Cl_s%%sp_s%] [Successful])

for %%# in (Cl_u,sp_u) do if defined %%# set s_u=1
if defined s_u (call :K38_color "Error in starting services              [ %Cl_u%%sp_u%]" %Red%)

::========================================================================================================================================

::  Thanks to @abbodi1406 for the WMI methods

echo:
set _channel=
set _Keyexist=
set _partial=

wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get ProductKeyChannel 2>nul | findstr /i GVLK 1>nul && (set _Keyexist=1)

if defined _Keyexist (
for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get PartialProductKey /value 2^>nul') do set "_partial=%%#"
call echo Checking Installed Product Key          [Volume:GVLK] [Partial Key -%%_partial%%]
)

if not defined _Keyexist (
set "InsKey=Installing KMS Client Setup Key        "
wmic path %sls% where __CLASS='%sls%' call InstallProductKey ProductKey="%key%" %nul% && (
call echo %%InsKey%% [%key%] [Successful]
) || (
call :K38_color "%%InsKey%% [%key%] [Unsuccessful]" %Red%
)
)

wmic path %sls% where __CLASS='%sls%' call RefreshLicenseStatus %nul%

::========================================================================================================================================

::  Set specific KMS host to Local Host
::  Thanks to @abbodi1406

::  By doing this, global KMS IP can not replace KMS38 activation but can be used with Office and other Windows Editions.

echo:
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
reg delete "HKLM\%SPPk%\%wApp%" /f %nul%
reg delete "HKU\S-1-5-20\%SPPk%\%wApp%" /f %nul%

set setkms_error=
set "setkms_=Setting Specific KMS Host to           "

wmic path %slp% where ID='%app%' call ClearKeyManagementServiceMachine %nul% || (set setkms_error=1)
wmic path %slp% where ID='%app%' call ClearKeyManagementServicePort %nul% || (set setkms_error=1)
wmic path %slp% where ID='%app%' call SetKeyManagementServiceMachine MachineName="127.0.0.2" %nul% || (set setkms_error=1)
wmic path %slp% where ID='%app%' call SetKeyManagementServicePort 1688 %nul% || (set setkms_error=1)

if not defined setkms_error (
echo %setkms_% [LocalHost 127.0.0.2] [Successful]
) else (
call :K38_color "%setkms_% [LocalHost 127.0.0.2] [Unsuccessful]" %Red%
)

::========================================================================================================================================

:: Files are copied to temp to generate ticket to avoid possible issues in case the path contains special character or non English names.

echo:
set "temp_=%SystemRoot%\Temp\_Ticket_Work"
if exist "%temp_%\" @RD /S /Q "%temp_%\" %nul%
md "%temp_%\" %nul%

cd /d "!_work!\BIN\"

set ARM64_file=
if "%arch%"=="ARM64" set ARM64_file=ARM64_

set "A_Cor_cl=%systemroot%\System32\clipup.exe"
copy /y /b "%ARM64_file%gatherosstate.exe" "%temp_%\gatherosstate.exe" %nul%
copy /y /b "%ARM64_file%slc.dll" "%temp_%\slc.dll" %nul%
if defined A_Cor (copy /y /b "ClipUp.exe" "%A_Cor_cl%" %nul%)

set cfailed=
if not exist "%temp_%\gatherosstate.exe" set cfailed=1
if not exist "%temp_%\slc.dll" set cfailed=1

set "copyfiles=Copying Required Files to Temp         "
set "copyclipup=Copying clipup.exe File to             "

if defined cfailed (
call :K38_color "%copyfiles% [%SystemRoot%\Temp\_Ticket_Work\] [Unsuccessful] Aborting..." %Red%
goto K38_Act_Cont
) else (
echo %copyfiles% [%SystemRoot%\Temp\_Ticket_Work\] [Successful]
)

if defined A_Cor (
if not exist "%A_Cor_cl%" call :K38_color "%copyclipup% [%systemroot%\System32\] [Unsuccessful] Aborting..." %Red% & goto K38_Act_Cont
if exist "%A_Cor_cl%" echo %copyclipup% [%systemroot%\System32\] [Successful]
)

cd /d "%temp_%\"
attrib -R -A -S -H *.*

::========================================================================================================================================

:: Multiple attempts to generate the ticket because in some cases, one attempt is not enough.

set "_noxml=if not exist "%temp_%\GenuineTicket.xml""

start /wait "" "%temp_%/gatherosstate.exe" %nul%
%_noxml% timeout /t 3 %nul%
%_noxml% call "%temp_%/gatherosstate.exe" %nul%
%_noxml% timeout /t 3 %nul%
%_noxml% "%temp_%/gatherosstate.exe" %nul%
%_noxml% timeout /t 3 %nul%

set "GenTicket=Generating GenuineTicket.xml           "
%_noxml% (
call :K38_color "%GenTicket% [Unsuccessful] Aborting..." %Red%
goto K38_Act_Cont
) else (
echo %GenTicket% [Successful]
)

:: clipup -v -o -altto <Ticket path> method to apply ticket was not used to avoid the certain issues in case the username have spaces or non English names.

set "InsTicket=Installing GenuineTicket.xml           "
set "TDir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if exist "%TDir%\*.xml" del /f /q "%TDir%\*.xml" %nul%
copy /y /b "%temp_%\GenuineTicket.xml" "%TDir%\GenuineTicket.xml" %nul%

if not exist "%TDir%\GenuineTicket.xml" (
call :K38_color "Failed to copy Ticket to [%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket\] Aborting..." %Red%
goto K38_Act_Cont
)

set "_xmlexist=if exist "%TDir%\GenuineTicket.xml""

%_psc% Restart-Service ClipSVC %nul%
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

%_xmlexist% (
%_psc% stop-Service ClipSVC %nul%
%_psc% start-Service ClipSVC %nul%
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%
)

set fallback_=
%_xmlexist% (
set fallback_=1
%nul% clipup -v -o
%_xmlexist% timeout /t 2 %nul%
)

%_xmlexist% (
call :K38_color "%InsTicket% [Unsuccessful] Aborting..." %Red%
if exist "%TDir%\*.xml" del /f /q "%TDir%\*.xml" %nul%
goto K38_Act_Cont
) else (
if not defined fallback_ echo %InsTicket% [Successful]
if defined fallback_ call :K38_color "%InsTicket% [Successful] [Fallback method: clipup -v -o]" %Red%
)

::========================================================================================================================================

echo:
echo Activating...
echo:

call :K38_CheckEXPtime
if %gpr% GTR 259200 (
call :K38_Actinfo
goto K38_Act_Cont
)

::========================================================================================================================================

::  Clear 180 Days KMS Activation lock without full Rearm and Restart

set RearmF=
set "Rearm=Applying SKU-APP ID Rearm              "
wmic path %sls% where __CLASS='%sls%' call ReArmApp ApplicationId="%wApp%" %nul% || (set RearmF=1)
wmic path %slp% where ID='%app%' call ReArmsku %nul% || (set RearmF=1)

if defined RearmF (
call :K38_color "%Rearm% [Unsuccessful]" %Red%
) else (
echo %Rearm% [Successful]
)

::========================================================================================================================================

echo:
call :K38_CheckEXPtime
if %gpr% GTR 259200 (
call :K38_Actinfo
goto K38_Act_Cont
)

%_psc% Restart-Service sppsvc %nul%
wmic path %sls% where __CLASS='%sls%' call RefreshLicenseStatus %nul%

call :K38_CheckEXPtime
if %gpr% GTR 259200 (
call :K38_Actinfo
goto K38_Act_Cont
)

call :K38_color "Activation Failed." %Red%
call :K38_color "Try the Troubleshoot Guide listed in the ReadMe." %Magenta%

:K38_Act_Cont

::  clipup.exe does not exist in server cor and acor editions by default, it was copied there with this script.

echo:
cd /d "!_work!\"
if exist "%temp_%\" @RD /S /Q "%temp_%\" %nul%
if defined A_Cor (if exist "%A_Cor_cl%" del /f /q "%A_Cor_cl%" %nul%)

set "delFiles=Cleaning Temp Files                    "
set "delclipup=Deleting copied clipup.exe file        "

if exist "%temp_%\" (
call :K38_color "%delFiles% [Unsuccessful]" %Red%
) else (
echo %delFiles% [Successful]
)

if defined A_Cor (
if exist "%A_Cor_cl%" call :K38_color "%delclipup% [Unsuccessful]" %Red%
if not exist "%A_Cor_cl%" echo %delclipup% [Successful]
)

goto K38_Done

::========================================================================================================================================

::  Echo all the missing files.
::  Written by @abbodi1406 (MDL)

:K38_check

for %%# in (%1 %2) do (if not exist "!_work!\BIN\%%#" (if defined _miss (set "_miss=!_miss! %%#") else (set "_miss=%%#")))
if defined _miss (
%ELine%
echo Following required file^(s^) is missing in 'BIN' folder. Aborting...
echo:
echo !_miss!
)
exit /b

::========================================================================================================================================

:K38_ServiceCheck

::  Detect Service status and start type
::  Written by @RPO

for /f "tokens=1,3 delims=: " %%a in ('sc query %1') do (if /i %%a==state set "state=%%b")
for /f "tokens=1-4 delims=: " %%a in ('sc qc %1') do (if /i %%a==start_type set "start_type=%%c %%d")

if /i "%state%"=="STOPPED" set state=Stopped
if /i "%state%"=="RUNNING" set state=Running

if /i "%start_type%"=="auto_start (delayed)" set start_type=Delayed-Auto
if /i "%start_type%"=="auto_start "          set start_type=Auto
if /i "%start_type%"=="demand_start "        set start_type=Demand
if /i "%start_type%"=="disabled "            set start_type=Disabled

for %%i in (%*) do (
if /i "%%i"=="%_3%" set "sp_start_type=%start_type%" & set "sp_state=%state%"
if /i "%%i"=="%_1%" set "Cl_start_type=%start_type%" & set "Cl_state=%state%"
)
exit /b

::========================================================================================================================================

:K38_CheckPermAct

::  Check Windows Permanent Activation status
::  Written by @abbodi1406

wmic path %slp% where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name 2>nul | findstr /i "Windows" 1>nul && set PermAct=1||set PermAct=
exit /b

::========================================================================================================================================

::  Check Expiration date with powershell
::  Written by @abbodi1406

:K38_Actinfo

for /f "tokens=* delims=" %%# in ('%_psc% "$([DateTime]::Now.addMinutes(%gpr%)).ToString('yyyy-MM-dd HH:mm:ss')" 2^>nul') do set "_xpr=%%#"
call :K38_color "%winos% is activated till %_xpr%" %Green%
exit /b

::========================================================================================================================================

:K38_CheckEXPtime

::  Check Activation remaining time
::  Written by @abbodi1406

for /f "tokens=2 delims==" %%# in ('"wmic path %slp% where (ApplicationID='%wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" ') do set "gpr=%%#"
exit /b

::========================================================================================================================================

:K38_color

%_psc% write-host '%1' -fore '%2' -back '%3'
exit /b

:K38_color1

%_psc% write-host '%1' -fore '%2' -back '%3'  -NoNewline
exit /b

::========================================================================================================================================

:K38_Done

echo:
if defined Unattended (
echo Exiting in 3 seconds...
if %winbuild% LSS 7600 (ping -n 3 127.0.0.1 > nul) else (timeout /t 3)
exit /b
)
echo Press any key to exit...
pause >nul
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

:: Generic Volume License Key (GVLK) List

:: Windows 10 [RS5]
:32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee
set "key=M7XTQ-FN8P6-TTKYV-9D4CC-J462D" &:: Enterprise LTSC 2019
exit /b

:7103a333-b8c8-49cc-93ce-d37c09687f92
set "key=92NFX-8DJQP-P6BBQ-THF9C-7CG2H" &:: Enterprise LTSC 2019 N
exit /b

:ec868e65-fadf-4759-b23e-93fe37f2cc29
set "key=CPWHC-NT2C7-VYW78-DHDB2-PG3GK" &:: Enterprise for Virtual Desktops
exit /b

:0df4f814-3f57-4b8b-9a9d-fddadcd69fac
set "key=NBTWJ-3DR69-3C4V8-C26MC-GQ9M6" &:: Lean
exit /b

:: Windows 10 [RS3]
:82bbc092-bc50-4e16-8e18-b74fc486aec3
set "key=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J" &:: Pro Workstation
exit /b

:4b1571d3-bafb-4b40-8087-a961be2caf65
set "key=9FNHH-K3HBT-3W4TD-6383H-6XYWF" &:: Pro Workstation N
exit /b

:e4db50ea-bda1-4566-b047-0ca50abc6f07
set "key=7NBT4-WGBQX-MP4H7-QXFF8-YP3KX" &:: Enterprise Remote Server
exit /b

:: Windows 10 [RS2]
:e0b2d383-d112-413f-8a80-97f373a5820c
set "key=YYVX9-NTFWV-6MDM3-9PT4T-4M68B" &:: Enterprise G
exit /b

:e38454fb-41a4-4f59-a5dc-25080e354730
set "key=44RPN-FTY23-9VTTB-MP9BX-T84FV" &:: Enterprise G N
exit /b

:: Windows 10 [RS1]
:2d5a5a60-3040-48bf-beb0-fcd770c20ce0
set "key=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ" &:: Enterprise 2016 LTSB
exit /b

:9f776d83-7156-45b2-8a5c-359b9c9f22a3
set "key=QFFDN-GRT3P-VKWWX-X7T3R-8B639" &:: Enterprise 2016 LTSB N
exit /b

:3f1afc82-f8ac-4f6c-8005-1d233e606eee
set "key=6TP4R-GNPTD-KYYHQ-7B7DP-J447Y" &:: Pro Education
exit /b

:5300b18c-2e33-4dc2-8291-47ffcec746dd
set "key=YVWGF-BXNMC-HTQYQ-CPQ99-66QFC" &:: Pro Education N
exit /b

:: Windows 10 [TH]
:58e97c99-f377-4ef1-81d5-4ad5522b5fd8
set "key=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99" &:: Home
exit /b

:7b9e1751-a8da-4f75-9560-5fadfe3d8e38
set "key=3KHY7-WNT83-DGQKR-F7HPR-844BM" &:: Home N
exit /b

:cd918a57-a41b-4c82-8dce-1a538e221a83
set "key=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH" &:: Home Single Language
exit /b

:a9107544-f4a0-4053-a96a-1479abdef912
set "key=PVMJN-6DFY6-9CCP6-7BKTT-D3WVR" &:: Home China
exit /b

:2de67392-b7a7-462a-b1ca-108dd189f588
set "key=W269N-WFGWX-YVC9B-4J6C9-T83GX" &:: Pro
exit /b

:a80b5abf-76ad-428b-b05d-a47d2dffeebf
set "key=MH37W-N47XK-V7XM9-C7227-GCQG9" &:: Pro N
exit /b

:e0c42288-980c-4788-a014-c080d2e1926e
set "key=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2" &:: Education
exit /b

:3c102355-d027-42c6-ad23-2e7ef8a02585
set "key=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ" &:: Education N
exit /b

:73111121-5638-40f6-bc11-f1d7b0d64300
set "key=NPPR9-FWDCX-D2C8J-H872K-2YT43" &:: Enterprise
exit /b

:e272e3e2-732f-4c65-a8f0-484747d0d947
set "key=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4" &:: Enterprise N
exit /b

:7b51a46c-0c04-4e8f-9af4-8496cca90d5e
set "key=WNMTR-4C88C-JK8YV-HQ7T2-76DF9" &:: Enterprise 2015 LTSB
exit /b

:87b838b7-41b6-4590-8318-5797951d8529
set "key=2F77B-TNFGY-69QQF-B8YKP-D69TJ" &:: Enterprise 2015 LTSB N
exit /b

:: Windows Server 2019 [RS5]
:de32eafd-aaee-4662-9444-c1befb41bde2
set "key=N69G4-B89J2-4G8F4-WWYCC-J464C" &:: Standard
exit /b

:34e1ae55-27f8-4950-8877-7a03be5fb181
set "key=WMDGN-G9PQG-XVVXX-R3X43-63DFG" &:: Datacenter
exit /b

:034d3cbb-5d4b-4245-b3f8-f84571314078
set "key=WVDHN-86M7X-466P6-VHXV7-YY726" &:: Essentials
exit /b

:a99cc1f0-7719-4306-9645-294102fbff95
set "key=FDNH6-VW9RW-BXPJ7-4XTYG-239TB" &:: Azure Core
exit /b

:73e3957c-fc0c-400d-9184-5f7b6f2eb409
set "key=N2KJX-J94YW-TQVFB-DG9YT-724CC" &:: Standard ACor
exit /b

:90c362e5-0da1-4bfd-b53b-b87d309ade43
set "key=6NMRW-2C8FM-D24W7-TQWMY-CWH2D" &:: Datacenter ACor
exit /b

:8de8eb62-bbe0-40ac-ac17-f75595071ea3
set "key=GRFBW-QNDC4-6QBHG-CCK3B-2PR88" &:: ServerARM64
exit /b

:: Windows Server 2016 [RS4]
:43d9af6e-5e86-4be8-a797-d072a046896c
set "key=K9FYF-G6NCK-73M32-XMVPY-F9DRR" &:: ServerARM64
exit /b

:: Windows Server 2016 [RS3]
:61c5ef22-f14f-4553-a824-c4b31e84b100
set "key=PTXN8-JFHJM-4WC78-MPCBR-9W4KR" &:: Standard ACor
exit /b

:e49c08e7-da82-42f8-bde2-b570fbcae76c
set "key=2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG" &:: Datacenter ACor
exit /b

:: Windows Server 2016 [RS1]
:8c1c5410-9f39-4805-8c9d-63a07706358f
set "key=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY" &:: Standard
exit /b

:21c56779-b449-4d20-adfc-eece0e1ad74b
set "key=CB7KF-BWN84-R7R2Y-793K2-8XDDG" &:: Datacenter
exit /b

:2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283
set "key=JCKRF-N37P4-C2D82-9YXRT-4M63B" &:: Essentials
exit /b

:7b4433f4-b1e7-4788-895a-c45378d38253
set "key=QN4C6-GBJD2-FB422-GHWJK-GJG2R" &:: Cloud Storage
exit /b

:3dbf341b-5f6c-4fa7-b936-699dce9e263f
set "key=VP34G-4NPPG-79JTQ-864T4-R3MQX" &:: Azure Core
exit /b

::========================================================================================================================================