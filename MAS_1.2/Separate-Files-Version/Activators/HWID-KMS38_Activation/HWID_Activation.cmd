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
title  [HWID] Digital License Activation
set Unattended=
set _args=
set _elev=
set "_arg1=%~1"
if not defined _arg1 goto :DL_NoProgArgs
set "_args=%~1"
set "_arg2=%~2"
if defined _arg2 set "_args=%~1 %~2"
for %%A in (%_args%) do (
if /i "%%A"=="-el" set _elev=1
if /i "%%A"=="/u" set Unattended=1)
:DL_NoProgArgs
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set "_psc=powershell -nop -ep bypass -c"
set "nul=1>nul 2>nul"
set "Red="white" "DarkRed""
set "Green="white" "DarkGreen""
set "Magenta="white" "darkmagenta""
set "Gray="white" "darkgray""
set "Black="white" "Black""
set "ELine=echo: &call :DL_color "==== ERROR ====" %Red% &echo:"
set slp=SoftwareLicensingProduct
set sls=SoftwareLicensingService
set wApp=55c92734-d682-4d71-983e-d6ec3f16059f

::========================================================================================================================================

for %%i in (powershell.exe) do if "%%~$path:i"=="" (
echo: &echo ==== ERROR ==== &echo:
echo Powershell is not installed in the system.
echo Aborting...
goto DL_Done
)

::========================================================================================================================================

if %winbuild% LSS 10240 (
%ELine%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 10.
goto DL_Done
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop
::  Thanks to @hearywarlot [ https://forums.mydigitallife.net/threads/.74332/ ] for the VBS method.
::  Thanks to @abbodi1406 for the powershell method and solving special characters issue in file path name.

%nul% reg query HKU\S-1-5-19 && (
  goto :DL_Passed
  ) || (
  if defined _elev goto :DL_E_Admin
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
goto :DL_E_Admin
)
)
exit /b

:DL_E_Admin
%ELine%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
goto DL_Done

:DL_Passed

::========================================================================================================================================

::  Fix for the special characters limitation in path name
::  Written by @abbodi1406

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

mode con: cols=102 lines=31

::  Check Windows OS name

set winos=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do if not errorlevel 1 set "winos=%%b"
if not defined winos for /f "tokens=2* delims== " %%a in ('"wmic os get caption /value" 2^>nul') do if not errorlevel 1 set "winos=%%b"

call :DL_CheckPermAct
if defined PermAct (

echo ___________________________________________________________________________________________
echo:
call :DL_color1 "     " %Black% &call :DL_color "Checking: %winos% is Permanently Activated." %Green%
call :DL_color1 "     " %Black% &call :DL_color "Activation is not required." %Gray%
echo ___________________________________________________________________________________________
echo:
if defined Unattended goto DL_Done

echo      Press [1] or [2] button in Keyboard :
echo ___________________________________________
echo:
choice /C:12 /N /M ">    [1] Activate [2] Exit : "

if errorlevel 2 exit /b
if errorlevel 1 Goto DL_Continue
)

:DL_Continue
cls

::========================================================================================================================================

cd /d "!_work!"
pushd "!_work!"

if not exist "!_work!\BIN\" (
%ELine%
echo 'BIN' Folder does not exist in current directory.
echo It's supposed to have files required for the Activation.
goto DL_Done
)

::========================================================================================================================================

echo %winos%| findstr /I Evaluation >nul && set Eval=1||set Eval=
if defined Eval (
%ELine%
echo [%winos% ^| %winbuild%] HWID Activation is Not Supported.
echo %winos%| findstr /I Server >nul && (
echo Server Evaluation cannot be activated. Convert it to full Server OS.
) || (
echo Evaluation Editions cannot be activated. Install full Windows OS.
)
goto DL_Done
)

::========================================================================================================================================

::  Check SKU value

set SKU=
for /f "tokens=2 delims==" %%a IN ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" 2^>nul') do if not errorlevel 1 (set osSKU=%%a)
if not defined SKU for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn 2^>nul') do if not errorlevel 1 (set osSKU=%%a)

if "%osSKU%"=="" (
%ELine%
echo SKU value was not detected properly. Aborting...
goto DL_Done
)

::  Check Windows Edition with SKU value for better accuracy

set osedition=
call :_CheckEdition %nul%

if "%osedition%"=="" (
%ELine%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%] HWID Activation is Not Supported.
goto DL_Done
)

::  Check Windows Architecture 

set arch=
reg Query "HKLM\Hardware\Description\System\CentralProcessor\0" | find /i "x86" > nul && set arch=x86|| set arch=x64
wmic os get osarchitecture | find /i "ARM" > nul && set arch=ARM64|| echo %PROCESSOR_ARCHITECTURE% | find /i "ARM" > nul && set arch=ARM64

if "%arch%"=="ARM64" call :DL_check ARM64_gatherosstate.exe ARM64_slc.dll
if not "%arch%"=="ARM64" call :DL_check gatherosstate.exe slc.dll
if defined _miss goto DL_Done

::========================================================================================================================================

cls
set key=
call :%osedition% %nul%

if "%key%"=="" (
%ELine%
echo [%winos% ^| %winbuild%] HWID Activation is Not Supported.
goto DL_Done
)

::========================================================================================================================================

cls
echo Checking OS Info                        [%winos% ^| %winbuild% ^| %arch%]

set "Chkint=Checking Internet Connection           "
set IntCon=
ping -n 1 www.microsoft.com %nul% && (
set IntCon=1
echo %Chkint% [Connected]
) || (
call :DL_color "%Chkint% [Not connected] [Ping www.microsoft.com - Failed]" %Red%
)

::========================================================================================================================================

echo:
set _1=ClipSVC
set _2=wlidsvc
set _3=sppsvc
set _4=wuauserv

for %%# in (%_1% %_2% %_3% %_4%) do call :DL_ServiceCheck %%#

set "CLecho=Checking %_1%                        [Service Status -%Cl_state%] [Startup Type -%Cl_start_type%]"
set "wlecho=Checking %_2%                        [Service Status -%wl_state%] [Startup Type -%wl_start_type%]"
set "specho=Checking %_3%                         [Service Status -%sp_state%] [Startup Type -%sp_start_type%]"
set "wuecho=Checking %_4%                       [Service Status -%wu_state%] [Startup Type -%wu_start_type%]"

if not "%Cl_start_type%"=="Demand"       (call :DL_color "%CLecho%" %Red% & set Clst_e=1) else (echo %CLecho%)
if not "%wl_start_type%"=="Demand"       (call :DL_color "%wlecho%" %Red% & set wlst_e=1) else (echo %wlecho%)
if not "%sp_start_type%"=="Delayed-Auto" (call :DL_color "%specho%" %Red% & set spst_e=1) else (echo %specho%)

if "%wu_start_type%"=="Disabled" (set "_C=%Red%") else (set "_C=%Gray%")
if not "%wu_start_type%"=="Auto"         (call :DL_color "%wuecho%" %_C% & set wust_e=1) else (echo %wuecho%)

echo:
if defined Clst_e (sc config %_1% start= Demand %nul%       && set Clst_s=%_1%-Demand || set Clst_u=%_1%-Demand )
if defined wlst_e (sc config %_2% start= Demand %nul%       && set wlst_s=%_2%-Demand || set wlst_u=%_2%-Demand )
if defined spst_e (sc config %_3% start= Delayed-Auto %nul% && set spst_s=%_3%-Delayed-Auto || set spst_u=%_3%-Delayed-Auto )
if defined wust_e (sc config %_4% start= Auto %nul%         && set wust_s=%_4%-Auto || set wust_u=%_4%-Auto )

for %%# in (Clst_s,wlst_s,spst_s,wust_s) do if defined %%# set st_s=1
if defined st_s (echo Changing services Startup Type to       [ %Clst_s%%wlst_s%%spst_s%%wust_s%] [Successful])

for %%# in (Clst_u,wlst_u,spst_u,wust_u) do if defined %%# set st_u=1
if defined st_u (call :DL_color "Error in changing Startup Type to       [ %Clst_u%%wlst_u%%spst_u%%wust_u%]" %Red%)

if not "%Cl_state%"=="Running" (%_psc% start-service %_1% %nul% && set Cl_s=%_1% || set Cl_u=%_1% )
if not "%wl_state%"=="Running" (%_psc% start-service %_2% %nul% && set wl_s=%_2% || set wl_u=%_2% )
if not "%sp_state%"=="Running" (%_psc% start-service %_3% %nul% && set sp_s=%_3% || set sp_u=%_3% )
if not "%wu_state%"=="Running" (%_psc% start-service %_4% %nul% && set wu_s=%_4% || set wu_u=%_4% )

for %%# in (Cl_s,wl_s,sp_s,wu_s) do if defined %%# set s_s=1
if defined s_s (echo Starting services                       [ %Cl_s%%wl_s%%sp_s%%wu_s%] [Successful])

for %%# in (Cl_u,wl_u,sp_u,wu_u) do if defined %%# set s_u=1
if defined s_u (call :DL_color "Error in starting services              [ %Cl_u%%wl_u%%sp_u%%wu_u%]" %Red%)

if defined wust_u (
call :DL_color "Most likely the Windows Update Service was blocked with a tool, identify and unblock it." %Magenta%
)

::========================================================================================================================================

::  Thanks to @abbodi1406 for the WMI methods

echo:
set _channel=
set _Keyexist=
set _partial=

for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get ProductKeyChannel /value 2^>nul') do set "_channel=%%#"
for %%A in (MAK, OEM, Retail) do echo %_channel%| findstr /i %%A >nul && set _Keyexist=1

if defined _Keyexist (
for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get PartialProductKey /value 2^>nul') do set "_partial=%%#"
call echo Checking Installed Product Key          [%_channel%] [Partial Key -%%_partial%%]
)

if not defined _Keyexist (
set "InsKey=Installing Generic Product Key         "
wmic path %sls% where __CLASS='%sls%' call InstallProductKey ProductKey="%key%" %nul% && (
for /f "tokens=2 delims==" %%# in ('wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" Get ProductKeyChannel /value 2^>nul') do set "_channel=%%#"
call echo %%InsKey%% [%key%] [%%_channel%%] [Successful]
) || (
call :DL_color "%%InsKey%% [%key%] [Unsuccessful]" %Red%
)
)

wmic path %sls% where __CLASS='%sls%' call RefreshLicenseStatus %nul%

::========================================================================================================================================

:: Files are copied to temp to generate ticket to avoid possible issues in case the path contains special character or non English names.

echo:
set "temp_=%SystemRoot%\Temp\_Ticket_Work"
if exist "%temp_%\" @RD /S /Q "%temp_%\" %nul%
md "%temp_%\" %nul%

cd /d "!_work!\BIN\"

set ARM64_file=
if "%arch%"=="ARM64" set ARM64_file=ARM64_

copy /y /b "%ARM64_file%gatherosstate.exe" "%temp_%\gatherosstate.exe" %nul%
copy /y /b "%ARM64_file%slc.dll" "%temp_%\slc.dll" %nul%

set cfailed=
if not exist "%temp_%\gatherosstate.exe" set cfailed=1
if not exist "%temp_%\slc.dll" set cfailed=1

set "copyfiles=Copying Required Files to Temp         "
if defined cfailed (
call :DL_color "%copyfiles% [%SystemRoot%\Temp\_Ticket_Work\] [Unsuccessful] Aborting..." %Red%
goto :DL_Act_Cont
) else (
echo %copyfiles% [%SystemRoot%\Temp\_Ticket_Work\] [Successful]
)

cd /d "%temp_%\"
attrib -R -A -S -H *.*

::========================================================================================================================================

set "GatherMod=Creating modified gatherosstate        "

if not "%arch%"=="ARM64" (
rundll32 "%temp_%\slc.dll",PatchGatherosstate %nul%
if not exist "%temp_%\gatherosstatemodified.exe" (
call :DL_color "%GatherMod% [Unsuccessful] Aborting" %Red%
call :DL_color "Most likely Antivirus program blocked the process, disable it and-or create proper exclsuions." %Magenta%
goto :DL_Act_Cont
) else (
echo %GatherMod% [Successful]
)
)

::========================================================================================================================================

set _gather=
if "%arch%"=="ARM64" (
set _gather=gatherosstate.exe
) else (
set _gather=gatherosstatemodified.exe
)

start /wait "" "%temp_%/%_gather%" %nul%
if not exist "%temp_%\GenuineTicket.xml" (
call "%temp_%/%_gather%" %nul%
)

set "GenTicket=Generating GenuineTicket.xml           "
if not exist "%temp_%\GenuineTicket.xml" (
call :DL_color "%GenTicket% [Unsuccessful] Aborting..." %Red%
goto :DL_Act_Cont
) else (
echo %GenTicket% [Successful]
)

:: clipup -v -o -altto <Ticket path> method to apply ticket was not used to avoid the certain issues in case the username have spaces or non English names.

set "InsTicket=Installing GenuineTicket.xml           "
set "TDir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if exist "%TDir%\*.xml" del /f /q "%TDir%\*.xml" %nul%
copy /y /b "%temp_%\GenuineTicket.xml" "%TDir%\GenuineTicket.xml" %nul%

if not exist "%TDir%\GenuineTicket.xml" (
call :DL_color "Failed to copy Ticket to [%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket\] Aborting..." %Red%
goto :DL_Act_Cont
)

%_psc% Restart-Service ClipSVC %nul%

if exist "%TDir%\GenuineTicket.xml" (
net stop ClipSVC %nul%
net start ClipSVC %nul%
)

if exist "%TDir%\GenuineTicket.xml" (
%nul% clipup -v -o
set "fallback_=[Fallback method: clipup -v -o]"
)

if not exist "%TDir%\GenuineTicket.xml" (
echo %InsTicket% [Successful] %fallback_%
) else (
call :DL_color "%InsTicket% [Unsuccessful] Aborting..." %Red%
if exist "%TDir%\*.xml" del /f /q "%TDir%\*.xml" %nul%
goto :DL_Act_Cont
)

::========================================================================================================================================

echo:
echo Activating...
echo:

wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" call Activate %nul%
call :DL_CheckPermAct
if defined PermAct goto DL_Act_successful

call :DL_ReTry
if not "%ErrCode%"=="" set "Error_Code_=[Error Code %ErrCode%]"
call :DL_CheckPermAct

:DL_Act_successful

if defined PermAct (
call :DL_color "%winos% is permanently activated." %Green%
goto :DL_Act_Cont
)

call :DL_color "Activation Failed %Error_Code_%" %Red%
call :DL_color "Try the Troubleshoot Guide listed in the ReadMe." %Magenta%

::========================================================================================================================================

:DL_Act_Cont

echo:
set "changing_wust_back=Changing wu Startup Type back to        [%wu_start_type%]"
if defined wust_s (
sc config %_4% start= %wu_start_type% %nul% && echo %changing_wust_back% [Successful]
) || (
call :DL_color "%changing_wust_back% [Unsuccessful]" %Red%
)

cd /d "!_work!\"
if exist "%temp_%\" @RD /S /Q "%temp_%\" %nul%
set "delFiles=Cleaning Temp Files                    "
if exist "%temp_%\" (
call :DL_color "%delFiles% [Unsuccessful]" %Red%
) else (
echo %delFiles% [Successful]
)

goto DL_Done

::========================================================================================================================================

::  Echo all the missing files.
::  Written by @abbodi1406 (MDL)

:DL_check

for %%# in (%1 %2) do (if not exist "!_work!\BIN\%%#" (if defined _miss (set "_miss=!_miss! %%#") else (set "_miss=%%#")))
if defined _miss (
%ELine%
echo Following required file^(s^) is missing in 'BIN' folder. Aborting...
echo:
echo !_miss!
)
exit /b

::========================================================================================================================================

:DL_ServiceCheck

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
if /i "%%i"=="%_4%" set "wu_start_type=%start_type%" & set "wu_state=%state%"
if /i "%%i"=="%_3%" set "sp_start_type=%start_type%" & set "sp_state=%state%"
if /i "%%i"=="%_1%" set "Cl_start_type=%start_type%" & set "Cl_state=%state%"
if /i "%%i"=="%_2%" set "wl_start_type=%start_type%" & set "wl_state=%state%"
)
exit /b

::========================================================================================================================================

:DL_CheckPermAct

::  Check Windows Permanent Activation status
::  Written by @abbodi1406

wmic path %slp% where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name 2>nul | findstr /i "Windows" 1>nul && set PermAct=1||set PermAct=
exit /b

::========================================================================================================================================

:DL_ReTry

if defined IntCon if not defined wust_u if not defined wu_u call :DL_ReTry_2

::  Detect Error Code in the Activation
::  Written by @abbodi1406

wmic path %slp% where "ApplicationID='%wApp%' and PartialProductKey<>null" call Activate %nul%
set errorcode=%errorlevel%
cmd /c exit /b %errorcode%
if %errorcode% NEQ 0 set "ErrCode=0x%=ExitCode%"
exit /b

:DL_ReTry_2

set app=
%_psc% Restart-Service sppsvc %nul%

::  Specific rearm - reset the licensing status of the Windows SKU and app, without the need to restart the system.
::  wmic method by @abbodi1406

for /f "tokens=2 delims==" %%G in ('"wmic path %slp% where (ApplicationID='%wApp%' AND ProductKeyID like '%%-%%') get ID /value" 2^>nul') do (set app=%%G)
wmic path %sls% where __CLASS='%sls%' call ReArmApp ApplicationId="%wApp%" %nul%
wmic path %slp% where ID='%app%' call ReArmsku %nul%

wmic path %sls% where __CLASS='%sls%' call RefreshLicenseStatus %nul%
cscript /nologo %windir%\system32\slmgr.vbs -ato %nul%
exit /b

::========================================================================================================================================

:DL_color

%_psc% write-host '%1' -fore '%2' -back '%3'
exit /b

:DL_color1

%_psc% write-host '%1' -fore '%2' -back '%3'  -NoNewline
exit /b

::========================================================================================================================================

:DL_Done

echo:
if defined Unattended (
echo Exiting in 3 seconds...
if %winbuild% LSS 7600 (ping -n 3 127.0.0.1 > nul) else (timeout /t 3)
:: set a value to use in certain conditions of setupcomplete.cmd file.
if defined key if not defined PermAct (endlocal & endlocal & set HWIDAct=1)
exit /b
)
echo Press any key to exit...
pause >nul
exit /b

::========================================================================================================================================

::  Check Windows Edition with SKU value for better accuracy

:_CheckEdition

for %%# in (
4:Enterprise
27:EnterpriseN
48:Professional
49:ProfessionalN
98:CoreN
99:CoreCountrySpecific
100:CoreSingleLanguage
101:Core
121:Education
122:EducationN
125:EnterpriseS
126:EnterpriseSN
161:ProfessionalWorkstation
162:ProfessionalWorkstationN
164:ProfessionalEducation
165:ProfessionalEducationN
175:ServerRdsh
188:IoTEnterprise
) do for /f "tokens=1,2 delims=:" %%A in ("%%#") do (
if %osSKU%==%%A set "osedition=%%B"
)
exit /b

::========================================================================================================================================

::  Retail_OEM Key List

:Core
set "key=YTMG3-N6DKC-DKB77-7M9GH-8HVX7"
exit /b

:CoreCountrySpecific
set "key=N2434-X9D7W-8PF6X-8DV9T-8TYMD"
exit /b

:CoreN
set "key=4CPRK-NM3K3-X6XXQ-RXX86-WXCHW"
exit /b

:CoreSingleLanguage
set "key=BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"
exit /b

:Education
set "key=YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"
exit /b

:EducationN
set "key=84NGF-MHBT6-FXBX8-QWJK7-DRR8H"
exit /b

:Enterprise
set "key=XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
exit /b

:EnterpriseN
set "key=3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT"
exit /b

:EnterpriseS
if "%winbuild%" EQU "10240" set "key=FWN7H-PF93Q-4GGP8-M8RF3-MDWWW"
if "%winbuild%" EQU "14393" set "key=NK96Y-D9CD8-W44CQ-R8YTK-DYJWX"
exit /b

:EnterpriseSN
if "%winbuild%" EQU "10240" set "key=8V8WN-3GXBH-2TCMG-XHRX3-9766K"
if "%winbuild%" EQU "14393" set "key=2DBW3-N2PJG-MVHW3-G7TDK-9HKR4"
exit /b

:Professional
set "key=VK7JG-NPHTM-C97JM-9MPGT-3V66T"
exit /b

:ProfessionalEducation
set "key=8PTT6-RNW4C-6V7J2-C2D3X-MHBPB"
exit /b

:ProfessionalEducationN
set "key=GJTYN-HDMQY-FRR76-HVGC7-QPF8P"
exit /b

:ProfessionalN
set "key=2B87N-8KFHP-DKV6R-Y2C8J-PKCKT"
exit /b

:ProfessionalWorkstation
set "key=DXG7C-N36C4-C4HTG-X4T3X-2YV77"
exit /b

:ProfessionalWorkstationN
set "key=WYPNQ-8C467-V2W6J-TX4WX-WT2RQ"
exit /b

:ServerRdsh
set "key=NJCF7-PW8QT-3324D-688JX-2YV66"
exit /b

:IoTEnterprise
set "key=XQQYW-NFFMW-XJPBH-K8732-CKFFD"
exit /b

::========================================================================================================================================