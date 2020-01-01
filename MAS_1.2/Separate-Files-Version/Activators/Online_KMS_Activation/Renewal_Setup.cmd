@setlocal DisableDelayedExpansion
@echo off






:: =======================================================================================================
::
::   This script is a part of 'Microsoft Activation Scripts' project.
::
::   Homepages-
::   NsaneForums: (Login Required) https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/
::   GitLab: https://gitlab.com/massgrave/microsoft-activation-scripts
::
::   Maintained by @WindowsAddict
::
:: =======================================================================================================












::========================================================================================================================================

cls
set Unattended=
set _args=
set _elev=
set RenTask=
set RenActTask=
set DeskMenu=
set _SkipWinAct=
set _end=
set "_arg1=%~1"
if not defined _arg1 goto :NoProgArgs
set "_args=%~1"
set "_arg2=%~2"
set "_arg3=%~3"
if defined _arg2 set "_args=%~1 %~2"
if defined _arg3 set "_args=%~1 %~2 %~3"
for %%A in (%_args%) do (
if /i "%%A"=="-el" set _elev=1
if /i "%%A"=="/swa" set _SkipWinAct=1
if /i "%%A"=="/rt" set RenTask=1&set Unattended=1
if /i "%%A"=="/rat" set RenActTask=1&set Unattended=1
if /i "%%A"=="/dcm" set DeskMenu=1&set Unattended=1)
:NoProgArgs
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set "_psc=powershell -nop -ep bypass -c"
set "nul=1>nul 2>nul"
set "EchoRed=%_psc% write-host -back Black -fore Red"
set "EchoGreen=%_psc% write-host -back Black -fore Green"
set "EchoYellow=%_psc% write-host -back Black -fore DarkYellow"
set "ELine=echo: & %EchoRed% ==== ERROR ==== &echo:"

::========================================================================================================================================

for %%i in (powershell.exe) do if "%%~$path:i"=="" (
echo: &echo ==== ERROR ==== &echo:
echo Powershell is not installed in the system.
echo Aborting...
set _end=1
goto Done
)

::========================================================================================================================================

if %winbuild% LSS 7600 (
%ELine%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10 and their Server equivalent.
set _end=1
goto Done
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
goto :E_Admin
)
)
exit /b

:E_Admin
%ELine%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'.
set _end=1
goto Done

:Passed

::========================================================================================================================================

::  Fix for the special characters limitation in path name
::  Written by @abbodi1406

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

if not exist "!_work!\Activate.cmd" (
%ELine%
echo File [Activate.cmd] does not exist in current folder..
echo It's required for the Task Creation.
set _end=1
goto Done
)

call :check cleanosppx64.exe cleanosppx86.exe
if defined _miss set _end=1&goto Done

::========================================================================================================================================

set "_dest=%ProgramData%\Online_KMS_Activation"
set "key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks"

:ActivationRenewal

cls
title Online KMS Activation Renewal
mode con cols=98 lines=30
set ActTask=
set error_=
set DelDeskCont=
set error_1=

if defined RenTask goto:Task
if defined RenActTask set ActTask=1&goto:Task
if defined DeskMenu goto:ContextMenu
echo:
echo:
echo:
echo                         You can apply the option [either 1 or 2] and [3].
echo                       ______________________________________________________
echo                      ^|                                                      ^|
echo                      ^|            Auto Renewal via Task Scheduler           ^|
echo                      ^|                                                      ^|
echo                      ^|   [1] Create Renewal Task                            ^|
echo                      ^|                                                      ^|
echo                      ^|   [2] Create Renewal and Activation Task             ^|
echo                      ^|______________________________________________________^|
echo                      ^|                                                      ^|
echo                      ^|        Manual Renewal via Desktop Context Menu       ^|
echo                      ^|                                                      ^|
echo                      ^|   [3] Add Desktop Context Menu                       ^|
echo                      ^|______________________________________________________^|
echo                      ^|                                                      ^|
echo                      ^|   [4] Exit                                           ^|
echo                      ^|                                                      ^|
echo                      ^|______________________________________________________^|
echo:
choice /C:1234 /N /M ".                     Enter Your Choice [1,2,3,4] : "

if errorlevel 4 exit /b
if errorlevel 3 goto:ContextMenu
if errorlevel 2 set ActTask=1&goto:Task
if errorlevel 1 goto:Task

:======================================================================================================================================================

:Task

cls
if defined ActTask (
title  Create Renewal And Activation Tasks
) else (
title  Create Renewal Task
)

reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Renewal" >nul && (
schtasks /delete /tn Online_KMS_Activation_Script-Renewal /f %nul%
)
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Run_Once" >nul && (
schtasks /delete /tn Online_KMS_Activation_Script-Run_Once /f %nul%
)
If exist "%_dest%\" (
@RD /s /q "%_dest%\" %nul%
)
If exist "%windir%\Online_KMS_Activation_Script\" (
@RD /s /q "%windir%\Online_KMS_Activation_Script\" %nul%
)
If exist "%ProgramData%\Online_KMS_Activation.cmd" (
Reg delete "HKCR\DesktopBackground\shell\Activate Windows - Office" /f %nul%
del /f /q "%ProgramData%\Online_KMS_Activation.cmd" %nul%
set DelDeskCont=1
)

md "%_dest%\BIN\" %nul%

set "_temp=%SystemRoot%\Temp\_KMS_Task_Work"
if exist "%_temp%\" @RD /S /Q "%_temp%\" %nul%
md "%_temp%\" %nul%

call :Export renewal "%_temp%\Renewal.xml" Unicode
if defined ActTask (call :Export run_once "%_temp%\Run_Once.xml" Unicode)

call :Export info "%_dest%\Info.txt" ASCII

copy /y /b "!_work!\BIN\cleanosppx64.exe" "%_dest%\BIN\cleanosppx64.exe" %nul%
copy /y /b "!_work!\BIN\cleanosppx86.exe" "%_dest%\BIN\cleanosppx86.exe" %nul%

cd /d "!_work!"

if defined _SkipWinAct (
%nul% %_psc% "(gc Activate.cmd) -replace 'set ActWindows=1', 'set ActWindows=0' | Out-File -encoding ASCII "%_dest%\Activate.cmd"" || (set error_=1)
) else (
copy /y /b "!_work!\Activate.cmd" "%_dest%\Activate.cmd" %nul%
)
schtasks /create /tn "Online_KMS_Activation_Script-Renewal" /ru "SYSTEM" /xml "%_temp%\Renewal.xml" %nul%
if defined ActTask (schtasks /create /tn "Online_KMS_Activation_Script-Run_Once" /ru "SYSTEM" /xml "%_temp%\Run_Once.xml" %nul%)

if exist "%_temp%\" @RD /S /Q "%_temp%\" %nul%

::========================================================================================================================================

reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Renewal" >nul || (set error_=1)
if defined ActTask reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Run_Once" >nul || (set error_=1)

If not exist "%_dest%\Activate.cmd" (set error_=1)
If not exist "%_dest%\Info.txt" (set error_=1)
If not exist "%_dest%\BIN\cleanosppx64.exe" (set error_=1)
If not exist "%_dest%\BIN\cleanosppx86.exe" (set error_=1)

if defined error_ (
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Renewal" >nul && (
schtasks /delete /tn Online_KMS_Activation_Script-Renewal /f %nul%
)
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Run_Once" >nul && (
schtasks /delete /tn Online_KMS_Activation_Script-Run_Once /f %nul%
)
reg delete "HKCR\DesktopBackground\shell\Activate Windows - Office" /f %nul%
If exist "%_dest%\" (
@RD /s /q "%_dest%\" %nul%
)
echo _________________________________________________________________
echo:
%ELine%
echo Run the Online KMS Complete Uninstall script and then try again.
echo _________________________________________________________________
) else (
echo:
echo __________________________________________________________________________________________
echo:
if defined DelDeskCont (
%EchoYellow% Previous desktop context menu entry for Online KMS Activation is deleted.
echo:
)
if defined _SkipWinAct (
%EchoYellow% %_dest%\Activate.cmd is set to skip Windows Activation.
echo:
)

echo Files created:
echo %_dest%\BIN\cleanosppx64.exe
echo %_dest%\BIN\cleanosppx86.exe
echo %_dest%\Activate.cmd
echo %_dest%\Info.txt
echo:
echo Scheduled Tasks created:
echo \Online_KMS_Activation_Script-Renewal
if defined ActTask (echo \Online_KMS_Activation_Script-Run_Once)
echo:
echo It's recommended to set exclusion for the following file in your Antivirus Program.
echo:
echo %_dest%\Activate.cmd
echo __________________________________________________________________________________________
echo:
if defined ActTask (
%EchoGreen% Online KMS Activation - Renewal and Activation Tasks are successfully created.
) else (
%EchoGreen% Online KMS Activation - Renewal Task is successfully created.
)
echo __________________________________________________________________________________________
echo:
)

goto Done

::========================================================================================================================================

:ContextMenu

cls
title Add Desktop Context Menu

If exist "%ProgramData%\Online_KMS_Activation.cmd" (
del /f /q "%ProgramData%\Online_KMS_Activation.cmd" %nul%
set DelDeskCont=1
)

reg delete "HKCR\DesktopBackground\shell\Activate Windows - Office" /f %nul%

if exist "%_dest%\BIN\" (
@RD /s /q "%_dest%\BIN\" %nul%
)

md "%_dest%\BIN\" %nul%
copy /y /b "!_work!\BIN\cleanosppx64.exe" "%_dest%\BIN\cleanosppx64.exe" %nul%
copy /y /b "!_work!\BIN\cleanosppx86.exe" "%_dest%\BIN\cleanosppx86.exe" %nul%

if exist "%_dest%\Activate.cmd" (
del /f /q "%_dest%\Activate.cmd" %nul%
)

cd /d "!_work!"

if defined _SkipWinAct (
%nul% %_psc% "(gc Activate.cmd) -replace 'set ActWindows=1', 'set ActWindows=0' | Out-File -encoding ASCII "%_dest%\Activate.cmd"" || (set error_=1)
) else (
copy /y /b "!_work!\Activate.cmd" "%_dest%\Activate.cmd" %nul%
)

if exist "%_dest%\Info.txt" (
del /f /q "%_dest%\Info.txt" %nul%
)

call :Export info "%_dest%\Info.txt" ASCII

reg add "HKCR\DesktopBackground\shell\Activate Windows - Office" /v "Icon" /t REG_SZ /d "%SystemRoot%%\System32\shell32.dll,71" /f >nul 2>&1 || (set error_1=1)
reg add "HKCR\DesktopBackground\shell\Activate Windows - Office\command" /ve /d "%_dest%\Activate.cmd" /f %nul% || (set error_1=1)

If not exist "%_dest%\Activate.cmd" (set error_=1)
If not exist "%_dest%\Info.txt" (set error_=1)
If not exist "%_dest%\BIN\cleanosppx64.exe" (set error_=1)
If not exist "%_dest%\BIN\cleanosppx86.exe" (set error_=1)

reg query "HKCR\DesktopBackground\shell\Activate Windows - Office" %nul% || (set error_1=1)

if defined error_1 (
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Renewal" >nul && (
schtasks /delete /tn Online_KMS_Activation_Script-Renewal /f %nul%
)
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Run_Once" >nul && (
schtasks /delete /tn Online_KMS_Activation_Script-Run_Once /f %nul%
)
reg delete "HKCR\DesktopBackground\shell\Activate Windows - Office" /f %nul%
If exist "%_dest%\" (
@RD /s /q "%_dest%\" %nul%
)
echo _________________________________________________________________
echo:
%ELine%
echo Run the Online KMS Complete Uninstall script and then try again.
echo _________________________________________________________________
) else (
echo:
echo __________________________________________________________________________________________
echo:
if defined DelDeskCont (
%EchoYellow% Previous desktop context menu entry for Online KMS Activation is deleted.
echo:
)
if defined _SkipWinAct (
%EchoYellow% %_dest%\Activate.cmd is set to skip Windows Activation.
echo:
)

echo Files created:
echo %_dest%\BIN\cleanosppx64.exe
echo %_dest%\BIN\cleanosppx86.exe
echo %_dest%\Activate.cmd
echo %_dest%\Info.txt
echo:
echo Registry entry added:
echo HKCR\DesktopBackground\shell\Activate Windows - Office
echo HKCR\DesktopBackground\shell\Activate Windows - Office\command
echo __________________________________________________________________________________________
echo:
%EchoGreen% Desktop context menu entry for Online KMS Activation is successfully created.
echo __________________________________________________________________________________________
echo:
)

::========================================================================================================================================

:Done
echo:
if defined Unattended (
echo Exiting in 3 seconds...
if %winbuild% LSS 7600 (ping -n 3 127.0.0.1 > nul) else (timeout /t 3)
exit /b
)
if defined _end (
echo Press any key to exit...
pause >nul
exit /b
) else (
echo Press any key to go back...
pause >nul
goto ActivationRenewal
)

::========================================================================================================================================

:info:
====================================================================================================
   Online KMS Activation:
====================================================================================================

   The use of this script is to activate / renew your Windows /Server /Office license 
   using online KMS.
   
 - Scheduled task name (If Renewal Task is created) (Weekly).
   \Online_KMS_Activation_Script-Renewal

 - Scheduled task name (If Activation Task is created).
   \Online_KMS_Activation_Script-Run_Once

   The scheduled task runs only if the system is connected to the Internet.
   Activation Task will run on the system login and after successful activation, this task will 
   delete itself.
   
 - If system preactivation is done via HWID + Online KMS, and HWID was applied but was not 
   successful due to lack of internet at the time of installation of Windows, in that case, 
   Online KMS script will be set to skip Windows activation.

 - Registry entry name and location (If desktop context menu is created).
   HKCR\DesktopBackground\shell\Activate Windows - Office

   For complete script and more info, browse the script homepage.

====================================================================================================
   File Details:
====================================================================================================

   d30a0e4e5911d3ca705617d17225372731c770e2 *cleanosppx64.exe                   Virus Total = 0/66
   39ed8659e7ca16aaccb86def94ce6cec4c847dd6 *cleanosppx86.exe                   Virus Total = 1/66

   Virus Total Report Date: 12-11-2019
   
   These files are official Microsoft files and in this script, these are used in 
   cleaning office license in C2R Retail office to VL conversion process.
   
   The source of these files is the 'old' version of Microsoft Tool O15CTRRemove.diagcab
   You can get the original file here https://s.put.re/WFuXpyWA.zip

====================================================================================================

   Online KMS Activation script is just a fork of @abbodi1406's KMS_VL_ALL Project.
   KMS_VL_ALL homepage: https://forums.mydigitallife.net/posts/838808

   This fork was made to avoid having any KMS binary files and system can be activated using 
   some manual commands or transparent batch script files.

   Online KMS Activation script is a part of 'Microsoft Activation Scripts'
   Maintained by @WindowsAddict
   Homepage: https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/

====================================================================================================
:info:

:renewal:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Source>Microsoft Corporation</Source>
    <Date>1999-01-01T12:00:00.34375</Date>
    <Author>RPO/WindowsAddict</Author>
    <Version>1.0</Version>
    <Description>Online_KMS_Activation_Script-Renewal - Weekly Activation Renewal Task</Description>
    <URI>\Online_KMS_Activation_Script-Renewal</URI>
    <SecurityDescriptor>D:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FRFX;;;LS)(A;;FRFW;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)(A;;FR;;;S-1-5-4)</SecurityDescriptor>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>1999-01-01T12:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Sunday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="LocalSystem">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT2M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="LocalSystem">
    <Exec>
      <Command>%ProgramData%\Online_KMS_Activation\Activate.cmd</Command>
    <Arguments>Task</Arguments>
    </Exec>
  </Actions>
</Task>
:renewal:

:run_once:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Source>Microsoft Corporation</Source>
    <Date>1999-01-01T12:00:00.34375</Date>
    <Author>RPO/WindowsAddict</Author>
    <Version>1.0</Version>
    <Description>Online_KMS_Activation_Script-Run_Once - Run and Delete itself on first Internet Contact</Description>
    <URI>\Online_KMS_Activation_Script-Run_Once</URI>
    <SecurityDescriptor>D:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FRFX;;;LS)(A;;FRFW;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)(A;;FR;;;S-1-5-4)</SecurityDescriptor>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="LocalSystem">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT2M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="LocalSystem">
    <Exec>
      <Command>%ProgramData%\Online_KMS_Activation\Activate.cmd</Command>
    <Arguments>Task</Arguments>
    </Exec>
  </Actions>
</Task>
:run_once:

::========================================================================================================================================

::  Echo all the missing files.
::  Written by @abbodi1406 (MDL)

:check

for %%# in (%1 %2) do (if not exist "!_work!\BIN\%%#" (if defined _miss (set "_miss=!_miss! %%#") else (set "_miss=%%#")))
if defined _miss (
%ELine%
echo Following required file^(s^) is missing in 'BIN' folder. Aborting...
echo:
echo !_miss!
)
exit /b

::========================================================================================================================================

::  Extract the text from batch script without character and file encoding issue
::  Thanks to @abbodi1406

:Export
%nul% %_psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('%~2',$f[1].Trim(),[System.Text.Encoding]::%~3);" &exit/b
exit /b

::========================================================================================================================================