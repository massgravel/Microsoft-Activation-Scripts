====================================================================================================
   $OEM$ Folders [Windows Pre-Activation]:
====================================================================================================

 - To create a Preactivated Windows installation .iso, do the following things:
   Copy the "$OEM$" folder to the "sources" folder in the Windows installation media (.iso or USB).
   The directory will appear like this: \sources\$OEM$ in your altered .iso or on your bootable 
   USB drive.
   Now use this .iso or bootable USB drive to install Windows and it will either already be activated 
   (KMS38) as soon as it boots, or will self-activate at first internet contact. 

----------------------------------------------------------------------------------------------------
   HWID:
----------------------------------------------------------------------------------------------------

 - When using Digital License (HWID), no files are stored on the system, and when connected to the 
   internet for the first time, the system will self-activate at that time.
   
----------------------------------------------------------------------------------------------------
   KMS38:
----------------------------------------------------------------------------------------------------

 - When using KMS38, no files are stored on the system, and Windows becomes activated immediately 
   without further actions or connectivity of any kind being required.

----------------------------------------------------------------------------------------------------
   Online KMS (separately, or in combination with HWID or KMS38):
----------------------------------------------------------------------------------------------------

 - It creates the following 2 Activation/Renewal Methods. You can turn off any of them in 
   setupcomplete.cmd file

   ----------------------------------------------------------

   1- Automatic Renewal via Task Scheduler---

   It creates following files and tasks,

   Files:
   C:\ProgramData\Online_KMS_Activation\BIN\cleanosppx64.exe
   C:\ProgramData\Online_KMS_Activation\BIN\cleanosppx86.exe
   C:\ProgramData\Online_KMS_Activation\Activate.cmd
   C:\ProgramData\Online_KMS_Activation\Info.txt

   Scheduled Tasks:
   \Online_KMS_Activation_Script-Renewal  (Weekly)
   \Online_KMS_Activation_Script-Run_Once (Activation Task)

   The scheduled task runs only if the system is connected to the Internet.
   Activation Task will run on the system login and after successful activation and registering 
   online KMS server, this task will delete itself. leaving behind only one task to run weekly 
   for the lifetime of the system. 

   ----------------------------------------------------------
   
   2- Manual Renewal via Desktop Context Menu---

   It creates Desktop context Menu for manual activation and renewal.
   It creates the following files and registry entries.

   Files:
   C:\ProgramData\Online_KMS_Activation\BIN\cleanosppx64.exe
   C:\ProgramData\Online_KMS_Activation\BIN\cleanosppx86.exe
   C:\ProgramData\Online_KMS_Activation\Activate.cmd
   C:\ProgramData\Online_KMS_Activation\Info.txt
   
   Registry entries:
   HKCR\DesktopBackground\shell\Activate Windows - Office
   HKCR\DesktopBackground\shell\Activate Windows - Office\command

   It creates an easy to reach the Desktop context menu for the manual activation renewal.
   
   ----------------------------------------------------------

   d30a0e4e5911d3ca705617d17225372731c770e2 *cleanosppx64.exe                   Virus Total = 0/66
   39ed8659e7ca16aaccb86def94ce6cec4c847dd6 *cleanosppx86.exe                   Virus Total = 1/66

   Virus Total Report Date: 12-11-2019
   
   These files are official Microsoft files and in this script, these are used in 
   cleaning office license in C2R Retail office to VL conversion process.
   
   The source of these files is the 'old' version of Microsoft Tool O15CTRRemove.diagcab
   You can get the original file here https://s.put.re/WFuXpyWA.zip

   ----------------------------------------------------------
   
   IMPORTANT NOTE - Some sensitive AV's may flag the Automatic Renewal via the Task, and not 
   because of KMS, because for them it's suspicious to run long scripts in the background as Tasks.

   It's recommended to set exclusions in Antivirus for
   C:\ProgramData\Online_KMS_Activation\Activate.cmd

   ----------------------------------------------------------
   
 - When using Online KMS plus HWID Digital License, Online KMS script will be set to skip Windows 
   activation (if the HWID activation was applied but was not successful due to lack of internet 
   at the time of installation of Windows) but will register the KMS for other products, and all 
   later installed Volume License (VL) products (MS Office) will self-activate when going online. 

 - When using Online KMS plus KMS38, Online KMS will not skip Windows activation but skip KMS38 
   activation and will register the KMS for other products, and all subsequently-installed Volume 
   License (VL) products (MS Office) will self-activate when going online.

----------------------------------------------------------------------------------------------------
   HWID (Fallback to KMS38):
----------------------------------------------------------------------------------------------------
  
 - In this method, KMS38 will be used for the activation in case the Windows version is not 
   supported by HWID. For example, Windows 10 LTSC and Windows server.

----------------------------------------------------------------------------------------------------
   Activation Type       Supported Product             Activation Period
----------------------------------------------------------------------------------------------------
   
   Digital License    -  Windows 10                 -  Permanent
   KMS38              -  Windows 10 / Server        -  Until the year 2038
   Online KMS         -  Windows / Server / Office  -  For 180 Days, renewal task needs to be 
                                                       created for lifetime auto-activation.
   
----------------------------------------------------------------------------------------------------
   
   * For more details, use the ReadMe.txt included in the respective activation folders.