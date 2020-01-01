====================================================================================================
   KMS38 Activation:
====================================================================================================

 - This activation method activates Windows 10 and Windows Server (14393 and later builds), 
   until the year 2038.
 - This activation method does not store any files on the system.

 - Make sure the following things have been accounted for, before applying KMS38 Activation:
   - Before the activation, if any KMS activator is installed, then make sure to uninstall it 
     completely.
   - After KMS38 activation for the Windows Operating System has been achieved, if you want to 
     additionally, use the 180 Days KMS Activator for MS Office, then you must make sure that 
     it (the 180 Days KMS Activator for MS Office) is compatible with Windows KMS38 activation. 
     FYI, the following activators are compatible and can activate Office 
     without disrupting the Windows KMS38 activation.

     KMS_VL_ALL by @abbodi1406     https://forums.mydigitallife.net/posts/838808
     Online KMS Activation Script  https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/

 - Any KMS Activator which is not compatible with KMS38, may overwrite the KMS38 activation for 
   Windows with its own 180 days activation, thereby destroying the KMS38 activation for Windows. 
   To prevent this accidental overwriting, you can apply KMS38 protection for Windows, check 
   the Extras folder for more details.

 - Why is the script setting the specific KMS host to 127.0.0.2 (localhost)?
   - By doing this, global KMS IP can not replace KMS38 activation but can be used with Office and
     other Windows Editions.
   - In case you don't like it, you can remove it with following codes, open CMD as admin and enter,

set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
reg delete "HKLM\%SPPk%\55c92734-d682-4d71-983e-d6ec3f16059f" /f
reg delete "HKU\S-1-5-20\%SPPk%\55c92734-d682-4d71-983e-d6ec3f16059f" /f

====================================================================================================

   Documentation by @mspaintmsi

   Included topics-

   How does it work?
   
   https://pastebin.com/raw/7Xyaf15Z
   Mirror Link-
   https://textuploader.com/1dg8d/raw

====================================================================================================
   Supported Products:
====================================================================================================

   Windows 10:

   Core
   CoreCountrySpecific
   CoreN
   CoreSingleLanguage
   Education
   EducationN
   Enterprise
   EnterpriseG
   EnterpriseGN
   EnterpriseN
   EnterpriseS       [LTSB 2016 & LTSC 2019]
   EnterpriseSN      [LTSB 2016 & LTSC 2019]
   Professional
   ProfessionalEducation
   ProfessionalEducationN
   ProfessionalN
   ProfessionalWorkstation
   ProfessionalWorkstationN
   ServerRdsh

   ------------------------

   Windows Server:

   ServerCloudStorage     [Server 2016]
   ServerDatacenter       [Server 2016 & 2019]
   ServerDatacenterCor    [Server 2016 & 2019]
   ServerSolution         [Server 2016 & 2019]
   ServerSolutionCor      [Server 2016 & 2019]
   ServerStandard         [Server 2016 & 2019]
   ServerStandardCor      [Server 2016 & 2019]
   ServerAzureCor         [Server 2016 & 2019]
   ServerDatacenterACor   [All versions]
   ServerStandardACor     [All versions]


   Note - X86-X64 and ARM64 architecture systems are supported.
        - Any Evaluation version of Windows and Server (i.e. 'EVAL' LTSB/C) cannot be activated.
        - KMS38 only supports Windows/server version 14393 (1607) and newer versions.

====================================================================================================
   Switches in the Script:
====================================================================================================

 - For unattended mode, run the script with /u parameter.

"KMS38_Activation.cmd" /u

====================================================================================================
   File Details:
====================================================================================================

   fabb5a0fc1e6a372219711152291339af36ed0b5 *gatherosstate.exe                  Virus Total = 0/68
   ca3a51fdfc8749b8be85f7904b1c238a6dfba135 *slc.dll                            Virus Total = 1/67
   578364cb2319da7999acd8c015b4ce8da8f1b282 *ARM64_gatherosstate.exe            Virus Total = 0/70
   5dbea3a580cf60391453a04a5c910a3ceca2b810 *ARM64_slc.dll                      Virus Total = 0/69
   48d928b1bec25a56fe896c430c2c034b7866aa7a *ClipUp.exe                         Virus Total = 0/67

   Virus Total Report Date: 08-12-2019

 - File Sources:

   - ClipUp.exe (Original):
     From Windows server 2016 x64 ISO

   - gatherosstate.exe (Original):
     From Windows 10 x86 14393 ADK

   - ARM64_gatherosstate.exe (Original):
     From Windows 10 ARM64 18362 ISO

   - ARM64_slc.dll and slc.dll:

     Original slshim
     https://github.com/vyvojar/slshim

     Improved by @mspaintmsi
     https://www.nsaneforums.com/topic/316668--/?do=findComment&comment=1497887
     https://gitlab.com/massgrave/massgrave

     Source code is included.
     slc.dll is based on Integrated_Patcher_2 method.
     It is currently in use in HWID/KMS38 Activation script.

____________________________________________________________________________________________________

     You can safely delete the following files if it's not required for you.

     ClipUp.exe - Required to KMS38 activate Server Cor and Acor editions.
     ARM64_gatherosstate.exe and ARM64_slc.dll - Required to activate ARM64 Arch Windows 10.

====================================================================================================
   Manual Activation Process:
====================================================================================================

 - Prerequisite:

   For Windows 10 / Server x86-x64 system, you need following files,
   48d928b1bec25a56fe896c430c2c034b7866aa7a *ClipUp.exe       
   fabb5a0fc1e6a372219711152291339af36ed0b5 *gatherosstate.exe
   ca3a51fdfc8749b8be85f7904b1c238a6dfba135 *slc.dll           
   * ClipUp.exe is only required to activate Server Cor and Acor editions.

   For Windows 10 ARM64 system, you need following files,
   578364cb2319da7999acd8c015b4ce8da8f1b282 *ARM64_gatherosstate.exe
   5dbea3a580cf60391453a04a5c910a3ceca2b810 *ARM64_slc.dll
   * Rename the ARM64 files to gatherosstate.exe and slc.dll respectively.

   Make a folder named 'Files' in C drive, [C:\Files] and copy the required files in that folder.

   -------------------------------------------------------------------------------------------------

           GVLK                      Windows 10 Editions          

   TX9XD-98N7V-6WMQ6-BX7FG-H8Q99     Core
   PVMJN-6DFY6-9CCP6-7BKTT-D3WVR     CoreCountrySpecific
   3KHY7-WNT83-DGQKR-F7HPR-844BM     CoreN
   7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH     CoreSingleLanguage
   NW6C2-QMPVW-D7KKK-3GKT6-VCFB2     Education
   2WH4N-8QGBV-H22JP-CT43Q-MDWWJ     EducationN
   NPPR9-FWDCX-D2C8J-H872K-2YT43     Enterprise
   YYVX9-NTFWV-6MDM3-9PT4T-4M68B     EnterpriseG
   44RPN-FTY23-9VTTB-MP9BX-T84FV     EnterpriseGN
   DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4     EnterpriseN
   DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ     EnterpriseS                                [LTSB 2016]
   M7XTQ-FN8P6-TTKYV-9D4CC-J462D     EnterpriseS                                [LTSC 2019]
   QFFDN-GRT3P-VKWWX-X7T3R-8B639     EnterpriseSN                               [LTSB 2016]
   92NFX-8DJQP-P6BBQ-THF9C-7CG2H     EnterpriseSN                               [LTSC 2019]
   W269N-WFGWX-YVC9B-4J6C9-T83GX     Professional
   6TP4R-GNPTD-KYYHQ-7B7DP-J447Y     ProfessionalEducation
   YVWGF-BXNMC-HTQYQ-CPQ99-66QFC     ProfessionalEducationN
   MH37W-N47XK-V7XM9-C7227-GCQG9     ProfessionalN
   NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J     ProfessionalWorkstation
   9FNHH-K3HBT-3W4TD-6383H-6XYWF     ProfessionalWorkstationN
   7NBT4-WGBQX-MP4H7-QXFF8-YP3KX     ServerRdsh                            [Less than 1809]
   CPWHC-NT2C7-VYW78-DHDB2-PG3GK     ServerRdsh                     [Greater or Equal 1809]
   
           GVLK                      Windows Server Editions    
   
   QN4C6-GBJD2-FB422-GHWJK-GJG2R     ServerCloudStorage                       [Server 2016]
   CB7KF-BWN84-R7R2Y-793K2-8XDDG     ServerDatacenter, ServerDatacenterCor    [Server 2016]
   WMDGN-G9PQG-XVVXX-R3X43-63DFG     ServerDatacenter, ServerDatacenterCor    [Server 2019]
   JCKRF-N37P4-C2D82-9YXRT-4M63B     ServerSolution, ServerSolutionCor        [Server 2016]
   WVDHN-86M7X-466P6-VHXV7-YY726     ServerSolution, ServerSolutionCor        [Server 2019]
   WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY     ServerStandard, ServerStandardCor        [Server 2016]
   N69G4-B89J2-4G8F4-WWYCC-J464C     ServerStandard, ServerStandardCor        [Server 2019]
   VP34G-4NPPG-79JTQ-864T4-R3MQX     ServerAzureCor                           [Server 2016]
   FDNH6-VW9RW-BXPJ7-4XTYG-239TB     ServerAzureCor                           [Server 2019]
   6Y6KB-N82V8-D8CQV-23MJW-BWTG6     ServerDatacenterACor           [Server 1709 and later]
   DPCNP-XQFKJ-BJF7R-FRC8D-GF6G4     ServerStandardACor             [Server 1709 and later]

   -------------------------------------------------------------------------------------------------

 - Make sure to properly and completely remove any previously-installed KMS activator if one already exists.
 - Open CMD as Admin, and enter the following listed commands in the sequence in which they are given.
 - Enter Generic Volume License Key (GVLK) (Replace '%key%' with the key from the above list) 
   with the following command:
   
cscript /nologo %windir%\system32\slmgr.vbs /ipk %key%

 - Set specific KMS host to 127.0.0.2 [Localhost] with the following command: (Run one by one)
   - By doing this, the global KMS IP can not replace the KMS38 activation, and can then safely be used with MS Office 
     and other Windows Editions.
   - It's optional.

set spp=SoftwareLicensingProduct
for /f "tokens=2 delims==" %G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%' and Name like 'Windows%%' and PartialProductKey is not NULL) get ID /VALUE"') do (set app=%G)
wmic path %spp% where ID='%app%' call ClearKeyManagementServiceMachine
wmic path %spp% where ID='%app%' call ClearKeyManagementServicePort
wmic path %spp% where ID='%app%' call SetKeyManagementServiceMachine MachineName="127.0.0.2"
wmic path %spp% where ID='%app%' call SetKeyManagementServicePort 1688

 - Make sure slc.dll and gatherosstate.exe files are located in the folder, "C:\Files" and enter 
   following command to generate GenuineTicket.xml file.

call "C:\Files\gatherosstate.exe"

 - Now a GenuineTicket.xml file should be created in the folder "C:\Files\", copy and paste this file in the 
   folder named, "C:\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\"

 - Now apply this ticket using the following commands in this sequence:
   (In case of server cor and acor editions, copy the clipup.exe file to the folder "C:\Windows\System32\")

net stop ClipSVC
net start ClipSVC

 - Check the expiry date of the activation with the following command: 

cscript /nologo %windir%\system32\slmgr.vbs /xpr

 - If the expiry date is not in the year 2038, then enter the following command: 

cscript /nologo %windir%\system32\slmgr.vbs /rearm-app 55c92734-d682-4d71-983e-d6ec3f16059f
set spp=SoftwareLicensingProduct
for /f "tokens=2 delims==" %G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%' and Name like 'Windows%%' and PartialProductKey is not NULL) get ID /VALUE"') do (set app=%G)
cscript /nologo %windir%\system32\slmgr.vbs /rearm-sku %app%

 - check expiry date again, now it should show activation until the year 2038.

 - Done.

====================================================================================================
   Troubleshoot activation issues:
====================================================================================================

 - Make sure to completely remove any previously-installed KMS activators if any exist, before 
   installing KMS38 activation.

 - Reboot the system.

 - Now run the script to activate Windows 10, and if unsuccessful, 
   Try the troubleshoot button in settings activation page.
   If still unsuccessful then read additional troubleshoot options listed below.

--------------------------------------------

   - Open CMD as Admin, and enter the following command:

Dism /online /Cleanup-Image /RestoreHealth

   - After its done, reboot the system and Open CMD as Admin, and enter the following command:

sfc.exe /scannow

   - After it's done, reboot the system and run the activation script, and if unsuccessful, 
     open CMD as administrator again, and enter the following command:

slmgr.vbs /rearm

   - Reboot the system (important) and run the activation script, and if unsuccessful, 
     You may try to rebuild licensing Tokens.dat as suggested in https://support.microsoft.com/en-us/help/2736303
     (this will require to repair Office afterwards.)

   - Reboot the system and run the activation script, and if unsuccessful, 
     try cleaning the clipup using the following commands, it will reset all the HWID/KMS38 installed
     licences in the current system installation. open CMD as administrator again, and enter the 
     following commands one by one:

net stop ClipSVC
rundll32 clipc.dll,ClipCleanUpState

   - Reboot the system (important) and run the activation script, and if unsuccessful, it may be 
     time to start over from the very beginning and do a clean install of windows :D 

----------------------------------------------------------------------------------------------------

   - Some machines are not able to generate GenuineTicket.xml file using gatherosstate.exe
     The reason is unknown (to me). Please contact me if it happens to you.

=========================================================================================================
   Credits:
=========================================================================================================

   @mspaintmsi   Original co-authors of HWID/KMS38 Activation without KMS or predecessor install/upgrade.
      and        Created various methods for HWID/KMS38 Activation
   *Anonymous    https://www.nsaneforums.com/topic/316668--/?do=findComment&comment=1497887
                 https://gitlab.com/massgrave/massgrave

   @vyvojar      Original slshim (slc.dll)
                 https://github.com/vyvojar/slshim/releases

---------------------------------------------------------------------------------------------------------

   HWID/KMS38 methods Suggestions and improvements:-
  
   @sponpa       New ideas for the HWID/KM38 Generation
                 https://www.nsaneforums.com/topic/316668--/page/21/?tab=comments#comment-1431257

   @leitek8      Improvements for the slc.dll
                 https://www.nsaneforums.com/topic/316668--/page/22/?tab=comments#comment-1438005

---------------------------------------------------------------------------------------------------------

   Kind Help:-

   Thanks for having my back and answering all of my queries. (In no particular order)
   
   @AveYo aka @BAU, @sponpa, @mspaintmsi @RPO, @leitek8, @mxman2k, @Yen, @abbodi1406

   @BorrowedWifi for providing support in fixing English grammar errors in the Read Me.
   @Chibi ANUBIS and @smashed for testing scripts for ARM64 system.

   Special thanks to @abbodi1406 for providing the great help.

---------------------------------------------------------------------------------------------------------

   This script is a part of 'Microsoft Activation Scripts' project.

   Homepages-
   NsaneForums: (Login Required) https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/
   GitLab: https://gitlab.com/massgrave/microsoft-activation-scripts

   Maintained by @WindowsAddict

   P.S. I (@WindowsAddict) did not help in the development of HWID/KMS38 Activation in any way, I only 
   manage batch script tool which is based on the above mentioned original co-authors activation methods.

=========================================================================================================