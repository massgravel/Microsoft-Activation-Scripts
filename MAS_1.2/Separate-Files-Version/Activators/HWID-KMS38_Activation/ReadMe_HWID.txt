====================================================================================================
   [HWID] Digital License Activation:
====================================================================================================

 - This activation is supported for Windows 10 ONLY.
 - This activation does not store any files in your system.
 - This activation is a permanent activation for your system Hardware.

 - On a system, this activation can be created for all the supported editions, and all can be 
   linked to Microsoft account without any issues.

 - Once a system is activated, this activation cannot be removed. (Because the license is stored in 
   the Microsoft servers and not in the user's system.)

 - Any significant changes in the Hardware (such as a motherboard) may deactivate the system. 
   It is possible to reactivate a system that was deactivated because of significant hardware
   changes, IF your activation, was linked to an online Microsoft account.

 - For activation to succeed, the Windows Update Service and internet connectivity must be enabled.
   If you are trying to activate without these conditions being met, then the system will auto-
   activate later when the conditions are met.

 - Auto activation scenario after the Windows reinstall:
   - The Internet is required. (Only at the time of activation)
   - The system will auto-activate if Retail (Consumer) media was used for the installation.
   - The system will NOT auto-activate if VL (Business) media was used for the installation.
     In this case, the user will have to insert that windows edition Retail/OEM key (find keys below 
     in this page) to activate, if the user doesn't wish to activate again using this script.

 - Possible reasons for activation failure:
   - The Internet is not connected.
   - Windows update service is disabled.
   - Use of a VPN, and/or a privacy-based hosts file, firewall rules.
   - Corrupt system files.
   - Microsoft servers block the activation request from some countries such as Iran.
   - Rarely, Microsoft's activation servers are the problem.
   - Some machines are not able to generate GenuineTicket.xml file using gatherosstate.exe
     The reason is unknown (to me). Please contact me if it happens to you.

   * Troubleshoot guide is listed below.

====================================================================================================

   Documentation by @mspaintmsi

   Included topics-

   How does it work?
   Is it possible that Microsoft can block these Digital Licenses (HWID)?
   
   https://pastebin.com/raw/7Xyaf15Z
   Mirror Link-
   https://textuploader.com/1dg8d/raw

====================================================================================================
   Supported Products:
====================================================================================================

   Windows 10 Versions that can be activated:

   Core
   CoreCountrySpecific
   CoreN
   CoreSingleLanguage
   Education
   EducationN
   Enterprise
   EnterpriseN
   EnterpriseS    [LTSB 2015 & 2016]
   EnterpriseSN   [LTSB 2015 & 2016]
   Professional
   ProfessionalEducation
   ProfessionalEducationN
   ProfessionalN
   ProfessionalWorkstation
   ProfessionalWorkstationN
   ServerRdsh
   IoTEnterprise


   Note - X86-X64 and ARM64 architecture systems are supported.
        - Any Evaluation version of Windows (i.e. 'EVAL' LTSB/C) cannot be activated.
        - LTSC 2019 is not supported.

====================================================================================================
   Switches for the Script:
====================================================================================================

 - To run the script in unattended mode, use /u parameter.
"HWID_Activation.cmd" /u

====================================================================================================
   File Details:
====================================================================================================

   fabb5a0fc1e6a372219711152291339af36ed0b5 *gatherosstate.exe                  Virus Total = 0/68
   ca3a51fdfc8749b8be85f7904b1c238a6dfba135 *slc.dll                            Virus Total = 1/67
   578364cb2319da7999acd8c015b4ce8da8f1b282 *ARM64_gatherosstate.exe            Virus Total = 0/70
   5dbea3a580cf60391453a04a5c910a3ceca2b810 *ARM64_slc.dll                      Virus Total = 0/69

   Virus Total Report Date: 08-12-2019

 - File Sources:

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

     ARM64_gatherosstate.exe and ARM64_slc.dll - Required to activate ARM64 Arch Windows 10.

====================================================================================================
   Manual Activation Process:
====================================================================================================

 - Prerequisite:

   For Windows 10 x86-x64 system, you need following files,
   fabb5a0fc1e6a372219711152291339af36ed0b5 *gatherosstate.exe
   ca3a51fdfc8749b8be85f7904b1c238a6dfba135 *slc.dll

   For Windows 10 ARM64 system, you need following files,
   578364cb2319da7999acd8c015b4ce8da8f1b282 *ARM64_gatherosstate.exe
   5dbea3a580cf60391453a04a5c910a3ceca2b810 *ARM64_slc.dll
   * Rename the ARM64 files to gatherosstate.exe and slc.dll respectively.


   Make a folder named 'Files' in C drive, [C:\Files] and copy the required files in that folder.

   -------------------------------------------------------------------------------------------------

         Retail/OEM Keys            Windows 10 Editions

   YTMG3-N6DKC-DKB77-7M9GH-8HVX7    Core
   4CPRK-NM3K3-X6XXQ-RXX86-WXCHW    CoreN
   N2434-X9D7W-8PF6X-8DV9T-8TYMD    CoreCountrySpecific
   BT79Q-G7N6G-PGBYW-4YWX6-6F4BT    CoreSingleLanguage
   YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY    Education
   84NGF-MHBT6-FXBX8-QWJK7-DRR8H    EducationN
   XGVPP-NMH47-7TTHJ-W3FW7-8HV2C    Enterprise
   3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT    EnterpriseN
   FWN7H-PF93Q-4GGP8-M8RF3-MDWWW    EnterpriseS      [LTSB 2015]
   8V8WN-3GXBH-2TCMG-XHRX3-9766K    EnterpriseSN     [LTSB 2015]
   NK96Y-D9CD8-W44CQ-R8YTK-DYJWX    EnterpriseS      [LTSB 2016]
   2DBW3-N2PJG-MVHW3-G7TDK-9HKR4    EnterpriseSN     [LTSB 2016]
   VK7JG-NPHTM-C97JM-9MPGT-3V66T    Professional
   2B87N-8KFHP-DKV6R-Y2C8J-PKCKT    ProfessionalN
   8PTT6-RNW4C-6V7J2-C2D3X-MHBPB    ProfessionalEducation
   GJTYN-HDMQY-FRR76-HVGC7-QPF8P    ProfessionalEducationN
   DXG7C-N36C4-C4HTG-X4T3X-2YV77    ProfessionalWorkstation
   WYPNQ-8C467-V2W6J-TX4WX-WT2RQ    ProfessionalWorkstationN
   NJCF7-PW8QT-3324D-688JX-2YV66    ServerRdsh
   XQQYW-NFFMW-XJPBH-K8732-CKFFD    IoTEnterprise
   
   -------------------------------------------------------------------------------------------------

 - Make sure the Windows Update Service and internet are both enabled.
 - Open a command prompt (run cmd.exe) as administrator, and enter following listed commands in the 
   the sequence in which they are given.
 - Enter Retail/OEM Key, (Replace '%key%' with the key from the above list) with the following command:

cscript /nologo %windir%\system32\slmgr.vbs /ipk %key%

 - Make sure slc.dll and gatherosstate.exe files are located in the folder, "C:\Files" and enter 
   the following commands to generate GenuineTicket.xml file.

   For x86-x64 systems,

pushd "C:\Files"
rundll32 "C:\Files\slc.dll",PatchGatherosstate
call "C:\Files\gatherosstatemodified.exe"

   For ARM64 systems,

call "C:\Files\gatherosstate.exe"

 - Now a GenuineTicket.xml file should be created in the folder "C:\Files\", copy and paste this file in the 
   folder named, "C:\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\"

 - Now apply this ticket using the following commands in this sequence:

net stop ClipSVC
net start ClipSVC

 - Activate Windows with the following command:

cscript /nologo %windir%\system32\slmgr.vbs /ato

 - Check Activation Status with the following command:

cscript /nologo %windir%\system32\slmgr.vbs /xpr

 - Done.

   * Note - [clipup -v -o -altto <ticket_path>] method to apply the ticket was not suggested because
            of the issues in case the username have spaces or non English characters.

====================================================================================================
   Troubleshoot activation issues:
====================================================================================================

 - Make sure the internet is connected.
 
 - Open CMD and type services.msc and hit Enter, When Services opens up, look for 'Windows Update'
   and Make sure its startup type is set to Automatic. Some update blocking tools and scripts 
   usually permanently block the update service, you need to make sure it's not the case.

 - VPN, privacy-based hosts and/or firewall rules may cause problems with the activation. Disable 
   them if you are facing problems in activation.

 - Reboot the system.

 - Now run the script to activate Windows 10, and if unsuccessful, 
   Try the troubleshoot button in the settings activation page.
   If still unsuccessful then read additional troubleshoot options listed below.

--------------------------------------------

   - Open CMD as administrator, and enter the following command:

Dism /online /Cleanup-Image /RestoreHealth

   - After it's done, reboot the system and open CMD as administrator again, and enter the 
     following command:

sfc.exe /scannow

   - After it's done, reboot the system and run the activation script, and if unsuccessful, 
     open CMD as administrator again, and enter the following command:

slmgr.vbs /rearm

   - Reboot the system and run the activation script, and if unsuccessful, 
     You may try to rebuild licensing Tokens.dat as suggested in https://support.microsoft.com/en-us/help/2736303
     (this will require to repair Office afterwards.)

   - Reboot the system and run the activation script, and if unsuccessful, 
     try cleaning the clipup using the following commands, it will reset all the HWID/KMS38 installed
     licences in the current system installation. open CMD as administrator again, and enter the 
     following commands one by one:

net stop ClipSVC
rundll32 clipc.dll,ClipCleanUpState

   - Reboot the system (important) and run the activation script, and if unsuccessful, 
     Make sure hardware component proper drivers are installed, check manufacturer site/Windows-
     update for drivers.

   - After it's done, reboot the system and run the activation script, and if unsuccessful,
     it may be time to start over from the very beginning and do a clean install of windows :D 

-------------------------------------------
 Activation is blocked in some countries -
-------------------------------------------

 - Microsoft servers block the activation request from some countries such as Iran,
   To activate the system in those countries, follow the below steps,
   - In the settings app, Change Region and Timezone to the USA location and use a VPN, choose the
     the location of the USA. Now run the script, it should activate now.

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