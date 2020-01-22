====================================================================================================

   Online KMS Activation script is just a fork of @abbodi1406's KMS_VL_ALL Project.
   KMS_VL_ALL homepage: https://forums.mydigitallife.net/posts/838808

   This fork was made to avoid having any KMS binary files and system can be activated using 
   some manual commands or transparent batch script files.

   This script is a part of 'Microsoft Activation Scripts' project.

   Homepages-
   NsaneForums: (Login Required) https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/
   GitHub: https://github.com/massgravel/Microsoft-Activation-Scripts
   GitLab: https://gitlab.com/massgrave/microsoft-activation-scripts

   Maintained by @WindowsAddict

====================================================================================================
   Online KMS Activation:
====================================================================================================

 - This KMS script skips the activation of any permanently / KMS38-activated product.
 - This KMS script can activate both Retail and VL Windows / Server installation.
 - This KMS script can activate C2R Retail and VL Office, but not 2010/2013 MSI Retail Office. 

 ----------------------
 - Activation Renewal
 ----------------------

 - KMS activates Windows / Server / Office for 180 Days. (For Core/ProWMC edition it is 30/45 Days)

 - By design, using the KMS activation method, the system contacts the registered server every 7 
   days, and if contacted successfully it will automatically renew and reset the activation for the 
   the full period of 180 days again, starting from the day of successful contact.
   If the system cannot contact the server, it will be deactivated after 180 days and it will 
   remain deactivated until contact can be restored.
   
 - The KMS servers I've added have been working steadily for two to three years, but there can be 
   no guarantee that they will remain online indefinitely. If a registered server goes 
   down, you will need to create a renewal task, or do a manual renewal, for the lifetime of the 
   activation.

   There are 3 ways you can renew the KMS server and as a result, renew the activation.

 1- Activate.cmd

   Run this file whenever the system needs activation. Depending upon the particular (never fully-knowable)
   circumstances, a successful activation may last for a period of a MINIMUM of 180 days, 
   or a maximum of the full life of the machine it's running on, and you may never need to run it again.

 2- Manual Renewal via Desktop Context Menu

   This method is exactly same as above but here we put the following files in,
   C:\ProgramData\Online_KMS_Activation\BIN\cleanosppx64.exe
   C:\ProgramData\Online_KMS_Activation\BIN\cleanosppx86.exe
   C:\ProgramData\Online_KMS_Activation\Activate.cmd
   C:\ProgramData\Online_KMS_Activation\Info.txt

   and create registry entries in,
   HKCR\DesktopBackground\shell\Activate Windows - Office
   HKCR\DesktopBackground\shell\Activate Windows - Office\command

   It creates an easy to reach Desktop context menu for the manual activation renewal.

 3- Automatic Renewal via Task Scheduler
   
   This method put the following files in,
   C:\ProgramData\Online_KMS_Activation\BIN\cleanosppx64.exe
   C:\ProgramData\Online_KMS_Activation\BIN\cleanosppx86.exe
   C:\ProgramData\Online_KMS_Activation\Activate.cmd
   C:\ProgramData\Online_KMS_Activation\Info.txt

   And creates a scheduled task to run the script every 7 days.

   The scheduled task runs only if the system is connected to the Internet.
   With this method, the Activation task can also be created which will run on the system login
   and after successful activation, this task will delete itself.
  
   IMPORTANT NOTE - Some sensitive AV's may flag the Automatic Renewal via the Task, and not 
   because of KMS, because for them it's suspicious to run long scripts in the background as Tasks.

   It's recommended to set exclusions in Antivirus for
   C:\ProgramData\Online_KMS_Activation\Activate.cmd
   or
   use the 1st or 2nd option for activation renewal.

----------------------------------------------------------------------------------------------------

 ----------------------
 - Remarks
 ----------------------

 - This Online KMS Activation provides immediate global activation for Windows 8.1 and Windows 10, which 
   means that in the following three scenarios, the system will self-activate when connected to the 
   internet, and also means that users will not need to manually run the activation script again.

   Scenario 1: Subsequent installation or alteration of any 2013, 2016, or 2019 Volume License 
               (VL) Office product.
   Scenario 2: Windows edition change (with GVLK).
   Scenario 3: Date change, system hardware change, etc.

 - What is left in the system in the activation process?
   - Activate.cmd
     After activation, it leaves only the KMS Server name in the registry, which helps you to get the
     above-mentioned global activation feature whereby the system auto-renews the activations,
     so it's a good thing if you leave the server name in the registry.
     However, you can clear this registered KMS Server name upon activation, and do that, open 
     the script with notepad and set Clear-KMS-Cache to 1 from 0.
     What is left in the system when Renewal methods are installed, has been mentioned.

 - This script includes the most-stable KMS servers (6+) list. The server selection process is 
   fully automatic. You don't need to worry about the server's availability.

 - If your system date is incorrect (beyond 180 days) and you are offline, the system will be 
   deactivated, but will automatically reactivate when you correct the system date. 

 - Why should you choose the Online KMS activation method over offline KMS?
   The main benefit of Online KMS activation is that it doesn't need any KMS binary file and system
   can be activated using some manual commands or transparent batch script files.
   So this is for those who don't like/have difficulties/trust issue in offline KMS because of its 
   binary files and antivirus detections.

   If you prefer offline KMS then checkout an open-source activator, 
   @abbodi1406's KMS_VL_ALL   https://forums.mydigitallife.net/posts/838808

----------------------------------------------------------------------------------------------------

 --------------------------------------
 - Office C2R Retail to VL conversion
 --------------------------------------

   This activation script will convert Office C2R Retail to Volume without needing separate tools.
   
   - Supports: Office 365, Office 2019, Office 2016, Office 2013
   - Activated Retail products will be skipped from conversion
     this includes valid Office 365 subscriptions, or perpetual Office (MAK, OEM, MSDN, Retail..)
   - Current Office licenses will be cleaned up (unless retail-activated Office detected)
     then, proper Volume licenses will be installed based on the detected Product IDs
   - Office Mondo suite cover all products, if detected, only its licenses will be installed
   - Office 365 products will be converted with Mondo licenses by default  
     also, corresponding Office 365 Retail Grace Key will be installed
   - Office 2016 products will be converted with corresponding Office 2019 licenses
   - Office Professional suite will be converted with Office 2019 ProPlus licenses
   - Office HomeBusiness/HomeStudent suites will be converted with Office 2019 Standard licenses
   - If Office 2019 RTM licenses are not detected, Office 2016 licenses will be used instead
   - Office 2013 products follow the same logic but handled separately
   - If main products SKUs are detected, single apps licenses will not be installed to avoid duplication
   
   - SKUs:  
   O365ProPlus, O365Business, O365SmallBusPrem, O365HomePrem, O365EduCloud  
   ProPlus, Professional, Standard, HomeBusiness, HomeStudent, Visio, Project
   
   * Apps:  
   Access, Excel, InfoPath, Onenote, Outlook, PowerPoint, Publisher, SkypeForBusiness, Word, 
   Groove (OneDrive for Business)
   
   - O365ProPlus, O365Business, O365SmallBusPrem, ProPlus cover all apps  
   Professional cover all apps except SkypeForBusiness  
   Standard cover all apps except Access, SkypeForBusiness
   
   ## Notice
   
   - On Windows 7, Office 2016/2019 licensing service require Universal C Runtime to work correctly
   - UCRT is available in the latest Monthly Rollup, or the separate update KB3118401
   - Additionally, Office programs themselves require recent Windows 7 updates to start properly

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

   How does it work?
   Is it safe?

   https://pastebin.com/raw/7Xyaf15Z
   Mirror:
   https://textuploader.com/1dg8d/raw

====================================================================================================
   Products Compatibility:
====================================================================================================

   Supported Products: [Only Volume-capable]

   Windows 8 / 8.1 / 10 (all official editions, except Windows 10 S)  
   Windows 7 (Enterprise /N/E, Professional /N/E, Embedded Standard/POSReady/ThinPC)  
   Windows Server 2008 R2 / 2012 / 2012 R2 / 2016 / 2019  
   Office Volume 2010 / 2013 / 2016 / 2019
   
   [Activation script will convert Office C2R Retail O365/2013/2016/2019 to Volume]

----------------------------------------------------------------------------------------------------

   Unsupported Products:

   Office Retail [Office MSI Retail 2010/2013]
   Windows Editions which do not support KMS activation by design:  
   Windows Evaluation Editions  
   Windows 7 (Starter, HomeBasic, HomePremium, Ultimate)  
   Windows 10 (Cloud "S", IoTEnterprise, IoTEnterpriseS, ProfessionalSingleLanguage... etc)  
   Windows Server (Server Foundation, Storage Server, Home Server 2011... etc) 

----------------------------------------------------------------------------------------------------

   These editions are only KMS-activatable for 45 days at max:
   Windows 10 Home edition variants  
   Windows 8.1 Core edition variants, Pro with Media Center, Pro Student

   These editions are only KMS-activatable for 30 days at max:
   Windows 8 Core edition variants, Pro with Media Center

   Notes:  
   Supported Windows products do need volume conversion, only the GVLK (KMS key) is needed, which 
   the script will install accordingly.
   KMS Activation works in all (MBR, GPT, UEFI, BIOS) systems.

====================================================================================================
   Switches in the Script:
====================================================================================================

   * Applies to MAS Separate Files version

 - For unattended mode, run the scripts with /u parameter.

Activate.cmd /u
Uninstall.cmd /u

To create Renewal Task in unattended mode,
Renewal_Setup.cmd /rt

To create Renewal and Activation Task in unattended mode,
Renewal_Setup.cmd /rat

To create desktop context menu in unattended mode,
Renewal_Setup.cmd /dcm

To skip Windows activation in renewal script,
Renewal_Setup.cmd /rat /swa
or
Renewal_Setup.cmd /dcm /swa

 ----------------------------------------------------------------------

 - Activate.cmd

   - To run the script in Debug mode to find out the cause of errors in activation or just details,
     search "set _Debug=" and change the value from 0 to 1. 

   - To replace KMS38 activation with KMS activation, search "set SkipKMS38=" and change the value 
     from 1 to 0. 

   - To skip Windows activation, search "set ActWindows=" and change the value from 1 to 0.
   - To skip Office activation, search "set ActOffice=" and change the value from 1 to 0.
     - This is not effective if Windows and/or Office installation is already Volume (GVLK installed)
     - In [Online KMS + HWID] $OEM$ preactivation, Windows KMS activation is turned off 
       by default.

   - To turn OFF auto conversion for Office C2R Retail to Volume, search "set AutoR2V=" and 
     change the value from 1 to 0.

   - To set the script to use only one specific KMS server address, search "set KMS_Server="
     paste the server address after the = sign.

   - To clear the KMS cache, search "set Clear-KMS-Cache=" and change the value from 0 to 1.
     - Registered KMS server address (cache) enables the system to automatically renew the license 
       (for next 180 days) every 7 days, as long as the server is online.
     - This process is the same as how the legal KMS works, so no security program will flag 
       this behavior.
     - Changing this option here won't have any effect if manual (Desktop Context menu) and/or auto, 
       renewal activation script is installed. [default (0)].
     - I recommend leaving this option as default (0).

====================================================================================================
   Manual Activation Process:
====================================================================================================

 - Prerequisite:

   online Public KMS Server List: 

   kms.srv.crsoo.com
   kms.loli.beer
   kms8.MSGuides.com
   
   kms9.MSGuides.com
   kms.zhuxiaole.org
   kms.lolico.moe
   kms.moeclub.org

   Generic Volume License Key (GVLK):
   Thanks to @abbodi1406 for the Key collection.

        GVLK                        Edition                
   
   Windows 10
   
   TX9XD-98N7V-6WMQ6-BX7FG-H8Q99    Home
   3KHY7-WNT83-DGQKR-F7HPR-844BM    Home N
   7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH    Home Single Language
   PVMJN-6DFY6-9CCP6-7BKTT-D3WVR    Home China
   W269N-WFGWX-YVC9B-4J6C9-T83GX    Pro
   MH37W-N47XK-V7XM9-C7227-GCQG9    Pro N
   6TP4R-GNPTD-KYYHQ-7B7DP-J447Y    Pro Education
   YVWGF-BXNMC-HTQYQ-CPQ99-66QFC    Pro Education N
   NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J    Pro Workstation
   9FNHH-K3HBT-3W4TD-6383H-6XYWF    Pro Workstation N
   NW6C2-QMPVW-D7KKK-3GKT6-VCFB2    Education
   2WH4N-8QGBV-H22JP-CT43Q-MDWWJ    Education N
   NPPR9-FWDCX-D2C8J-H872K-2YT43    Enterprise
   DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4    Enterprise N
   YYVX9-NTFWV-6MDM3-9PT4T-4M68B    Enterprise G
   44RPN-FTY23-9VTTB-MP9BX-T84FV    Enterprise G N
   WNMTR-4C88C-JK8YV-HQ7T2-76DF9    Enterprise 2015 LTSB
   2F77B-TNFGY-69QQF-B8YKP-D69TJ    Enterprise 2015 LTSB N
   DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ    Enterprise 2016 LTSB
   QFFDN-GRT3P-VKWWX-X7T3R-8B639    Enterprise 2016 LTSB N
   M7XTQ-FN8P6-TTKYV-9D4CC-J462D    Enterprise LTSC 2019
   92NFX-8DJQP-P6BBQ-THF9C-7CG2H    Enterprise LTSC 2019 N
   CPWHC-NT2C7-VYW78-DHDB2-PG3GK    Enterprise for Virtual Desktops
   7NBT4-WGBQX-MP4H7-QXFF8-YP3KX    Remote Server
   NBTWJ-3DR69-3C4V8-C26MC-GQ9M6    Lean
   
   Windows 8.1
   
   M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK    Core
   7B9N3-D94CG-YTVHR-QBPX3-RJP64    Core N
   BB6NG-PQ82V-VRDPW-8XVD2-V8P66    Core Single Language
   NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3    Core China
   XYTND-K6QKT-K2MRH-66RTM-43JKP    Core ARM
   GCRJD-8NW9H-F2CDX-CCM8D-9D6T9    Pro
   HMCNV-VVBFX-7HMBH-CTY9B-B4FXY    Pro N
   789NJ-TQK6T-6XTH8-J39CJ-J8D3P    Pro with Media Center
   MHF9N-XY6XB-WVXMC-BTDCT-MKKG7    Enterprise
   TT4HM-HN7YT-62K67-RGRQJ-JFFXW    Enterprise N
   NMMPB-38DD4-R2823-62W8D-VXKJB    Embedded Industry Pro
   FNFKF-PWTVT-9RC8H-32HB2-JB34X    Embedded Industry Enterprise
   VHXM3-NR6FT-RY6RT-CK882-KW2CJ    Embedded Industry Automotive
   3PY8R-QHNP9-W7XQD-G6DPH-3J2C9    with Bing
   Q6HTR-N24GM-PMJFP-69CD8-2GXKR    with Bing N
   KF37N-VDV38-GRRTV-XH8X6-6F3BB    with Bing Single Language
   R962J-37N87-9VVK2-WJ74P-XTMHR    with Bing China
   MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B    Pro for Students
   TNFGH-2R6PB-8XM3K-QYHX2-J4296    Pro for Students N
   
   Windows 8
   
   BN3D2-R7TKB-3YPBD-8DRP2-27GG4    Core
   8N2M2-HWPGY-7PGT9-HGDD8-GVGGY    Core N
   2WN2H-YGCQR-KFX6K-CD6TF-84YXQ    Core Single Language
   4K36P-JN4VD-GDC6V-KDT89-DYFKP    Core China
   DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV    Core ARM
   NG4HW-VH26C-733KW-K6F98-J8CK4    Pro
   XCVCF-2NXM9-723PB-MHCB7-2RYQQ    Pro N
   GNBB8-YVD74-QJHX6-27H4K-8QHDG    Pro with Media Center
   32JNW-9KQ84-P47T8-D8GGY-CWCK7    Enterprise
   JMNMF-RHW7P-DMY6X-RF3DR-X2BQT    Enterprise N
   RYXVT-BNQG7-VD29F-DBMRY-HT73M    Embedded Industry Pro
   NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2    Embedded Industry Enterprise
   
   Windows 7
   
   FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4    Professional
   MRPKT-YTG23-K7D7T-X2JMM-QY7MG    Professional N
   W82YF-2Q76Y-63HXB-FGJG9-GF7QX    Professional E
   33PXH-7Y6KF-2VJC9-XBBR8-HVTHH    Enterprise
   YDRBP-3D83W-TY26F-D46B2-XCKRJ    Enterprise N
   C29WB-22CC8-VJ326-GHFJW-H9DH4    Enterprise E
   YBYF6-BHCR3-JPKRB-CDW7B-F9BK4    Embedded POSReady 7
   XGY72-BRBBT-FF8MH-2GG8H-W7KCW    Embedded Standard
   73KQT-CD9G6-K7TQG-66MRP-CQ22C    Embedded ThinPC
   
   Windows Server 2019
   
   N69G4-B89J2-4G8F4-WWYCC-J464C    Standard
   WMDGN-G9PQG-XVVXX-R3X43-63DFG    Datacenter
   WVDHN-86M7X-466P6-VHXV7-YY726    Essentials
   FDNH6-VW9RW-BXPJ7-4XTYG-239TB    Azure Core
   N2KJX-J94YW-TQVFB-DG9YT-724CC    Standard ACor
   6NMRW-2C8FM-D24W7-TQWMY-CWH2D    Datacenter ACor
   GRFBW-QNDC4-6QBHG-CCK3B-2PR88    ServerARM64
   
   Windows Server 2016
   
   WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY    Standard
   CB7KF-BWN84-R7R2Y-793K2-8XDDG    Datacenter
   JCKRF-N37P4-C2D82-9YXRT-4M63B    Essentials
   QN4C6-GBJD2-FB422-GHWJK-GJG2R    Cloud Storage
   VP34G-4NPPG-79JTQ-864T4-R3MQX    Azure Core
   PTXN8-JFHJM-4WC78-MPCBR-9W4KR    Standard ACor
   2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG    Datacenter ACor
   K9FYF-G6NCK-73M32-XMVPY-F9DRR    ServerARM64
   
   Windows Server 2012 R2
   
   D2N9P-3P6X9-2R39C-7RTCD-MDVJX    Standard
   W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9    Datacenter
   KNC87-3J2TX-XB4WP-VCPJV-M4FWM    Essentials
   3NPTF-33KPT-GGBPR-YX76B-39KDD    Cloud Storage
   
   Windows Server 2012
   
   XC9B7-NBPP2-83J2H-RHMBY-92BT4    Standard
   48HP8-DN98B-MYWDG-T2DCC-8W83P    Datacenter
   HM7DN-YVMH3-46JC3-XYTG7-CYQJJ    MultiPoint Standard
   XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G    MultiPoint Premium
   
   Windows Server 2008 R2
   
   6TPJF-RBVHG-WBW2R-86QPH-6RTM4    Web
   TT8MH-CG224-D3D7Q-498W2-9QCTX    HPC
   YC6KT-GKW9T-YTKYR-T4X34-R7VHC    Standard
   74YFP-3QFB3-KQT8W-PMXWJ-7M648    Datacenter
   489J6-VHDMP-X63PK-3K798-CPX3Y    Enterprise
   GT63C-RJFQ3-4GMB6-BRFB9-CB83V    Itanium
   736RG-XDKJK-V34PF-BHK87-J6X3K    MultiPoint Server
     
   ----------------------------------------------------------------------------------------------------
   
   ----------------------------------------------
    Windows /Server (All VL Supported Versions) 
   ----------------------------------------------
   
 - Connect to the internet. 
 - Open CMD as Admin, and enter the following listed commands in the sequence in which they are given.
 - Enter Generic Volume License Key (GVLK) (Replace %key% with the key from above list) with 
   the following command:

slmgr.vbs /ipk %key%

 - Register the KMS Server, (Replace %server% with one of the above-listed servers) 
   (If activation is unsuccessful then try a different server) with the following command:

slmgr.vbs /skms %server%

 - Activate Windows with the following command:
   
slmgr.vbs /ato
   
 - Check Activation Status with the following command:

slmgr.vbs /dli
   
 - Check Activation Expiry Date with the following command:
   
slmgr.vbs /xpr
   
 - Clear the name of the KMS server (Optional) (It'll prevent activation auto-renewal) with the following command: 
   
slmgr.vbs /ckms
   
 - Done. 

   ----------------------------------------------------------------------------
    Office VL Activation (Office 2010, 2013, 2016, 2019) -
   ----------------------------------------------------------------------------
   
 - Connect to the internet. 
 - Open CMD as Admin, and enter the following listed commands in the sequence in which they are given.
 - If Office is installed as VL (Volume License) then there is no need to enter its key. 
 - If Office is installed as Retail, then you need to convert it to VL, by using C2R-R2V by @abbodi1406
   https://forums.mydigitallife.net/posts/1150042 

 - Change to the directory where Office is installed.
   If your system is 32-bit Office on 32-bit Windows or 64-bit Office on 64-bit Windows use the following:
   
   For Office 2016 or 2019 enter the command:
   
cd "C:\Program Files\Microsoft Office\Office16"
   
   For Office 2013 enter the command:
   
cd "C:\Program Files\Microsoft Office\Office15"
   
   For Office 2010 enter the command:
   
cd "C:\Program Files\Microsoft Office\Office14"
   
   --------------------------------------------------------------------------------
   
   If your system is 32-bit Office on 64-bit Windows, use the following:
   
   For Office 2016 or 2019 enter the command:
   
cd "C:\Program Files (x86)\Microsoft Office\Office16"
   
   For Office 2013 enter the command:
   
cd "C:\Program Files (x86)\Microsoft Office\Office15"
   
   For Office 2010 enter the command:
   
cd "C:\Program Files (x86)\Microsoft Office\Office14"
   
   --------------------------------------------------------------------------------
   
 - Once all of that is done correctly, you must register the KMS Server, (In the following,
   replace %server% with one of the above-listed servers.) (If activation is unsuccessful 
   then try a different server.) with the following command:

cscript ospp.vbs /sethst:%server% 

 - Activate Office with the following command: 
   
cscript ospp.vbs /act 
   
 - Check Activation Status with the following command:
   
cscript ospp.vbs /dstatus
 
 - Clear the name of the KMS server, (Optional) (It'll prevent activation auto-renewal)
   with the appropriate following commands:
   
   To clear the KMS Server name for Office in Win 7, or Office 2010 on Win 8 or Win 10,
   enter each of the following commands in the sequence which is given:
   
set "OSPP=HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
reg delete "%OSPP%" /f /v KeyManagementServiceName 2>nul
reg delete "%OSPP%" /f /v KeyManagementServicePort 2>nul
reg delete "%OSPP%\59a52881-a989-479d-af46-f275c6370663" /f 2>nul
reg delete "%OSPP%\0ff1ce15-a989-479d-af46-f275c6370663" /f 2>nul

   To clear the KMS server name for Office (except Office 2010) on Win 8 or Win 10, run the following command:

set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName 2>nul
reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort 2>nul
reg delete "HKLM\%SPPk%\0ff1ce15-a989-479d-af46-f275c6370663" /f 2>nul
reg delete "HKEY_USERS\S-1-5-20\%SPPk%\0ff1ce15-a989-479d-af46-f275c6370663" /f 2>nul

 - Done. 

====================================================================================================
   Troubleshoot activation issues:
====================================================================================================

 - Make sure the Internet is connected.

 - Reboot the system and run the activation script, and if unsuccessful,
   Open CMD as Admin, and enter the following command: (For Windows 10\8\8.1)

Dism /online /Cleanup-Image /RestoreHealth

   - After its done, reboot the system and open CMD as Admin and enter the following command:

sfc.exe /scannow

   - After it's done, reboot the system and run the activation script, and if unsuccessful, 
     open CMD as administrator again, and enter the following command:

slmgr.vbs /rearm

   - Reboot the system and run the activation script, and if unsuccessful, 
     You may try to rebuild licensing Tokens.dat as suggested in https://support.microsoft.com/en-us/help/2736303
     (this will require to repair Office afterwards.)

   - Reboot the system and run the activation script, and if unsuccessful, 
     Try KMS_VL_ALL by @abbodi1406 https://forums.mydigitallife.net/posts/838808/
     If still unsuccessful, it may be time to start over from the very beginning 
     and do a clean install of windows :D 

   -------------------------------

   - If you have issues with Office activation, or got undesired or duplicate licenses (e.g. Office 2016 and 2019):
     Download Office Scrubber pack from https://forums.mydigitallife.net/posts/1466365/
     To get rid of any conflicted licenses, run Uninstall_Licenses.cmd, then you must start any 
     Office program to repair the licensing. You may also try Uninstall_Keys.cmd for similar manner.

     If you wish to remove Office and leftovers completely and start clean:
     Uninstall Office normally from Control Panel / Programs and Feature then run Full_Scrub.cmd
     afterwards, install new Office.

   - Can't activate Windows 7 with KMS: [Error 0xC004F035]
     Some OEM licensed computers cannot be activated with KMS on WINDOWS 7.
     Quote from the MS page https://tinyurl.com/yy8wfu5m
     'Computers obtained through OEM channels that have an ACPI_SLIC table in the (BIOS) are 
     required to have a valid Windows marker in the same ACPI_SLIC table.
     ---Computers that have an ACPI_SLIC table without a valid Windows marker generate an error 
     when a volume edition of Windows 7 is installed.'

====================================================================================================
   Credits:
====================================================================================================

   @abbodi1406   Activate.cmd (KMS_VL_ALL)
                 https://forums.mydigitallife.net/posts/838808
                 (* With the great help from @RPO, Forked it to work with Multi KMS Servers,
                 Renewal task, Desktop context menu, $OEM$, etc for Online KMS)

                 Clear-KMS-Cache.cmd
                 https://forums.mydigitallife.net/posts/1511883
                 (*Applied it as it is)

                 Check-Activation-Status-wmic.cmd
                 https://forums.mydigitallife.net/posts/838808
                 (*Applied it as it is)

----------------------------------------------------------------------------------------------------

   Online Public KMS Servers:

   kms.srv.crsoo.com
   kms.loli.beer
   kms8.MSGuides.com

   kms9.MSGuides.com
   kms.zhuxiaole.org
   kms.lolico.moe
   kms.moeclub.org

----------------------------------------------------------------------------------------------------

   Special Thanks to @RPO

   For providing great support in making and improvements of this script,
   To name a few,

   Internet test with Powershell (No ping)
   KMS server 1688 port test with Powershell
   Multi KMS server integration
   Scheduled task to renew the activation
                    
   And for solving countless problems in this batch script.
   
----------------------------------------------------------------------------------------------------

   Kind Help:-

   Thanks for having my back and answering all of my queries. (In no particular order)

   @AveYo aka @BAU, @RPO, @leitek8, @mxman2k, @Yen, @abbodi1406

   @BorrowedWifi For providing support in fixing English grammar errors in the Read Me.

----------------------------------------------------------------------------------------------------

   This script is a part of 'Microsoft Activation Scripts' project.

   Homepages-
   NsaneForums: (Login Required) https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/
   GitHub: https://github.com/massgravel/Microsoft-Activation-Scripts
   GitLab: https://gitlab.com/massgrave/microsoft-activation-scripts

   Maintained by @WindowsAddict

====================================================================================================