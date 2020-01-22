====================================================================================================
   KMS38 Protection:
====================================================================================================

 - If you don't know what KMS38 is, then first check it in the Read Me.

 - By default, a KMS38 Activation is vulnerable to unintended overwriting/replacement and 
   neutralization by a 180-Day KMS Activator (non-KMS38 Activator).
   However, with a few tricks you can ensure that no alternative KMS Activator can replace KMS38 
   Activation by accident or even on purpose. This script demonstrate how to do/undo that.

 - Protect KMS38:
   - How does KMS38 Protection work?
     In the KMS activation method, the Windows Operating System first checks the KMS IP registered 
     as a specific KMS, and if that is not defined then it checks the Global KMS IP.
     Another fact is that if LocalHost (127.0.0.2) is defined as KMS IP in the Windows 8.1 and 10 OS's
     then Windows will not accept it as a valid KMS IP.
     This script simply utilizes the above facts to protect the KMS38 activation from being 
     overwritten by any alternative 'normal' 180-Day KMS Activation.

     Script steps-
     - Check if Windows is activated with KMS38, if yes,
     - Set that Windows edition specific KMS IP to LocalHost (127.0.0.2),
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f\XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
where X is Windows edition Activation ID.

     - Lock this Registry with Reg_takeownership snippet by @AveYo aka @BAU
       pastebin.com/XTPt0JSC
     - Done.

 - Unprotect KMS38:
   - Just undo above steps,
     - Give administrator full control of that mentioned registry key.
     - Delete that registry key.
     - Done.

=======================================================================================================

  This script is a part of 'Microsoft Activation Scripts' project.

  Homepages-
  NsaneForums: (Login Required) https://www.nsaneforums.com/topic/316668-microsoft-activation-scripts/
  GitHub: https://github.com/massgravel/Microsoft-Activation-Scripts
  GitLab: https://gitlab.com/massgrave/microsoft-activation-scripts

  Maintained by @WindowsAddict

=======================================================================================================