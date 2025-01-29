@echo off


::  Check-Activation-Status
::  Written by @abbodi1406
::  https://gravesoft.dev/cas


::  Set Environment variables, it helps if they are misconfigured in the system

setlocal EnableExtensions
setlocal DisableDelayedExpansion

set "PathExt=.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC"

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)

set "ComSpec=%SysPath%\cmd.exe"
set "PSModulePath=%ProgramFiles%\WindowsPowerShell\Modules;%SysPath%\WindowsPowerShell\v1.0\Modules"

set "_psc=powershell -nop -c"
set "_err===== ERROR ===="
set _pwsh=1
for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" set _pwsh=0
cmd /c "%_psc% "$ExecutionContext.SessionState.LanguageMode"" | find /i "FullLanguage" 1>nul || (set _pwsh=0)
if %_pwsh% equ 0 (
echo %_err%
cmd /c "%_psc% "$ExecutionContext.SessionState.LanguageMode""
echo Windows PowerShell is not working correctly.
echo It is required for this script to work.
goto :E_Exit
)
set "_batf=%~f0"
set "_batp=%_batf:'=''%"
setlocal EnableDelayedExpansion
%_psc% "$f=[IO.File]::ReadAllText('!_batp!') -split ':sppmgr\:.*';iex ($f[1])"

:E_Exit
echo.
echo Press 0 key to exit.
choice /c 0 /n
exit /b

:sppmgr:
param (
    [Parameter()]
    [switch]
    $All,
    [Parameter()]
    [switch]
    $Dlv,
    [Parameter()]
    [switch]
    $IID,
    [Parameter()]
    [switch]
    $Pass
)

function CONOUT($strObj)
{
	Out-Host -Input $strObj
}

function ExitScript($ExitCode = 0)
{
	Exit $ExitCode
}

if (-Not $PSVersionTable) {
	"==== ERROR ====`r`n"
	"Windows PowerShell 1.0 is not supported by this script."
	ExitScript 1
}

if ($ExecutionContext.SessionState.LanguageMode.value__ -NE 0) {
	"==== ERROR ====`r`n"
	"Windows PowerShell is not running in Full Language Mode."
	ExitScript 1
}

$winbuild = 1
try {
	$winbuild = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$env:SystemRoot\System32\kernel32.dll").FileBuildPart
} catch {
	$winbuild = [int]([wmi]'Win32_OperatingSystem=@').BuildNumber
}

if ($winbuild -EQ 1) {
	"==== ERROR ====`r`n"
	"Could not detect Windows build."
	ExitScript 1
}

if ($winbuild -LT 2600) {
	"==== ERROR ====`r`n"
	"This build of Windows is not supported by this script."
	ExitScript 1
}

$SysPath = "$env:SystemRoot\System32"
if (Test-Path "$env:SystemRoot\Sysnative\reg.exe") {
	$SysPath = "$env:SystemRoot\Sysnative"
}

if (Test-Path "$SysPath\sppc.dll") {
	$SLdll = 'sppc.dll'
} elseif (Test-Path "$SysPath\slc.dll") {
	$SLdll = 'slc.dll'
} else {
	"==== ERROR ====`r`n"
	"Software Licensing Client Dll is not detected."
	ExitScript 1
}

if ($All.IsPresent)
{
	$isAll = {CONOUT "`r"}
	$noAll = {$null}
}
else
{
	$isAll = {$null}
	$noAll = {CONOUT "`r"}
}
$Dlv = $Dlv.IsPresent
$IID = $IID.IsPresent -Or $Dlv.IsPresent

$NT6 = $winbuild -GE 6000
$NT7 = $winbuild -GE 7600
$NT9 = $winbuild -GE 9600

$Admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$line2 = "============================================================"
$line3 = "____________________________________________________________"

function UnQuickEdit
{
	$t=[AppDomain]::CurrentDomain.DefineDynamicAssembly((Get-Random), 1).DefineDynamicModule((Get-Random), $False).DefineType((Get-Random))
	$t.DefinePInvokeMethod('GetStdHandle', 'kernel32.dll', 22, 1, [IntPtr], @([Int32]), 1, 3).SetImplementationFlags(128)
	$t.DefinePInvokeMethod('SetConsoleMode', 'kernel32.dll', 22, 1, [Boolean], @([IntPtr], [Int32]), 1, 3).SetImplementationFlags(128)
	$t.DefinePInvokeMethod('GetConsoleWindow', 'kernel32.dll', 22, 1, [IntPtr], @(), 1, 3).SetImplementationFlags(128)
	$t.DefinePInvokeMethod('SendMessageW', 'user32.dll', 22, 1, [IntPtr], @([IntPtr], [UInt32], [IntPtr], [IntPtr]), 1, 3).SetImplementationFlags(128)
	$k=$t.CreateType()
	if ($winbuild -GE 17763) {
		if ($k::SendMessageW($k::GetConsoleWindow(), 127, 0, 0) -EQ [IntPtr]::Zero) {
			return
		}
	}
	$v=(0x0080, 0x00A0)[!($winbuild -GE 10586)]
	$b=$k::SetConsoleMode($k::GetStdHandle(-10), $v)
}

function echoWindows
{
	CONOUT "$line2"
	CONOUT "===                   Windows Status                     ==="
	CONOUT "$line2"
	& $noAll
}

function echoOffice
{
	if ($doMSG -EQ 0) {
		return
	}

	& $isAll
	CONOUT "$line2"
	CONOUT "===                   Office Status                      ==="
	CONOUT "$line2"
	& $noAll

	$script:doMSG = 0
}

function strGetRegistry($strKey, $strName)
{
	try {
		return [Microsoft.Win32.Registry]::GetValue($strKey, $strName, $null)
	} catch {
		return $null
	}
}

function CheckOhook
{
	$ohook = 0
	$paths = "${env:ProgramFiles}", "${env:ProgramW6432}", "${env:ProgramFiles(x86)}"

	15, 16 | foreach `
	{
		$A = $_; $paths | foreach `
		{
			if (Test-Path "$($_)$('\Microsoft Office\Office')$($A)$('\sppc*dll')") {$ohook = 1}
		}
	}

	"System", "SystemX86" | foreach `
	{
		$A = $_; "Office 15", "Office" | foreach `
		{
			$B = $_; $paths | foreach `
			{
				if (Test-Path "$($_)$('\Microsoft ')$($B)$('\root\vfs\')$($A)$('\sppc*dll')") {$ohook = 1}
			}
		}
	}

	if ($ohook -EQ 0) {
		return
	}

	& $isAll
	CONOUT "$line2"
	CONOUT "===                Office Ohook Status                   ==="
	CONOUT "$line2"
	$host.UI.WriteLine('Yellow', 'Black', "`r`nOhook for permanent Office activation is installed.`r`nYou can ignore the below mentioned Office activation status.")
	& $noAll
}

#region WMI
function DetectID($strSLP, $strAppId)
{
	$ppk = (" AND PartialProductKey <> NULL)", ")")[$All.IsPresent]
	$fltr = "SELECT ID FROM $strSLP WHERE (ApplicationID='$strAppId'"
	$clause = $fltr + $ppk
	$sWmi = [wmisearcher]$clause
	$sWmi.Options.Rewindable = $false
	return ($sWmi.Get().Count -GT 0)
}

function GetID($strSLP, $strAppId)
{
	$NT5 = ($strSLP -EQ $wslp -And $winbuild -LT 6001)
	$IDs = [Collections.ArrayList]@()
	$isAdd = (" AND LicenseDependsOn <> NULL)", ")")[$NT5]
	$noAdd = " AND LicenseDependsOn IS NULL)"
	$query = "SELECT ID FROM $strSLP WHERE (ApplicationID='$strAppId' AND PartialProductKey"

	if ($All.IsPresent) {
		$fltr = $query + " IS NULL"
		$clause = $fltr + $isAdd
		$sWmi = [wmisearcher]$clause
		$sWmi.Options.Rewindable = $false
		try {$sWmi.Get() | select -Expand Properties -EA 0 | foreach {$IDs += $_.Value}} catch {}
		if (-Not $NT5) {
		$clause = $fltr + $noAdd
		$sWmi = [wmisearcher]$clause
		$sWmi.Options.Rewindable = $false
		try {$sWmi.Get() | select -Expand Properties -EA 0 | foreach {$IDs += $_.Value}} catch {}
		}
	}

	$fltr = $query + " <> NULL"
	$clause = $fltr + $isAdd
	$sWmi = [wmisearcher]$clause
	$sWmi.Options.Rewindable = $false
	try {$sWmi.Get() | select -Expand Properties -EA 0 | foreach {$IDs += $_.Value}} catch {}
	if (-Not $NT5) {
	$clause = $fltr + $noAdd
	$sWmi = [wmisearcher]$clause
	$sWmi.Options.Rewindable = $false
	try {$sWmi.Get() | select -Expand Properties -EA 0 | foreach {$IDs += $_.Value}} catch {}
	}

	return $IDs
}

function DetectSubscription {
	if ($null -EQ $objSvc.SubscriptionType -Or $objSvc.SubscriptionType -EQ 120) {
		return
	}

	if ($objSvc.SubscriptionType -EQ 1) {
		$SubMsgType = "Device based"
	} else {
		$SubMsgType = "User based"
	}

	if ($objSvc.SubscriptionStatus -EQ 120) {
		$SubMsgStatus = "Expired"
	} elseif ($objSvc.SubscriptionStatus -EQ 100) {
		$SubMsgStatus = "Disabled"
	} elseif ($objSvc.SubscriptionStatus -EQ 1) {
		$SubMsgStatus = "Active"
	} else {
		$SubMsgStatus = "Not active"
	}

	$SubMsgExpiry = "Unknown"
	if ($objSvc.SubscriptionExpiry) {
		if ($objSvc.SubscriptionExpiry.Contains("unspecified") -EQ $false) {$SubMsgExpiry = $objSvc.SubscriptionExpiry}
	}

	$SubMsgEdition = "Unknown"
	if ($objSvc.SubscriptionEdition) {
		if ($objSvc.SubscriptionEdition.Contains("UNKNOWN") -EQ $false) {$SubMsgEdition = $objSvc.SubscriptionEdition}
	}

	CONOUT "`nSubscription information:"
	CONOUT "    Edition: $SubMsgEdition"
	CONOUT "    Type   : $SubMsgType"
	CONOUT "    Status : $SubMsgStatus"
	CONOUT "    Expiry : $SubMsgExpiry"
}

function DetectAdbaClient
{
	CONOUT "`nAD Activation client information:"
	CONOUT "    Object Name: $ADActivationObjectName"
	CONOUT "    Domain Name: $ADActivationObjectDN"
	CONOUT "    CSVLK Extended PID: $ADActivationCsvlkPid"
	CONOUT "    CSVLK Activation ID: $ADActivationCsvlkSkuId"
}

function DetectAvmClient
{
	CONOUT "`nAutomatic VM Activation client information:"
	if (-Not [String]::IsNullOrEmpty($IAID)) {
		CONOUT "    Guest IAID: $IAID"
	} else {
		CONOUT "    Guest IAID: Not Available"
	}
	if (-Not [String]::IsNullOrEmpty($AutomaticVMActivationHostMachineName)) {
		CONOUT "    Host machine name: $AutomaticVMActivationHostMachineName"
	} else {
		CONOUT "    Host machine name: Not Available"
	}
	if ($AutomaticVMActivationLastActivationTime.Substring(0,4) -NE "1601") {
		$EED = [DateTime]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($AutomaticVMActivationLastActivationTime),$null,48).ToString('yyyy-MM-dd hh:mm:ss tt')
		CONOUT "    Activation time: $EED UTC"
	} else {
		CONOUT "    Activation time: Not Available"
	}
	if (-Not [String]::IsNullOrEmpty($AutomaticVMActivationHostDigitalPid2)) {
		CONOUT "    Host Digital PID2: $AutomaticVMActivationHostDigitalPid2"
	} else {
		CONOUT "    Host Digital PID2: Not Available"
	}
}

function DetectKmsHost
{
	if ($Vista -Or $NT5) {
		$KeyManagementServiceListeningPort = strGetRegistry $SLKeyPath "KeyManagementServiceListeningPort"
		$KeyManagementServiceDnsPublishing = strGetRegistry $SLKeyPath "DisableDnsPublishing"
		$KeyManagementServiceLowPriority = strGetRegistry $SLKeyPath "EnableKmsLowPriority"
		if (-Not $KeyManagementServiceDnsPublishing) {$KeyManagementServiceDnsPublishing = "TRUE"}
		if (-Not $KeyManagementServiceLowPriority) {$KeyManagementServiceLowPriority = "FALSE"}
	} else {
		$KeyManagementServiceListeningPort = $objSvc.KeyManagementServiceListeningPort
		$KeyManagementServiceDnsPublishing = $objSvc.KeyManagementServiceDnsPublishing
		$KeyManagementServiceLowPriority = $objSvc.KeyManagementServiceLowPriority
	}

	if (-Not $KeyManagementServiceListeningPort) {$KeyManagementServiceListeningPort = 1688}
	if ($KeyManagementServiceDnsPublishing -EQ "TRUE") {
		$KeyManagementServiceDnsPublishing = "Enabled"
	} else {
		$KeyManagementServiceDnsPublishing = "Disabled"
	}
	if ($KeyManagementServiceLowPriority -EQ "TRUE") {
		$KeyManagementServiceLowPriority = "Low"
	} else {
		$KeyManagementServiceLowPriority = "Normal"
	}

	CONOUT "`nKey Management Service host information:"
	CONOUT "    Current count: $KeyManagementServiceCurrentCount"
	CONOUT "    Listening on Port: $KeyManagementServiceListeningPort"
	CONOUT "    DNS publishing: $KeyManagementServiceDnsPublishing"
	CONOUT "    KMS priority: $KeyManagementServiceLowPriority"
	if (-Not [String]::IsNullOrEmpty($KeyManagementServiceTotalRequests)) {
		CONOUT "`nKey Management Service cumulative requests received from clients:"
		CONOUT "    Total: $KeyManagementServiceTotalRequests"
		CONOUT "    Failed: $KeyManagementServiceFailedRequests"
		CONOUT "    Unlicensed: $KeyManagementServiceUnlicensedRequests"
		CONOUT "    Licensed: $KeyManagementServiceLicensedRequests"
		CONOUT "    Initial grace period: $KeyManagementServiceOOBGraceRequests"
		CONOUT "    Expired or Hardware out of tolerance: $KeyManagementServiceOOTGraceRequests"
		CONOUT "    Non-genuine grace period: $KeyManagementServiceNonGenuineGraceRequests"
		if ($null -NE $KeyManagementServiceNotificationRequests) {CONOUT "    Notification: $KeyManagementServiceNotificationRequests"}
	}
}

function DetectKmsClient
{
	if ($null -NE $VLActivationTypeEnabled) {CONOUT "Configured Activation Type: $($VLActTypes[$VLActivationTypeEnabled])"}
	CONOUT "`r"
	if ($LicenseStatus -NE 1) {
		CONOUT "Please activate the product in order to update KMS client information values."
		return
	}

	if ($Vista) {
		$KeyManagementServicePort = strGetRegistry $SLKeyPath "KeyManagementServicePort"
		$DiscoveredKeyManagementServiceMachineName = strGetRegistry $NSKeyPath "DiscoveredKeyManagementServiceName"
		$DiscoveredKeyManagementServiceMachinePort = strGetRegistry $NSKeyPath "DiscoveredKeyManagementServicePort"
	}

	if ([String]::IsNullOrEmpty($KeyManagementServiceMachine)) {
		$KmsReg = $null
	} else {
		if (-Not $KeyManagementServicePort) {$KeyManagementServicePort = 1688}
		$KmsReg = "Registered KMS machine name: ${KeyManagementServiceMachine}:${KeyManagementServicePort}"
	}

	if ([String]::IsNullOrEmpty($DiscoveredKeyManagementServiceMachineName)) {
		$KmsDns = "DNS auto-discovery: KMS name not available"
		if ($Vista -And -Not $Admin) {$KmsDns = "DNS auto-discovery: Run the script as administrator to retrieve info"}
	} else {
		if (-Not $DiscoveredKeyManagementServiceMachinePort) {$DiscoveredKeyManagementServiceMachinePort = 1688}
		$KmsDns = "KMS machine name from DNS: ${DiscoveredKeyManagementServiceMachineName}:${DiscoveredKeyManagementServiceMachinePort}"
	}

	if ($null -NE $objSvc.KeyManagementServiceHostCaching) {
		if ($objSvc.KeyManagementServiceHostCaching -EQ "TRUE") {
			$KeyManagementServiceHostCaching = "Enabled"
		} else {
			$KeyManagementServiceHostCaching = "Disabled"
		}
	}

	CONOUT "Key Management Service client information:"
	CONOUT "    Client Machine ID (CMID): $($objSvc.ClientMachineID)"
	if ($null -EQ $KmsReg) {
		CONOUT "    $KmsDns"
		CONOUT "    Registered KMS machine name: KMS name not available"
	} else {
		CONOUT "    $KmsReg"
	}
	if ($null -NE $DiscoveredKeyManagementServiceMachineIpAddress) {CONOUT "    KMS machine IP address: $DiscoveredKeyManagementServiceMachineIpAddress"}
	CONOUT "    KMS machine extended PID: $KeyManagementServiceProductKeyID"
	CONOUT "    Activation interval: $VLActivationInterval minutes"
	CONOUT "    Renewal interval: $VLRenewalInterval minutes"
	if ($null -NE $KeyManagementServiceHostCaching) {CONOUT "    KMS host caching: $KeyManagementServiceHostCaching"}
	if (-Not [String]::IsNullOrEmpty($KeyManagementServiceLookupDomain)) {CONOUT "    KMS SRV record lookup domain: $KeyManagementServiceLookupDomain"}
}

function GetResult($strSLP, $strSLS, $strID)
{
	try
	{
		$objPrd = [wmisearcher]"SELECT * FROM $strSLP WHERE ID='$strID'"
		$objPrd.Options.Rewindable = $false
		$objPrd.Get() | select -Expand Properties -EA 0 | foreach { if (-Not [String]::IsNullOrEmpty($_.Value)) {set $_.Name $_.Value} }
		$objPrd.Dispose()
	}
	catch
	{
		return
	}

	$winID = ($ApplicationID -EQ $winApp)
	$winPR = ($winID -And -Not $LicenseIsAddon)
	$Vista = ($winID -And $NT6 -And -Not $NT7)
	$NT5 = ($strSLP -EQ $wslp -And $winbuild -LT 6001)
	$reapp = ("Windows", "App")[!$winID]
	$prmnt = ("machine", "product")[!$winPR]

	if ($Description | Select-String "VOLUME_KMSCLIENT") {$cKmsClient = 1; $_mTag = "Volume"}
	if ($Description | Select-String "TIMEBASED_") {$cTblClient = 1; $_mTag = "Timebased"}
	if ($Description | Select-String "VIRTUAL_MACHINE_ACTIVATION") {$cAvmClient = 1; $_mTag = "Automatic VM"}
	if ($null -EQ $cKmsClient) {
		if ($Description | Select-String "VOLUME_KMS") {$cKmsHost = 1}
	}

	$_gpr = [Math]::Round($GracePeriodRemaining/1440)
	if ($_gpr -GT 0) {
		$_xpr = [DateTime]::Now.addMinutes($GracePeriodRemaining).ToString('yyyy-MM-dd hh:mm:ss tt')
	}

	if ($null -EQ $LicenseStatusReason) {$LicenseStatusReason = -1}
	$LicenseReason = '0x{0:X}' -f $LicenseStatusReason
	$LicenseMsg = "Time remaining: $GracePeriodRemaining minute(s) ($_gpr day(s))"
	if ($LicenseStatus -EQ 0) {
		$LicenseInf = "Unlicensed"
		$LicenseMsg = $null
	}
	if ($LicenseStatus -EQ 1) {
		$LicenseInf = "Licensed"
		$LicenseMsg = $null
		if ($GracePeriodRemaining -EQ 0) {
			$ExpireMsg = "The $prmnt is permanently activated."
		} else {
			$LicenseMsg = "$_mTag activation expiration: $GracePeriodRemaining minute(s) ($_gpr day(s))"
			if ($null -NE $_xpr) {$ExpireMsg = "$_mTag activation will expire $_xpr"}
		}
	}
	if ($LicenseStatus -EQ 2) {
		$LicenseInf = "Initial grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Initial grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 3) {
		$LicenseInf = "Additional grace period (KMS license expired or hardware out of tolerance)"
		if ($null -NE $_xpr) {$ExpireMsg = "Additional grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 4) {
		$LicenseInf = "Non-genuine grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Non-genuine grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 5 -And -Not $NT5) {
		$LicenseInf = "Notification"
		$LicenseMsg = "Notification Reason: $LicenseReason"
		if ($LicenseReason -EQ "0xC004F00F") {if ($null -NE $cKmsClient) {$LicenseMsg = $LicenseMsg + " (KMS license expired)."} else {$LicenseMsg = $LicenseMsg + " (hardware out of tolerance)."}}
		if ($LicenseReason -EQ "0xC004F200") {$LicenseMsg = $LicenseMsg + " (non-genuine)."}
		if ($LicenseReason -EQ "0xC004F009" -Or $LicenseReason -EQ "0xC004F064") {$LicenseMsg = $LicenseMsg + " (grace time expired)."}
	}
	if ($LicenseStatus -GT 5 -Or ($LicenseStatus -GT 4 -And $NT5)) {
		$LicenseInf = "Unknown"
		$LicenseMsg = $null
	}
	if ($LicenseStatus -EQ 6 -And -Not $Vista -And -Not $NT5) {
		$LicenseInf = "Extended grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Extended grace period ends $_xpr"}
	}

	if ($winPR -And $PartialProductKey -And -Not $NT9) {
		$dp4 = strGetRegistry "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" "DigitalProductId4"
		if ($null -NE $dp4) {
			$ProductKeyChannel = ([System.Text.Encoding]::Unicode.GetString($dp4, 1016, 128)).Trim([char]$null)
		}
	}

	if ($winPR -And $Dlv -And $null -EQ $RemainingAppReArmCount) {
		try
		{
			$tmp = [wmisearcher]"SELECT RemainingWindowsReArmCount FROM $strSLS"
			$tmp.Options.Rewindable = $false
			$tmp.Get() | select -Expand Properties -EA 0 | foreach {set $_.Name $_.Value}
			$tmp.Dispose()
		}
		catch
		{
		}
	}

	$add_on = $Name.IndexOf("add-on for", 5)

	& $isAll
	if ($add_on -EQ -1) {CONOUT "Name: $Name"} else {CONOUT "Name: $($Name.Substring(0, $add_on + 7))"}
	CONOUT "Description: $Description"
	CONOUT "Activation ID: $ID"
	if ($null -NE $ProductKeyID) {CONOUT "Extended PID: $ProductKeyID"}
	if ($null -NE $ProductKeyID2 -And $Dlv) {CONOUT "Product ID: $ProductKeyID2"}
	if ($null -NE $OfflineInstallationId -And $IID) {CONOUT "Installation ID: $OfflineInstallationId"}
	if ($null -NE $ProductKeyChannel) {CONOUT "Product Key Channel: $ProductKeyChannel"}
	if ($null -NE $PartialProductKey) {CONOUT "Partial Product Key: $PartialProductKey"}
	CONOUT "License Status: $LicenseInf"
	if ($null -NE $LicenseMsg) {CONOUT "$LicenseMsg"}
	if ($LicenseStatus -NE 0 -And $EvaluationEndDate.Substring(0,4) -NE "1601") {
		$EED = [DateTime]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($EvaluationEndDate),$null,48).ToString('yyyy-MM-dd hh:mm:ss tt')
		CONOUT "Evaluation End Date: $EED UTC"
	}
	if ($Dlv) {
		if ($null -NE $RemainingWindowsReArmCount) {
			CONOUT "Remaining Windows rearm count: $RemainingWindowsReArmCount"
		}
		if ($null -NE $RemainingSkuReArmCount -And $RemainingSkuReArmCount -NE 4294967295) {
			CONOUT "Remaining $reapp rearm count: $RemainingAppReArmCount"
			CONOUT "Remaining SKU rearm count: $RemainingSkuReArmCount"
		}
		if ($null -NE $TrustedTime -And $LicenseStatus -NE 0) {
			$TTD = [DateTime]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($TrustedTime),$null,32).ToString('yyyy-MM-dd hh:mm:ss tt')
			CONOUT "Trusted time: $TTD"
		}
	}
	if ($LicenseStatus -EQ 0) {
		return
	}

	if ($strSLP -EQ $wslp -And $null -NE $PartialProductKey -And $null -NE $ADActivationObjectName -And $VLActivationType -EQ 1) {
		DetectAdbaClient
	}

	if ($winID -And $null -NE $cAvmClient -And $null -NE $PartialProductKey) {
		DetectAvmClient
	}

	$chkSub = ($winPR -And $cSub)

	$chkSLS = ($null -NE $PartialProductKey) -And ($null -NE $cKmsClient -Or $null -NE $cKmsHost -Or $chkSub)

	if (!$chkSLS) {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
		return
	}

	try
	{
		$objSvc = New-Object PSObject
		$wmiSvc = [wmisearcher]"SELECT * FROM $strSLS"
		$wmiSvc.Options.Rewindable = $false
		$wmiSvc.Get() | select -Expand Properties -EA 0 | foreach {
			if (-Not [String]::IsNullOrEmpty($_.Value))
			{
				$objSvc | Add-Member 8 $_.Name $_.Value
				if ($null -EQ $IsKeyManagementServiceMachine) {set $_.Name $_.Value}
			}
		}
		$wmiSvc.Dispose()
	}
	catch
	{
		return
	}

	if ($strSLS -EQ $wsls -And $NT9) {
		if ([String]::IsNullOrEmpty($DiscoveredKeyManagementServiceMachineIpAddress)) {
			$DiscoveredKeyManagementServiceMachineIpAddress = "not available"
		}
	}

	if ($null -NE $cKmsHost -And $IsKeyManagementServiceMachine -GT 0) {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
		DetectKmsHost
	}

	if ($null -NE $cKmsClient) {
		DetectKmsClient
	}

	if ($null -EQ $cKmsHost) {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
	}

	if ($chkSub) {
		DetectSubscription
	}

}
#endregion

#region vNextDiag
if ($PSVersionTable.PSVersion.Major -Lt 3)
{
	function ConvertFrom-Json
	{
		[CmdletBinding()]
		Param(
			[Parameter(ValueFromPipeline=$true)][Object]$item
		)
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
		$psjs = New-Object System.Web.Script.Serialization.JavaScriptSerializer
		Return ,$psjs.DeserializeObject($item)
	}
	function ConvertTo-Json
	{
		[CmdletBinding()]
		Param(
			[Parameter(ValueFromPipeline=$true)][Object]$item
		)
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
		$psjs = New-Object System.Web.Script.Serialization.JavaScriptSerializer
		Return $psjs.Serialize($item)
	}
}

function PrintModePerPridFromRegistry
{
	$vNextRegkey = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext"
	$vNextPrids = Get-Item -Path $vNextRegkey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'property' -ErrorAction SilentlyContinue | Where-Object -FilterScript {$_.ToLower() -like "*retail" -or $_.ToLower() -like "*volume"}
	If ($null -Eq $vNextPrids)
	{
		CONOUT "`nNo registry keys found."
		Return
	}
	CONOUT "`r"
	$vNextPrids | ForEach `
	{
		$mode = (Get-ItemProperty -Path $vNextRegkey -Name $_).$_
		Switch ($mode)
		{
			2 { $mode = "vNext"; Break }
			3 { $mode = "Device"; Break }
			Default { $mode = "Legacy"; Break }
		}
		CONOUT "$_ = $mode"
	}
}

function PrintSharedComputerLicensing
{
	$scaRegKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	$scaValue = Get-ItemProperty -Path $scaRegKey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	$scaRegKey2 = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing"
	$scaValue2 = Get-ItemProperty -Path $scaRegKey2 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	$scaPolicyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Licensing"
	$scaPolicyValue = Get-ItemProperty -Path $scaPolicyKey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	If ($null -Eq $scaValue -And $null -Eq $scaValue2 -And $null -Eq $scaPolicyValue)
	{
		CONOUT "`nNo registry keys found."
		Return
	}
	$scaModeValue = $scaValue -Or $scaValue2 -Or $scaPolicyValue
	If ($scaModeValue -Eq 0)
	{
		$scaMode = "Disabled"
	}
	If ($scaModeValue -Eq 1)
	{
		$scaMode = "Enabled"
	}
	CONOUT "`nStatus: $scaMode"
	CONOUT "`r"
	$tokenFiles = $null
	$tokenPath = "${env:LOCALAPPDATA}\Microsoft\Office\16.0\Licensing"
	If (Test-Path $tokenPath)
	{
		$tokenFiles = Get-ChildItem -Path $tokenPath -Filter "*authString*" -Recurse | Where-Object { !$_.PSIsContainer }
	}
	If ($null -Eq $tokenFiles)
	{
		CONOUT "No tokens found."
		Return
	}
	If ($tokenFiles.Length -Eq 0)
	{
		CONOUT "No tokens found."
		Return
	}
	$tokenFiles | ForEach `
	{
		$tokenParts = (Get-Content -Encoding Unicode -Path $_.FullName).Split('_')
		$output = New-Object PSObject
		$output | Add-Member 8 'ACID' $tokenParts[0];
		$output | Add-Member 8 'User' $tokenParts[3];
		$output | Add-Member 8 'NotBefore' $tokenParts[4];
		$output | Add-Member 8 'NotAfter' $tokenParts[5];
		Write-Output $output
	}
}

function PrintLicensesInformation
{
	Param(
		[ValidateSet("NUL", "Device")]
		[String]$mode
	)
	If ($mode -Eq "NUL")
	{
		$licensePath = "${env:LOCALAPPDATA}\Microsoft\Office\Licenses"
	}
	ElseIf ($mode -Eq "Device")
	{
		$licensePath = "${env:PROGRAMDATA}\Microsoft\Office\Licenses"
	}
	$licenseFiles = $null
	If (Test-Path $licensePath)
	{
		$licenseFiles = Get-ChildItem -Path $licensePath -Recurse | Where-Object { !$_.PSIsContainer }
	}
	If ($null -Eq $licenseFiles -Or $licenseFiles.Length -Eq 0)
	{
		CONOUT "`nNo licenses found."
		Return
	}
	$licenseFiles | ForEach `
	{
		$license = (Get-Content -Encoding Unicode $_.FullName | ConvertFrom-Json).License
		$decodedLicense = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($license)) | ConvertFrom-Json
		$licenseType = $decodedLicense.LicenseType
		If ($null -Ne $decodedLicense.ExpiresOn)
		{
			$expiry = [System.DateTime]::Parse($decodedLicense.ExpiresOn, $null, 'AdjustToUniversal')
		}
		Else
		{
			$expiry = New-Object System.DateTime
		}
		$licenseState = "Grace"
		If ((Get-Date) -Gt (Get-Date $decodedLicense.Metadata.NotAfter))
		{
			$licenseState = "RFM"
		}
		ElseIf ((Get-Date) -Lt (Get-Date $expiry))
		{
			$licenseState = "Licensed"
		}
		$output = New-Object PSObject
		$output | Add-Member 8 'File' $_.PSChildName;
		$output | Add-Member 8 'Version' $_.Directory.Name;
		$output | Add-Member 8 'Type' "User|${licenseType}";
		$output | Add-Member 8 'Product' $decodedLicense.ProductReleaseId;
		$output | Add-Member 8 'Acid' $decodedLicense.Acid;
		If ($mode -Eq "Device") { $output | Add-Member 8 'DeviceId' $decodedLicense.Metadata.DeviceId; }
		$output | Add-Member 8 'LicenseState' $licenseState;
		$output | Add-Member 8 'EntitlementStatus' $decodedLicense.Status;
		$output | Add-Member 8 'EntitlementExpiration' ("N/A", $decodedLicense.ExpiresOn)[!($null -eq $decodedLicense.ExpiresOn)];
		$output | Add-Member 8 'ReasonCode' ("N/A", $decodedLicense.ReasonCode)[!($null -eq $decodedLicense.ReasonCode)];
		$output | Add-Member 8 'NotBefore' $decodedLicense.Metadata.NotBefore;
		$output | Add-Member 8 'NotAfter' $decodedLicense.Metadata.NotAfter;
		$output | Add-Member 8 'NextRenewal' $decodedLicense.Metadata.RenewAfter;
		$output | Add-Member 8 'TenantId' ("N/A", $decodedLicense.Metadata.TenantId)[!($null -eq $decodedLicense.Metadata.TenantId)];
		#$output.PSObject.Properties | foreach { $ht = @{} } { $ht[$_.Name] = $_.Value } { $output = $ht | ConvertTo-Json }
		Write-Output $output
	}
}

function vNextDiagRun
{
	$fNUL = ([IO.Directory]::Exists("${env:LOCALAPPDATA}\Microsoft\Office\Licenses")) -and ([IO.Directory]::GetFiles("${env:LOCALAPPDATA}\Microsoft\Office\Licenses", "*", 1).Length -GE 0)
	$fDev = ([IO.Directory]::Exists("${env:PROGRAMDATA}\Microsoft\Office\Licenses")) -and ([IO.Directory]::GetFiles("${env:PROGRAMDATA}\Microsoft\Office\Licenses", "*", 1).Length -GE 0)
	$rPID = $null -NE (GP "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext" -EA 0 | select -Expand 'property' -EA 0 | where -Filter {$_.ToLower() -like "*retail" -or $_.ToLower() -like "*volume"})
	$rSCA = $null -NE (GP "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -EA 0 | select -Expand "SharedComputerLicensing" -EA 0)
	$rSCL = $null -NE (GP "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing" -EA 0 | select -Expand "SharedComputerLicensing" -EA 0)

	if (($fNUL -Or $fDev -Or $rPID -Or $rSCA -Or $rSCL) -EQ $false) {
		Return
	}

	& $isAll
	CONOUT "$line2"
	CONOUT "===                  Office vNext Status                 ==="
	CONOUT "$line2"
	CONOUT "`n========== Mode per ProductReleaseId =========="
	PrintModePerPridFromRegistry
	CONOUT "`n========== Shared Computer Licensing =========="
	PrintSharedComputerLicensing
	CONOUT "`n========== vNext licenses ==========="
	PrintLicensesInformation -Mode "NUL"
	CONOUT "`n========== Device licenses =========="
	PrintLicensesInformation -Mode "Device"
	CONOUT "$line3"
	CONOUT "`r"
}
#endregion

#region clic

<#
;;; Source: https://github.com/asdcorp/clic
;;; Powershell port: abbodi1406

Copyright 2023 asdcorp

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

function BoolToWStr($bVal) {
	("TRUE", "FALSE")[!$bVal]
}

function InitializePInvoke {
	$Marshal = [System.Runtime.InteropServices.Marshal]
	$Module = [AppDomain]::CurrentDomain.DefineDynamicAssembly((Get-Random), 'Run').DefineDynamicModule((Get-Random))

	$Class = $Module.DefineType('NativeMethods', 'Public, Abstract, Sealed, BeforeFieldInit', [Object], 0)
	$Class.DefinePInvokeMethod('SLIsWindowsGenuineLocal', 'slc.dll', 'Public, Static', 'Standard', [Int32], @([UInt32].MakeByRefType()), 'Winapi', 'Unicode').SetImplementationFlags('PreserveSig')
	$Class.DefinePInvokeMethod('SLGetWindowsInformationDWORD', 'slc.dll', 22, 1, [Int32], @([String], [UInt32].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGetWindowsInformation', 'slc.dll', 22, 1, [Int32], @([String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)

	if ($DllSubscription) {
		$Class.DefinePInvokeMethod('ClipGetSubscriptionStatus', 'Clipc.dll', 22, 1, [Int32], @([IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
		$Struct = $Class.DefineNestedType('SubStatus', 'NestedPublic, SequentialLayout, Sealed, BeforeFieldInit', [ValueType], 0)
		[void]$Struct.DefineField('dwEnabled', [UInt32], 'Public')
		[void]$Struct.DefineField('dwSku', [UInt32], 6)
		[void]$Struct.DefineField('dwState', [UInt32], 6)
		$SubStatus = $Struct.CreateType()
	}

	$Win32 = $Class.CreateType()
}

function InitializeDigitalLicenseCheck {
	$CAB = [System.Reflection.Emit.CustomAttributeBuilder]

	$ICom = $Module.DefineType('EUM.IEUM', 'Public, Interface, Abstract, Import')
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.ComImportAttribute].GetConstructor(@()), @()))
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.GuidAttribute].GetConstructor(@([String])), @('F2DCB80D-0670-44BC-9002-CD18688730AF')))
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.InterfaceTypeAttribute].GetConstructor(@([Int16])), @([Int16]1)))

	1..4 | % { [void]$ICom.DefineMethod('VF'+$_, 'Public, Virtual, HideBySig, NewSlot, Abstract', 'Standard, HasThis', [Void], @()) }
	[void]$ICom.DefineMethod('AcquireModernLicenseForWindows', 1478, 33, [Int32], @([Int32], [Int32].MakeByRefType()))

	$IEUM = $ICom.CreateType()
}

function PrintStateData {
	$pwszStateData = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-Action-StateData",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pwszStateData
	)) {
		return $FALSE
	}

	[string[]]$pwszStateString = $Marshal::PtrToStringUni($pwszStateData) -replace ";", "`n    "
	CONOUT ("    $pwszStateString")

	$Marshal::FreeHGlobal($pwszStateData)
	return $TRUE
}

function PrintLastActivationHResult {
	$pdwLastHResult = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-LastWindowsActivationHResult",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pdwLastHResult
	)) {
		return $FALSE
	}

	CONOUT ("    LastActivationHResult=0x{0:x8}" -f $Marshal::ReadInt32($pdwLastHResult))

	$Marshal::FreeHGlobal($pdwLastHResult)
	return $TRUE
}

function PrintLastActivationTime {
	$pdwLastTime = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-LastWindowsActivationTime",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pdwLastTime
	)) {
		return $FALSE
	}

	$actTime = $Marshal::ReadInt64($pdwLastTime)
	if ($actTime -ne 0) {
		CONOUT ("    LastActivationTime={0}" -f [DateTime]::FromFileTimeUtc($actTime).ToString("yyyy/MM/dd:HH:mm:ss"))
	}

	$Marshal::FreeHGlobal($pdwLastTime)
	return $TRUE
}

function PrintIsWindowsGenuine {
	$dwGenuine = 0
	$ppwszGenuineStates = @(
		"SL_GEN_STATE_IS_GENUINE",
		"SL_GEN_STATE_INVALID_LICENSE",
		"SL_GEN_STATE_TAMPERED",
		"SL_GEN_STATE_OFFLINE",
		"SL_GEN_STATE_LAST"
	)

	if ($Win32::SLIsWindowsGenuineLocal([ref]$dwGenuine)) {
		return $FALSE
	}

	if ($dwGenuine -lt 5) {
		CONOUT ("    IsWindowsGenuine={0}" -f $ppwszGenuineStates[$dwGenuine])
	} else {
		CONOUT ("    IsWindowsGenuine={0}" -f $dwGenuine)
	}

	return $TRUE
}

function PrintDigitalLicenseStatus {
	try {
		. InitializeDigitalLicenseCheck
		$ComObj = New-Object -Com EditionUpgradeManagerObj.EditionUpgradeManager
	} catch {
		return $FALSE
	}

	$parameters = 1, $null

	if ([EUM.IEUM].GetMethod("AcquireModernLicenseForWindows").Invoke($ComObj, $parameters)) {
		return $FALSE
	}

	$dwReturnCode = $parameters[1]
	[bool]$bDigitalLicense = $FALSE

	$bDigitalLicense = (($dwReturnCode -ge 0) -and ($dwReturnCode -ne 1))
	CONOUT ("    IsDigitalLicense={0}" -f (BoolToWStr $bDigitalLicense))

	return $TRUE
}

function PrintSubscriptionStatus {
	$dwSupported = 0

	if ($winbuild -ge 15063) {
		$pwszPolicy = "ConsumeAddonPolicySet"
	} else {
		$pwszPolicy = "Allow-WindowsSubscription"
	}

	if ($Win32::SLGetWindowsInformationDWORD($pwszPolicy, [ref]$dwSupported)) {
		return $FALSE
	}

	CONOUT ("    SubscriptionSupportedEdition={0}" -f (BoolToWStr $dwSupported))

	$pStatus = $Marshal::AllocHGlobal($Marshal::SizeOf([Type]$SubStatus))
	if ($Win32::ClipGetSubscriptionStatus([ref]$pStatus)) {
		return $FALSE
	}

	$sStatus = [Activator]::CreateInstance($SubStatus)
	$sStatus = $Marshal::PtrToStructure($pStatus, [Type]$SubStatus)
	$Marshal::FreeHGlobal($pStatus)

	CONOUT ("    SubscriptionEnabled={0}" -f (BoolToWStr $sStatus.dwEnabled))

	if ($sStatus.dwEnabled -eq 0) {
		return $TRUE
	}

	CONOUT ("    SubscriptionSku={0}" -f $sStatus.dwSku)
	CONOUT ("    SubscriptionState={0}" -f $sStatus.dwState)

	return $TRUE
}

function ClicRun
{
	& $isAll
	CONOUT "Client Licensing Check information:"

	$null = PrintStateData
	$null = PrintLastActivationHResult
	$null = PrintLastActivationTime
	$null = PrintIsWindowsGenuine

	if ($DllDigital) {
		$null = PrintDigitalLicenseStatus
	}

	if ($DllSubscription) {
		$null = PrintSubscriptionStatus
	}

	CONOUT "$line3"
	& $noAll
}
#endregion

$Host.UI.RawUI.WindowTitle = "Check Activation Status"
UnQuickEdit
if ($All.IsPresent) {
	$B=$Host.UI.RawUI.BufferSize;$B.Height=3000;$Host.UI.RawUI.BufferSize=$B;
	if (!$Pass.IsPresent) {clear;}
}

$wslp = "SoftwareLicensingProduct"
$wsls = "SoftwareLicensingService"
$oslp = "OfficeSoftwareProtectionProduct"
$osls = "OfficeSoftwareProtectionService"
$winApp = "55c92734-d682-4d71-983e-d6ec3f16059f"
$o14App = "59a52881-a989-479d-af46-f275c6370663"
$o15App = "0ff1ce15-a989-479d-af46-f275c6370663"
$cSub = ($winbuild -GE 19041) -And (Select-String -Path "$SysPath\wbem\sppwmi.mof" -Encoding unicode -Pattern "SubscriptionType")
$DllDigital = ($winbuild -GE 14393) -And (Test-Path "$SysPath\EditionUpgradeManagerObj.dll")
$DllSubscription = ($winbuild -GE 14393) -And (Test-Path "$SysPath\Clipc.dll")
$VLActTypes = @("All", "AD", "KMS", "Token")
$SLKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"
$NSKeyPath = "HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"

'cW1nd0ws', 'c0ff1ce15', 'c0ff1ce14', 'ospp14', 'ospp15' | foreach {set $_ $false}

$offsvc = "osppsvc"
if ($NT7 -Or -Not $NT6) {$winsvc = "sppsvc"} else {$winsvc = "slsvc"}

try {gsv $winsvc -EA 1 | Out-Null; $WsppHook = 1} catch {$WsppHook = 0}
try {gsv $offsvc -EA 1 | Out-Null; $OsppHook = 1} catch {$OsppHook = 0}

if ($WsppHook -NE 0) {
	try {sasv $winsvc -EA 1} catch {}
	$cW1nd0ws  = DetectID $wslp $winApp
	$c0ff1ce15 = DetectID $wslp $o15App
	$c0ff1ce14 = DetectID $wslp $o14App
}

if ($OsppHook -NE 0) {
	try {sasv $offsvc -EA 1} catch {}
	$ospp15 = DetectID $oslp $o15App
	$ospp14 = DetectID $oslp $o14App
}

if ($cW1nd0ws)
{
	echoWindows
	GetID $wslp $winApp | foreach -EA 1 {
	GetResult $wslp $wsls $_
	CONOUT "$line3"
	& $noAll
	}
}
elseif ($NT6)
{
	echoWindows
	CONOUT "`nError: product key not found."
}

if ($winbuild -GE 9200) {
	. InitializePInvoke
	ClicRun
}

if ($c0ff1ce15 -Or $ospp15) {
	CheckOhook
}

$doMSG = 1

if ($c0ff1ce15)
{
	echoOffice
	GetID $wslp $o15App | foreach -EA 1 {
	GetResult $wslp $wsls $_
	CONOUT "$line3"
	& $noAll
	}
}

if ($c0ff1ce14)
{
	echoOffice
	GetID $wslp $o14App | foreach -EA 1 {
	GetResult $wslp $wsls $_
	CONOUT "$line3"
	& $noAll
	}
}

if ($ospp15)
{
	echoOffice
	GetID $oslp $o15App | foreach -EA 1 {
	GetResult $oslp $osls $_
	CONOUT "$line3"
	& $noAll
	}
}

if ($ospp14)
{
	echoOffice
	GetID $oslp $o14App | foreach -EA 1 {
	GetResult $oslp $osls $_
	CONOUT "$line3"
	& $noAll
	}
}

if ($NT7) {
	vNextDiagRun
}

ExitScript 0
:sppmgr:
