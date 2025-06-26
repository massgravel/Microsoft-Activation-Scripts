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

set "ps=%SysPath%\WindowsPowerShell\v1.0\powershell.exe"
set "_psc=%ps% -nop -c"
set "_err===== ERROR ===="
set _pwsh=1
if not exist %ps% set _pwsh=0
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
$NT8 = $winbuild -GE 9200
$NT9 = $winbuild -GE 9600

$Admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$line2 = "============================================================"
$line3 = "____________________________________________________________"

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

#region SSSS
function BoolToWStr($bVal) {
	("TRUE", "FALSE")[!$bVal]
}

function InitializePInvoke($LaDll, $bOffice) {
	$LaName = [IO.Path]::GetFileNameWithoutExtension($LaDll)
	$SLApp = $NT7 -Or $bOffice -Or ($LaName -EQ 'sppc' -And [Diagnostics.FileVersionInfo]::GetVersionInfo("$SysPath\sppc.dll").FilePrivatePart -GE 16501)
	$Win32 = $null

	$Marshal = [System.Runtime.InteropServices.Marshal]
	$Module = [AppDomain]::CurrentDomain.DefineDynamicAssembly(($LaName+"_Assembly"), 'Run').DefineDynamicModule(($LaName+"_Module"), $False)
	$Class = $Module.DefineType(($LaName+"_Methods"), 'Public, Abstract, Sealed, BeforeFieldInit', [Object], 0)

	$Class.DefinePInvokeMethod('SLClose', $LaDll, 22, 1, [Int32], @([IntPtr]), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLOpen', $LaDll, 22, 1, [Int32], @([IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGenerateOfflineInstallationId', $LaDll, 22, 1, [Int32], @([IntPtr], [Guid].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGetSLIDList', $LaDll, 22, 1, [Int32], @([IntPtr], [UInt32], [Guid].MakeByRefType(), [UInt32], [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGetLicensingStatusInformation', $LaDll, 22, 1, [Int32], @([IntPtr], [Guid].MakeByRefType(), [Guid].MakeByRefType(), [IntPtr], [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGetPKeyInformation', $LaDll, 22, 1, [Int32], @([IntPtr], [Guid].MakeByRefType(), [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGetProductSkuInformation', $LaDll, 22, 1, [Int32], @([IntPtr], [Guid].MakeByRefType(), [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGetServiceInformation', $LaDll, 22, 1, [Int32], @([IntPtr], [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	if ($SLApp) {
		$Class.DefinePInvokeMethod('SLGetApplicationInformation', $LaDll, 22, 1, [Int32], @([IntPtr], [Guid].MakeByRefType(), [String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	}
	if ($bOffice) {
		$Win32 = $Class.CreateType()
		return
	}
	if ($NT6) {
		$Class.DefinePInvokeMethod('SLGetWindowsInformation', 'slc.dll', 22, 1, [Int32], @([String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
		$Class.DefinePInvokeMethod('SLGetWindowsInformationDWORD', 'slc.dll', 22, 1, [Int32], @([String], [UInt32].MakeByRefType()), 1, 3).SetImplementationFlags(128)
		$Class.DefinePInvokeMethod('SLIsGenuineLocal', 'slwga.dll', 22, 1, [Int32], @([Guid].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr]), 1, 3).SetImplementationFlags(128)
	}
	if ($NT7) {
		$Class.DefinePInvokeMethod('SLIsWindowsGenuineLocal', 'slc.dll', 'Public, Static', 'Standard', [Int32], @([UInt32].MakeByRefType()), 'Winapi', 'Unicode').SetImplementationFlags('PreserveSig')
	}

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

function SlGetInfoIID($SkuId)
{
	$bData = 0

	if ($Win32::SLGenerateOfflineInstallationId(
		$hSLC,
		[ref][Guid]$SkuId,
		[ref]$bData
	))
	{
		return $null
	}
	else
	{
		return $Marshal::PtrToStringUni($bData)
	}
}

function SlReturnData($hrRet, $tData, $cData, $bData) {
	if ($hrRet -NE 0 -Or $cData -EQ 0)
	{
		return $null
	}
	if ($tData -EQ 1)
	{
		return $Marshal::PtrToStringUni($bData)
	}
	elseif ($tData -EQ 4)
	{
		return $Marshal::ReadInt32($bData)
	}
	elseif ($tData -EQ 3 -And $cData -EQ 8)
	{
		return $Marshal::ReadInt64($bData)
	}
	else
	{
		return $null
	}
}

function SlGetInfoPKey($PkeyId, $Value)
{
	$tData = 0
	$cData = 0
	$bData = 0

	$hrRet = $Win32::SLGetPKeyInformation(
		$hSLC,
		[ref][Guid]$PkeyId,
		$Value,
		[ref]$tData,
		[ref]$cData,
		[ref]$bData
	)

	return SlReturnData $hrRet $tData $cData $bData
}

function SlGetInfoSku($SkuId, $Value)
{
	$tData = 0
	$cData = 0
	$bData = 0

	$hrRet = $Win32::SLGetProductSkuInformation(
		$hSLC,
		[ref][Guid]$SkuId,
		$Value,
		[ref]$tData,
		[ref]$cData,
		[ref]$bData
	)

	return SlReturnData $hrRet $tData $cData $bData
}

function SlGetInfoApp($AppId, $Value)
{
	$tData = 0
	$cData = 0
	$bData = 0

	$hrRet = $Win32::SLGetApplicationInformation(
		$hSLC,
		[ref][Guid]$AppId,
		$Value,
		[ref]$tData,
		[ref]$cData,
		[ref]$bData
	)

	return SlReturnData $hrRet $tData $cData $bData
}

function SlGetInfoService($Value)
{
	$tData = 0
	$cData = 0
	$bData = 0

	$hrRet = $Win32::SLGetServiceInformation(
		$hSLC,
		$Value,
		[ref]$tData,
		[ref]$cData,
		[ref]$bData
	)

	return SlReturnData $hrRet $tData $cData $bData
}

function SlGetInfoSvcApp($strApp, $Value)
{
	if ($SLApp)
	{
		return SlGetInfoApp $strApp $Value
	}
	else
	{
		return SlGetInfoService $Value
	}
}

function SlGetInfoLicensing($AppId, $SkuId)
{
	$dwStatus = 0
	$dwGrace = 0
	$hrReason = 0
	$qwValidity = 0

	$cStatus = 0
	$pStatus = 0

	$hrRet = $Win32::SLGetLicensingStatusInformation(
		$hSLC,
		[ref][Guid]$AppId,
		[ref][Guid]$SkuId,
		0,
		[ref]$cStatus,
		[ref]$pStatus
	)

	if ($hrRet -NE 0 -Or $cStatus -EQ 0)
	{
		return
	}

	[IntPtr]$ppStatus = [Int64]$pStatus + [Int64]40 * ($cStatus - 1)
	$dwStatus = $Marshal::ReadInt32($ppStatus, 16)
	$dwGrace = $Marshal::ReadInt32($ppStatus, 20)
	$hrReason = $Marshal::ReadInt32($ppStatus, 28)
	$qwValidity = $Marshal::ReadInt64($ppStatus, 32)

	if ($dwStatus -EQ 3)
	{
		$dwStatus = 5
	}
	if ($dwStatus -EQ 2)
	{
		if ($hrReason -EQ 0x4004F00D)
		{
			$dwStatus = 3
		}
		elseif ($hrReason -EQ 0x4004F065)
		{
			$dwStatus = 4
		}
		elseif ($hrReason -EQ 0x4004FC06)
		{
			$dwStatus = 6
		}
	}

	return
}

function SlGetInfoSLID($AppId)
{
	$cReturnIds = 0
	$pReturnIds = 0

	$hrRet = $Win32::SLGetSLIDList(
		$hSLC,
		0,
		[ref][Guid]$AppId,
		1,
		[ref]$cReturnIds,
		[ref]$pReturnIds
	)

	if ($hrRet -NE 0 -Or $cReturnIds -EQ 0)
	{
		return
	}

	$a1List = @()
	$a2List = @()
	$a3List = @()
	$a4List = @()

	foreach ($i in 0..($cReturnIds - 1))
	{
		$bytes = New-Object byte[] 16
		$Marshal::Copy([Int64]$pReturnIds + [Int64]16 * $i, $bytes, 0, 16)
		$actid = ([Guid]$bytes).Guid
		$gPPK = SlGetInfoSku $actid "pkeyId"
		$gAdd = SlGetInfoSku $actid "DependsOn"
		if ($All.IsPresent) {
			if ($null -EQ $gPPK -And $null -NE $gAdd) { $a1List += @{id = $actid; pk = $null; ex = $true} }
			if ($null -EQ $gPPK -And $null -EQ $gAdd) { $a2List += @{id = $actid; pk = $null; ex = $false} }
		}
		if ($null -NE $gPPK -And $null -NE $gAdd) { $a3List += @{id = $actid; pk = $gPPK; ex = $true} }
		if ($null -NE $gPPK -And $null -EQ $gAdd) { $a4List += @{id = $actid; pk = $gPPK; ex = $false} }
	}

	return ($a1List + $a2List + $a3List + $a4List)
}

function DetectSubscription {
	try
	{
		$objSvc = New-Object PSObject
		$wmiSvc = [wmisearcher]"SELECT SubscriptionType, SubscriptionStatus, SubscriptionEdition, SubscriptionExpiry FROM SoftwareLicensingService"
		$wmiSvc.Options.Rewindable = $false
		$wmiSvc.Get() | select -Expand Properties -EA 0 | foreach { $objSvc | Add-Member 8 $_.Name $_.Value }
		$wmiSvc.Dispose()
	}
	catch
	{
		return
	}

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
	CONOUT "    Type   : $SubMsgType"
	CONOUT "    Status : $SubMsgStatus"
	CONOUT "    Edition: $SubMsgEdition"
	CONOUT "    Expiry : $SubMsgExpiry"
}

function DetectAdbaClient
{
	$propADBA | foreach { set $_ (SlGetInfoSku $licID $_) }
	CONOUT "`nAD Activation client information:"
	CONOUT "    Object Name: $ADActivationObjectName"
	CONOUT "    Domain Name: $ADActivationObjectDN"
	CONOUT "    CSVLK Extended PID: $ADActivationCsvlkPID"
	CONOUT "    CSVLK Activation ID: $ADActivationCsvlkSkuID"
}

function DetectAvmClient
{
	$propAVMA | foreach { set $_ (SlGetInfoSku $licID $_) }
	CONOUT "`nAutomatic VM Activation client information:"
	if (-Not [String]::IsNullOrEmpty($InheritedActivationId)) {
		CONOUT "    Guest IAID: $InheritedActivationId"
	} else {
		CONOUT "    Guest IAID: Not Available"
	}
	if (-Not [String]::IsNullOrEmpty($InheritedActivationHostMachineName)) {
		CONOUT "    Host machine name: $InheritedActivationHostMachineName"
	} else {
		CONOUT "    Host machine name: Not Available"
	}
	if (-Not [String]::IsNullOrEmpty($InheritedActivationHostDigitalPid2)) {
		CONOUT "    Host Digital PID2: $InheritedActivationHostDigitalPid2"
	} else {
		CONOUT "    Host Digital PID2: Not Available"
	}
	if ($InheritedActivationActivationTime) {
		$IAAT = [DateTime]::FromFileTime($InheritedActivationActivationTime).ToString('yyyy-MM-dd hh:mm:ss tt')
		CONOUT "    Activation time: $IAAT"
	} else {
		CONOUT "    Activation time: Not Available"
	}
}

function DetectKmsHost
{
	$IsKeyManagementService = SlGetInfoSvcApp $strApp 'IsKeyManagementService'
	if (-Not $IsKeyManagementService) {
		return
	}

	if ($Vista -Or $NT5) {
		$regk = $SLKeyPath
	} elseif ($strSLP -EQ $oslp) {
		$regk = $OPKeyPath
	} else {
		$regk = $SPKeyPath
	}
	$KMSListening = strGetRegistry $regk "KeyManagementServiceListeningPort"
	$KMSPublishing = strGetRegistry $regk "DisableDnsPublishing"
	$KMSPriority = strGetRegistry $regk "EnableKmsLowPriority"

	if (-Not $KMSListening) {$KMSListening = 1688}
	if (-Not $KMSPublishing) {$KMSPublishing = "TRUE"} else {$KMSPublishing = BoolToWStr (!$KMSPublishing)}
	if (-Not $KMSPriority) {$KMSPriority = "FALSE"} else {$KMSPriority = BoolToWStr $KMSPriority}

	if ($KMSPublishing -EQ "TRUE") {$KMSPublishing = "Enabled"} else {$KMSPublishing = "Disabled"}
	if ($KMSPriority -EQ "TRUE") {$KMSPriority = "Low"} else {$KMSPriority = "Normal"}

	if ($SLApp)
	{
		$propKMSServer | foreach { set $_ (SlGetInfoApp $strApp $_) }
	}
	else
	{
		$propKMSServer | foreach { set $_ (SlGetInfoService $_) }
	}

	$KMSRequests = $KeyManagementServiceTotalRequests
	$NoRequests = ($null -EQ $KMSRequests) -Or ($KMSRequests -EQ -1) -Or ($KMSRequests -EQ 4294967295)

	CONOUT "`nKey Management Service host information:"
	CONOUT "    Current count: $KeyManagementServiceCurrentCount"
	CONOUT "    Listening on Port: $KMSListening"
	CONOUT "    DNS publishing: $KMSPublishing"
	CONOUT "    KMS priority: $KMSPriority"
	if ($NoRequests) {
		return
	}
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

function DetectKmsClient
{
	if ($win8)
	{
		$VLType = strGetRegistry ($SPKeyPath + '\' + $strApp + '\' + $licID) "VLActivationType"
		if ($null -EQ $VLType) {$VLType = strGetRegistry ($SPKeyPath + '\' + $strApp) "VLActivationType"}
		if ($null -EQ $VLType) {$VLType = strGetRegistry ($SPKeyPath) "VLActivationType"}
		if ($null -EQ $VLType -Or $VLType -GT 3) {$VLType = 0}
	}
	if ($null -NE $VLType) {CONOUT "Configured Activation Type: $($VLActTypes[$VLType])"}

	CONOUT "`r"
	if ($LicenseStatus -NE 1) {
		CONOUT "Please activate the product in order to update KMS client information values."
		return
	}

	if ($NT7 -Or $strSLP -EQ $oslp) {
		$propKMSClient | foreach { set $_ (SlGetInfoSku $licID $_) }
		if ($strSLP -EQ $oslp) {$regk = $OPKeyPath} else {$regk = $SPKeyPath}
		$KMSCaching = strGetRegistry $regk "DisableKeyManagementServiceHostCaching"
		if (-Not $KMSCaching) {$KMSCaching = "TRUE"} else {$KMSCaching = BoolToWStr (!$KMSCaching)}
	}

	"ClientMachineID" | foreach { set $_ (SlGetInfoService $_) }

	if ($Vista) {
		$propKMSVista | foreach { set $_ (SlGetInfoService $_) }
		$KeyManagementServicePort = strGetRegistry $SLKeyPath "KeyManagementServicePort"
		$DiscoveredKeyManagementServiceName = strGetRegistry $NSKeyPath "DiscoveredKeyManagementServiceName"
		$DiscoveredKeyManagementServicePort = strGetRegistry $NSKeyPath "DiscoveredKeyManagementServicePort"
	}

	if ([String]::IsNullOrEmpty($KeyManagementServiceName)) {
		$KmsReg = $null
	} else {
		if (-Not $KeyManagementServicePort) {$KeyManagementServicePort = 1688}
		$KmsReg = "Registered KMS machine name: ${KeyManagementServiceName}:${KeyManagementServicePort}"
	}

	if ([String]::IsNullOrEmpty($DiscoveredKeyManagementServiceName)) {
		$KmsDns = "DNS auto-discovery: KMS name not available"
		if ($Vista -And -Not $Admin) {$KmsDns = "DNS auto-discovery: Run the script as administrator to retrieve info"}
	} else {
		if (-Not $DiscoveredKeyManagementServicePort) {$DiscoveredKeyManagementServicePort = 1688}
		$KmsDns = "KMS machine name from DNS: ${DiscoveredKeyManagementServiceName}:${DiscoveredKeyManagementServicePort}"
	}

	if ($null -NE $KMSCaching) {
		if ($KMSCaching -EQ "TRUE") {$KMSCaching = "Enabled"} else {$KMSCaching = "Disabled"}
	}

	if ($strSLP -EQ $wslp -And $NT9) {
		if ([String]::IsNullOrEmpty($DiscoveredKeyManagementServiceIpAddress)) {
			$DiscoveredKeyManagementServiceIpAddress = "not available"
		}
	}

	CONOUT "Key Management Service client information:"
	CONOUT "    Client Machine ID (CMID): $ClientMachineID"
	if ($null -EQ $KmsReg) {
		CONOUT "    $KmsDns"
		CONOUT "    Registered KMS machine name: KMS name not available"
	} else {
		CONOUT "    $KmsReg"
	}
	if ($null -NE $DiscoveredKeyManagementServiceIpAddress) {CONOUT "    KMS machine IP address: $DiscoveredKeyManagementServiceIpAddress"}
	CONOUT "    KMS machine extended PID: $CustomerPID"
	CONOUT "    Activation interval: $VLActivationInterval minutes"
	CONOUT "    Renewal interval: $VLRenewalInterval minutes"
	if ($null -NE $KMSCaching) {CONOUT "    KMS host caching: $KMSCaching"}
	if (-Not [String]::IsNullOrEmpty($KeyManagementServiceLookupDomain)) {CONOUT "    KMS SRV record lookup domain: $KeyManagementServiceLookupDomain"}
}

function GetResult($strSLP, $strApp, $entry)
{
	$licID = $entry.id
	$propPrd | foreach { set $_ (SlGetInfoSku $licID $_) }
	. SlGetInfoLicensing $strApp $licID
	$LicenseStatus = $dwStatus
	$LicReason = $hrReason
	$EvaluationEndDate = $qwValidity
	$gprMnt = $dwGrace

	$pkid = $entry.pk
	$isPPK = $null -NE $pkid

	$add_on = $Name.IndexOf("add-on for", 5)
	if ($add_on -NE -1) {
		$Name = $Name.Substring(0, $add_on + 7)
	}

	$licPHN = "empty"
	if ($Dlv -Or $All.IsPresent) {
		$licPHN = SlGetInfoSku $licID "msft:sl/EUL/PHONE/PUBLIC"
	}

	if ($LicenseStatus -EQ 0 -And !$isPPK) {
		& $isAll
		CONOUT "Name: $Name"
		CONOUT "Description: $Description"
		CONOUT "Activation ID: $licID"
		CONOUT "License Status: Unlicensed"
		if ($licPHN -NE "empty") {
			$gPHN = [String]::IsNullOrEmpty($licPHN) -NE $true
			CONOUT "Phone activatable: $($gPHN.ToString())"
		}
		return
	}

	$winID = ($strApp -EQ $winApp)
	$winPR = ($winID -And -Not $entry.ex)
	$Vista = ($winID -And $NT6 -And -Not $NT7)
	$NT5 = ($strSLP -EQ $wslp -And $winbuild -LT 6001)
	$win8 = ($strSLP -EQ $wslp -And $NT8)
	$reapp = ("Windows", "App")[!$winID]
	$prmnt = ("machine", "product")[!$winPR]

	if ($Description.Contains("VOLUME_KMSCLIENT")) {$cKmsClient = 1; $actTag = "Volume"}
	if ($Description.Contains("TIMEBASED_")) {$cTblClient = 1; $actTag = "Timebased"}
	if ($Description.Contains("VIRTUAL_MACHINE_ACTIVATION")) {$cAvmClient = 1; $actTag = "Automatic VM"}
	if ($null -EQ $cKmsClient -And $Description.Contains("VOLUME_KMS")) {$cKmsServer = 1}

	$gprDay = [Math]::Round($gprMnt/1440)
	$_xpr = ""
	$inGrace = $false
	if ($gprMnt -GT 0) {
		$_xpr = [DateTime]::Now.AddMinutes($gprMnt).ToString('yyyy-MM-dd hh:mm:ss tt')
		$inGrace = $true
	}

	$LicenseMsg = "Time remaining: $gprMnt minute(s) ($gprDay day(s))"
	if ($LicenseStatus -EQ 0) {
		$LicenseInf = "Unlicensed"
		$LicenseMsg = $null
	}
	if ($LicenseStatus -EQ 1) {
		$LicenseInf = "Licensed"
		if ($gprMnt -EQ 0) {
			$LicenseMsg = $null
			$ExpireMsg = "The $prmnt is permanently activated."
		} else {
			$LicenseMsg = "$actTag activation expiration: $gprMnt minute(s) ($gprDay day(s))"
			if ($inGrace) {$ExpireMsg = "$actTag activation will expire $_xpr"}
		}
	}
	if ($LicenseStatus -EQ 2) {
		$LicenseInf = "Initial grace period"
		if ($inGrace) {$ExpireMsg = "$LicenseInf ends $_xpr"}
	}
	if ($LicenseStatus -EQ 3) {
		$LicenseInf = "Additional grace period (KMS license expired or hardware out of tolerance)"
		if ($inGrace) {$ExpireMsg = "Additional grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 4) {
		$LicenseInf = "Non-genuine grace period"
		if ($inGrace) {$ExpireMsg = "$LicenseInf ends $_xpr"}
	}
	if ($LicenseStatus -EQ 5 -And -Not $NT5) {
		$LicenseReason = '0x{0:X}' -f $LicReason
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
		if ($inGrace) {$ExpireMsg = "$LicenseInf ends $_xpr"}
	}

	if ($isPPK) {
		$propPkey | foreach { set $_ (SlGetInfoPKey $pkid $_) }
	}

	if ($winPR -And $isPPK -And -Not $NT8) {
		$uxd = SlGetInfoSku $licID 'UXDifferentiator'
		$script:primary += @{
			aid = $licID;
			ppk = $PartialProductKey;
			chn = $Channel;
			lst = $LicenseStatus;
			lcr = $LicReason;
			ged = $gprMnt;
			evl = $EvaluationEndDate;
			dff = $uxd
		}
	}

	if ($IID -And $isPPK) {
		$OfflineInstallationId = SlGetInfoIID $licID
	}

	if ($Dlv) {
		if ($win8)
		{
			$RemainingSkuReArmCount = SlGetInfoSku $licID 'RemainingRearmCount'
			$RemainingAppReArmCount = SlGetInfoApp $strApp 'RemainingRearmCount'
		}
		else
		{
			if (($winID -And $NT7) -Or $strSLP -EQ $oslp)
			{
				$RemainingSLReArmCount = SlGetInfoApp $strApp 'RemainingRearmCount'
			}
			else
			{
				$RemainingSLReArmCount = SlGetInfoService 'RearmCount'
			}
		}
		if ($null -EQ $TrustedTime)
		{
			$TrustedTime = SlGetInfoSvcApp $strApp 'TrustedTime'
		}
	}

	& $isAll
	CONOUT "Name: $Name"
	CONOUT "Description: $Description"
	CONOUT "Activation ID: $licID"
	if ($null -NE $DigitalPID) {CONOUT "Extended PID: $DigitalPID"}
	if ($null -NE $DigitalPID2 -And $Dlv) {CONOUT "Product ID: $DigitalPID2"}
	if ($null -NE $OfflineInstallationId -And $IID) {CONOUT "Installation ID: $OfflineInstallationId"}
	if ($null -NE $Channel) {CONOUT "Product Key Channel: $Channel"}
	if ($null -NE $PartialProductKey) {CONOUT "Partial Product Key: $PartialProductKey"}
	CONOUT "License Status: $LicenseInf"
	if ($null -NE $LicenseMsg) {CONOUT "$LicenseMsg"}
	if ($LicenseStatus -NE 0 -And $EvaluationEndDate) {
		$EED = [DateTime]::FromFileTimeUtc($EvaluationEndDate).ToString('yyyy-MM-dd hh:mm:ss tt')
		CONOUT "Evaluation End Date: $EED UTC"
	}
	if ($LicenseStatus -NE 1 -And $licPHN -NE "empty") {
		$gPHN = [String]::IsNullOrEmpty($licPHN) -NE $true
		CONOUT "Phone activatable: $($gPHN.ToString())"
	}
	if ($Dlv) {
		if ($null -NE $RemainingSLReArmCount) {
			CONOUT "Remaining $reapp rearm count: $RemainingSLReArmCount"
		}
		if ($null -NE $RemainingSkuReArmCount) {
			CONOUT "Remaining $reapp rearm count: $RemainingAppReArmCount"
			CONOUT "Remaining SKU rearm count: $RemainingSkuReArmCount"
		}
		if ($LicenseStatus -NE 0 -And $TrustedTime) {
			$TTD = [DateTime]::FromFileTime($TrustedTime).ToString('yyyy-MM-dd hh:mm:ss tt')
			CONOUT "Trusted time: $TTD"
		}
	}
	if (!$isPPK) {
		return
	}

	if ($win8 -And $VLActivationType -EQ 1) {
		DetectAdbaClient
	}

	if ($winID -And $null -NE $cAvmClient) {
		DetectAvmClient
	}

	$chkSub = ($winPR -And $isSub)

	$chkSLS = ($null -NE $cKmsClient -Or $null -NE $cKmsServer -Or $chkSub)

	if (!$chkSLS) {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
		return
	}

	if ($null -NE $cKmsClient) {
		DetectKmsClient
	}

	if ($null -NE $cKmsServer) {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
		DetectKmsHost
	} else {
		if ($null -NE $ExpireMsg) {CONOUT "`n    $ExpireMsg"}
	}

	if ($chkSub) {
		DetectSubscription
	}

}

function ParseList($strSLP, $strApp, $arrList)
{
	foreach ($entry in $arrList)
	{
		GetResult $strSLP $strApp $entry
		CONOUT "$line3"
		& $noAll
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
	If ($null -Eq $tokenFiles -Or $tokenFiles.Length -Eq 0)
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
	$fNUL = ([IO.Directory]::Exists("${env:LOCALAPPDATA}\Microsoft\Office\Licenses")) -and ([IO.Directory]::GetFiles("${env:LOCALAPPDATA}\Microsoft\Office\Licenses", "*", 1).Length -GT 0)
	$fDev = ([IO.Directory]::Exists("${env:PROGRAMDATA}\Microsoft\Office\Licenses")) -and ([IO.Directory]::GetFiles("${env:PROGRAMDATA}\Microsoft\Office\Licenses", "*", 1).Length -GT 0)
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
	$pqwLastTime = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-LastWindowsActivationTime",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pqwLastTime
	)) {
		return $FALSE
	}

	$actTime = $Marshal::ReadInt64($pqwLastTime)
	if ($actTime -ne 0) {
		CONOUT ("    LastActivationTime={0}" -f [DateTime]::FromFileTimeUtc($actTime).ToString("yyyy/MM/dd:HH:mm:ss"))
	}

	$Marshal::FreeHGlobal($pqwLastTime)
	return $TRUE
}

function PrintIsWindowsGenuine {
	$dwGenuine = 0

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

#region clc
function clcGetExpireKrn
{
	$tData = 0
	$cData = 0
	$bData = 0

	$hrRet = $Win32::SLGetWindowsInformation(
		"Kernel-ExpirationDate",
		[ref]$tData,
		[ref]$cData,
		[ref]$bData
	)

	if ($hrRet -Or !$cData -Or $tData -NE 3)
	{
		return $null
	}

	$year = $Marshal::ReadInt16($bData, 0)
	if ($year -EQ 0 -Or $year -EQ 1601)
	{
		$rData = $null
	}
	else
	{
		$rData = '{0}/{1}/{2}:{3}:{4}:{5}' -f $year, $Marshal::ReadInt16($bData, 2), $Marshal::ReadInt16($bData, 4), $Marshal::ReadInt16($bData, 6), $Marshal::ReadInt16($bData, 8), $Marshal::ReadInt16($bData, 10)
	}

	#$Marshal::FreeHGlobal($bData)
	return $rData
}

function clcGetExpireSys
{
	$kuser = $Marshal::ReadInt64((New-Object IntPtr(0x7FFE02C8)))

	if ($kuser -EQ 0)
	{
		return $null
	}

	$rData = [DateTime]::FromFileTimeUtc($kuser).ToString('yyyy/MM/dd:HH:mm:ss')
	return $rData
}

function clcGetLicensingState($dwState)
{
	if ($dwState -EQ 5) {
		$dwState = 3
	} elseif ($dwState -EQ 3 -Or $dwState -EQ 4 -Or $dwState -EQ 6) {
		$dwState = 2
	} elseif ($dwState -GT 6) {
		$dwState = 4
	}

	$rData = '{0}' -f $ppwszLicensingStates[$dwState]
	return $rData
}

function clcGetGenuineState($AppId)
{
	$dwGenuine = 0

	if ($NT7) {
		$hrRet = $Win32::SLIsWindowsGenuineLocal([ref]$dwGenuine)
	} else {
		$hrRet = $Win32::SLIsGenuineLocal([ref][Guid]$AppId, [ref]$dwGenuine, 0)
	}

	if ($hrRet)
	{
		$dwGenuine = 4
	}

	if ($dwGenuine -LT 5) {
		$rData = '{0}' -f $ppwszGenuineStates[$dwGenuine]
	} else {
		$rData = $dwGenuine
	}
	return $rData
}

function ClcRun
{
	$prs = $script:primary[0]
	if ($null -EQ $prs) {
		return
	}

	$lState = clcGetLicensingState $prs.lst
	$uState = clcGetGenuineState $winApp
	$TbbKrn = clcGetExpireKrn
	$TbbSys = clcGetExpireSys
	if ($null -NE $TbbKrn) {
		$ked = $TbbKrn
	} elseif ($null -NE $TbbSys) {
		$ked = $TbbSys
	}

	& $isAll
	CONOUT "Client Licensing Check information:"

	CONOUT ("    AppId={0}" -f $winApp)
	if ($prs.ged) { CONOUT ("    GraceEndDate={0}" -f ([DateTime]::UtcNow.AddMinutes($prs.ged).ToString('yyyy/MM/dd:HH:mm:ss'))) }
	if ($null -NE $ked) { CONOUT ("    KernelTimebombDate={0}" -f $ked) }
	CONOUT ("    LastConsumptionReason=0x{0:x8}" -f $prs.lcr)
	if ($prs.evl) { CONOUT ("    LicenseExpirationDate={0}" -f ([DateTime]::FromFileTimeUtc($prs.evl).ToString('yyyy/MM/dd:HH:mm:ss'))) }
	CONOUT ("    LicenseState={0}" -f $lState)
	CONOUT ("    PartialProductKey={0}" -f $prs.ppk)
	CONOUT ("    ProductKeyType={0}" -f $prs.chn)
	CONOUT ("    SkuId={0}" -f $prs.aid)
	CONOUT ("    uxDifferentiator={0}" -f $prs.dff)
	CONOUT ("    IsWindowsGenuine={0}" -f $uState)

	CONOUT "$line3"
	& $noAll
}
#endregion

$Host.UI.RawUI.WindowTitle = "Check Activation Status"
if ($All.IsPresent) {
	$B=$Host.UI.RawUI.BufferSize;$B.Height=3000;$Host.UI.RawUI.BufferSize=$B;
	if (!$Pass.IsPresent) {clear;}
}

$SysPath = "$env:SystemRoot\System32"
if (Test-Path "$env:SystemRoot\Sysnative\reg.exe") {
	$SysPath = "$env:SystemRoot\Sysnative"
}

$wslp = "SoftwareLicensingProduct"
$wsls = "SoftwareLicensingService"
$oslp = "OfficeSoftwareProtectionProduct"
$osls = "OfficeSoftwareProtectionService"
$winApp = "55c92734-d682-4d71-983e-d6ec3f16059f"
$o14App = "59a52881-a989-479d-af46-f275c6370663"
$o15App = "0ff1ce15-a989-479d-af46-f275c6370663"
$isSub = ($winbuild -GE 26000) -And (Select-String -Path "$SysPath\wbem\sppwmi.mof" -Encoding unicode -Pattern "SubscriptionType")
$DllDigital = ($winbuild -GE 14393) -And (Test-Path "$SysPath\EditionUpgradeManagerObj.dll")
$DllSubscription = ($winbuild -GE 14393) -And (Test-Path "$SysPath\Clipc.dll")
$VLActTypes = @("All", "AD", "KMS", "Token")
$OPKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
$SPKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
$SLKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"
$NSKeyPath = "HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"
$propPrd = 'Name', 'Description', 'TrustedTime', 'VLActivationType'
$propPkey = 'PartialProductKey', 'Channel', 'DigitalPID', 'DigitalPID2'
$propKMSServer = 'KeyManagementServiceCurrentCount', 'KeyManagementServiceTotalRequests', 'KeyManagementServiceFailedRequests', 'KeyManagementServiceUnlicensedRequests', 'KeyManagementServiceLicensedRequests', 'KeyManagementServiceOOBGraceRequests', 'KeyManagementServiceOOTGraceRequests', 'KeyManagementServiceNonGenuineGraceRequests', 'KeyManagementServiceNotificationRequests'
$propKMSClient = 'CustomerPID', 'KeyManagementServiceName', 'KeyManagementServicePort', 'DiscoveredKeyManagementServiceName', 'DiscoveredKeyManagementServicePort', 'DiscoveredKeyManagementServiceIpAddress', 'VLActivationInterval', 'VLRenewalInterval', 'KeyManagementServiceLookupDomain'
$propKMSVista  = 'CustomerPID', 'KeyManagementServiceName', 'VLActivationInterval', 'VLRenewalInterval'
$propADBA = 'ADActivationObjectName', 'ADActivationObjectDN', 'ADActivationCsvlkPID', 'ADActivationCsvlkSkuID'
$propAVMA = 'InheritedActivationId', 'InheritedActivationHostMachineName', 'InheritedActivationHostDigitalPid2', 'InheritedActivationActivationTime'
$primary = @()
$ppwszGenuineStates = @(
	"SL_GEN_STATE_IS_GENUINE",
	"SL_GEN_STATE_INVALID_LICENSE",
	"SL_GEN_STATE_TAMPERED",
	"SL_GEN_STATE_OFFLINE",
	"SL_GEN_STATE_LAST"
)
$ppwszLicensingStates = @(
	"SL_LICENSING_STATUS_UNLICENSED",
	"SL_LICENSING_STATUS_LICENSED",
	"SL_LICENSING_STATUS_IN_GRACE_PERIOD",
	"SL_LICENSING_STATUS_NOTIFICATION",
	"SL_LICENSING_STATUS_LAST"
)

'cW1nd0ws', 'c0ff1ce15', 'c0ff1ce14', 'ospp14', 'ospp15' | foreach {set $_ @()}

$offsvc = "osppsvc"
if ($NT7 -Or -Not $NT6) {$winsvc = "sppsvc"} else {$winsvc = "slsvc"}

try {gsv $winsvc -EA 1 | Out-Null; $WsppHook = 1} catch {$WsppHook = 0}
try {gsv $offsvc -EA 1 | Out-Null; $OsppHook = 1} catch {$OsppHook = 0}

if (Test-Path "$SysPath\sppc.dll") {
	$SLdll = 'sppc.dll'
} elseif (Test-Path "$SysPath\slc.dll") {
	$SLdll = 'slc.dll'
} else {
	$WsppHook = 0
}

if ($OsppHook -NE 0) {
	$OLdll = (strGetRegistry $OPKeyPath "Path") + 'osppc.dll'
	if (!(Test-Path "$OLdll")) {$OsppHook = 0}
}

if ($WsppHook -NE 0) {
	if ($NT6 -And -Not $NT7 -And -Not $Admin) {
		if ($null -EQ [Diagnostics.Process]::GetProcessesByName("$winsvc")[0].ProcessName) {$WsppHook = 0; CONOUT "`nError: failed to start $winsvc Service.`n"}
	} else {
		try {sasv $winsvc -EA 1} catch {$WsppHook = 0; CONOUT "`nError: failed to start $winsvc Service.`n"}
	}
}

if ($WsppHook -NE 0) {
	. InitializePInvoke $SLdll $false
	$hSLC = 0
	[void]$Win32::SLOpen([ref]$hSLC)

	$cW1nd0ws  = SlGetInfoSLID $winApp
	$c0ff1ce15 = SlGetInfoSLID $o15App
	$c0ff1ce14 = SlGetInfoSLID $o14App
}

if ($cW1nd0ws.Count -GT 0)
{
	echoWindows
	ParseList $wslp $winApp $cW1nd0ws
}
elseif ($NT6)
{
	echoWindows
	CONOUT "Error: product key not found.`n"
}

if ($NT6 -And -Not $NT8) {
	ClcRun
}

if ($NT8) {
	ClicRun
}

$doMSG = 1

if ($c0ff1ce15.Count -GT 0)
{
	CheckOhook
	echoOffice
	ParseList $wslp $o15App $c0ff1ce15
}

if ($c0ff1ce14.Count -GT 0)
{
	echoOffice
	ParseList $wslp $o14App $c0ff1ce14
}

if ($hSLC) {
	[void]$Win32::SLClose($hSLC)
}

if ($OsppHook -NE 0) {
	try {sasv $offsvc -EA 1} catch {$OsppHook = 0; CONOUT "`nError: failed to start $offsvc Service.`n"}
}

if ($OsppHook -NE 0) {
	. InitializePInvoke "$OLdll" $true
	$hSLC = 0
	[void]$Win32::SLOpen([ref]$hSLC)

	$ospp15 = SlGetInfoSLID $o15App
	$ospp14 = SlGetInfoSLID $o14App
}

if ($ospp15.Count -GT 0)
{
	echoOffice
	ParseList $oslp $o15App $ospp15
}

if ($ospp14.Count -GT 0)
{
	echoOffice
	ParseList $oslp $o14App $ospp14
}

if ($hSLC) {
	[void]$Win32::SLClose($hSLC)
}

if ($NT7) {
	vNextDiagRun
}

ExitScript 0
:sppmgr:
