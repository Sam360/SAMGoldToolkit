 ##########################################################################
 #
 # Sam360UISync
 # SAM Gold Toolkit
 # Original Source: Sam360
 #
 ##########################################################################
 
 <#
.SYNOPSIS
#>

 Param(
    [alias("log")]
    [string] $LogFile = "$($env:Temp)\LogFile.txt",
    [string] $UserName,
    [string] $Password,
    [string] $ClientOrganisationId = "0",
    [string] $APIServer = "https://api.sam360.com",
    [string] $UIDatabaseServerName, 
    [string] $UIDatabaseName)

function LogText {
    param(
        [Parameter(Position=0, ValueFromRemainingArguments=$true, ValueFromPipeline=$true)]
        [Object] $Object,
        [System.ConsoleColor]$color = [System.Console]::ForegroundColor,
        [switch]$NoNewLine = $false  
    )

    # Display text on screen
    Write-Host -Object $Object -ForegroundColor $color -NoNewline:$NoNewLine

    if ($LogFile) {
        $Object | Out-File $LogFile -Encoding utf8 -Append 
    }
}

function LogError([string[]]$errorDescription){
    if ($Verbose){
        LogText ""
    }

    $output = Get-Date -Format HH:mm:ss.ff
    $output += " - "
    $output += $errorDescription -join "`r`n              "
    LogText $output -Color Red
    Start-Sleep -s 3
}

function InitialiseLogFile {
    if ($LogFile -and (Test-Path $LogFile)) {
        Remove-Item $LogFile
    }
}

function LogLastException() {
    $currentException = $Error[0].Exception;

    while ($currentException)
    {
        LogText -Color Red $currentException
        LogText -Color Red $currentException.Data
        LogText -Color Red $currentException.HelpLink
        LogText -Color Red $currentException.HResult
        LogText -Color Red $currentException.Message
        LogText -Color Red $currentException.Source
        LogText -Color Red $currentException.StackTrace
        LogText -Color Red $currentException.TargetSite

        $currentException = $currentException.InnerException
    }

    Start-Sleep -s 3
}

function LogProgress($progressDescription){
	$output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
	LogText $output -Color Green
}

function GetScriptPath
{
    if($PSCommandPath){
        return $PSCommandPath; }
        
    if($MyInvocation.ScriptName){
        return $MyInvocation.ScriptName }
        
    if($script:MyInvocation.MyCommand.Path){
        return $script:MyInvocation.MyCommand.Path }

    return $script:MyInvocation.MyCommand.Definition
}

function QueryUser([string]$Message, [string]$Prompt, [switch]$AsSecureString = $false, [string]$DefaultValue){
    $strResult = ""
    
    if ($Message) {
        LogText $Message -color Yellow
    }

    if ($DefaultValue) {
        $Prompt += " (Default [$DefaultValue])"
    }

    $Prompt += ": "
    LogText $Prompt -color Yellow -NoNewLine
    
    if ($Headless) {
        LogText " (Headless - Using Default Value)" -color Yellow
    }
    else {
        $strResult = Read-Host -AsSecureString:$AsSecureString
    }

    if(!$strResult) {
        $strResult = $DefaultValue
        if ($AsSecureString) {
            $strResult = ConvertTo-SecureString $strResult -AsPlainText -Force
        }
    }

    return $strResult
}

function Get-ConsoleCredential([String] $Message, [String] $DefaultUsername)
{
    $strUsername = QueryUser -Message $Message -Prompt "Username" -DefaultValue $DefaultUsername
    if (!$strUsername){
        return $null
    }

    $strSecurePassword = QueryUser -Prompt "Password" -AsSecureString
    if (!$strSecurePassword){
        return $null
    }

    $bstrSecurePassword = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($strSecurePassword)
    $strUnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrSecurePassword)

    $Creds = New-Object PSObject
    $Creds | Add-Member -MemberType NoteProperty -Name "UserName" -Value $strUsername
    $Creds | Add-Member -MemberType NoteProperty -Name "Password" -Value $strUnsecurePassword

    return $Creds
}

function LogEnvironmentDetails {
    LogText -Color Gray " "
    LogText -Color Gray "   _____         __  __    _____       _     _   _______          _ _    _ _   "
    LogText -Color Gray "  / ____|  /\   |  \/  |  / ____|     | |   | | |__   __|        | | |  (_) |  "
    LogText -Color Gray " | (___   /  \  | \  / | | |  __  ___ | | __| |    | | ___   ___ | | | ___| |_ "
    LogText -Color Gray "  \___ \ / /\ \ | |\/| | | | |_ |/ _ \| |/ _`` |    | |/ _ \ / _ \| | |/ / | __|"
    LogText -Color Gray "  ____) / ____ \| |  | | | |__| | (_) | | (_| |    | | (_) | (_) | |   <| | |_ "
    LogText -Color Gray " |_____/_/    \_\_|  |_|  \_____|\___/|_|\__,_|    |_|\___/ \___/|_|_|\_\_|\__|"
    LogText -Color Gray " "
    LogText -Color Gray " Get-Sam360Inventory.ps1"
    LogText -Color Gray " "

    $OSDetails = Get-WmiObject Win32_OperatingSystem
    $ScriptPath = GetScriptPath
    $Elevated = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
    LogText -Color Gray "Computer Name:        $($env:COMPUTERNAME)"
    LogText -Color Gray "User Name:            $($env:USERNAME)@$($env:USERDNSDOMAIN)"
    LogText -Color Gray "Windows Version:      $($OSDetails.Caption)($($OSDetails.Version))"
    LogText -Color Gray "PowerShell Host:      $($host.Version.Major)"
    LogText -Color Gray "PowerShell Version:   $($PSVersionTable.PSVersion)"
    LogText -Color Gray "PowerShell Word size: $($([IntPtr]::size) * 8) bit"
    LogText -Color Gray "CLR Version:          $($PSVersionTable.CLRVersion)"
    LogText -Color Gray "Elevated:             $Elevated"
    LogText -Color Gray "Current Date Time:    $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
    LogText -Color Gray "Script Path:          $ScriptPath"
    LogText -Color Gray "API User Name:        $UserName"
    LogText -Color Gray "API Server:           $APIServer"
    LogText -Color Gray ""
}

function SetupDateFormats {
    # Standardise date/time output to ISO 8601'ish format
    $bDateFormatConfigured = $false
    $currentThread = [System.Threading.Thread]::CurrentThread
    
    try {
        $CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = 'yyyy-MM-dd'
        $CurrentThread.CurrentCulture.DateTimeFormat.LongDatePattern = 'yyyy-MM-dd HH:mm:ss'
        $bDateFormatConfigured = $true
    }
    catch {
    }

    if (!($bDateFormatConfigured)) {
        try {
            $cultureCopy = $CurrentThread.CurrentCulture.Clone()
            $cultureCopy.DateTimeFormat.ShortDatePattern = 'yyyy-MM-dd'
            $cultureCopy.DateTimeFormat.LongDatePattern = 'yyyy-MM-dd HH:mm:ss'
            $currentThread.CurrentCulture = $cultureCopy
        }
        catch {
        }
    }
}

$MappingDevices = @{
	EntityName = "Device";
	Table = "tblDevices";
	Sam360ReportId = "UniversalInventoryUniversalInventoryDevicesTable";
	FieldMappings = @(
		@{DBFieldName="SrcId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="Site"; DBFieldType="varchar"; DBFieldWidth=32; S3APIFieldName="DeviceADSite";},
		@{DBFieldName="Type"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceComputerType";},
		@{DBFieldName="Name"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceName";},
		@{DBFieldName="DnsHostName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceName";},
		@{DBFieldName="Fqdn"; DBFieldType="nvarchar"; DBFieldWidth=512; S3APIFieldName="DeviceFQDN";},
		@{DBFieldName="Domain"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceDomain";},
		@{DBFieldName="PrimaryUserName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="LastUserofDeviceWindowsUserDomainName";},
		@{DBFieldName="Description"; DBFieldType="nvarchar"; DBFieldWidth=512; S3APIFieldName="DeviceADDescription";},
		@{DBFieldName="LastBoot"; DBFieldType="datetime"; S3APIFieldName="DeviceLastBootTime";},
		@{DBFieldName="LastScan"; DBFieldType="datetime"; S3APIFieldName="DeviceLastUpdate";},
		@{DBFieldName="Manufacturer"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceSystemManufacturer";},
		@{DBFieldName="Model"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceSystemProductName";},
		@{DBFieldName="TotalMemoryMb"; DBFieldType="bigint"; S3APIFieldName="DeviceTotalRAM";},
		@{DBFieldName="TotalDiskSpaceMB"; DBFieldType="bigint"; S3APIFieldName="DeviceTotalDisksSizeGB";},
		@{DBFieldName="SerialNumber"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceSerialNumber";},
		@{DBFieldName="Virtual"; DBFieldType="bit"; S3APIFieldName="DeviceIsVirtual";},
		@{DBFieldName="WmiStatus"; DBFieldType="nvarchar"; DBFieldWidth=4000; S3APIFieldName="DeviceWMIConnectResult";},
		@{DBFieldName="SamAccountName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceADSAMAccountName";},
		@{DBFieldName="SamAccountType"; DBFieldType="int"; S3APIFieldName="DeviceADSAMAccountType"; },
		@{DBFieldName="Enabled"; DBFieldType="bit"; S3APIFieldName="DeviceExcludeDevice"; ConverterFn="GetInverseBoolValue"},
		@{DBFieldName="PwdRequired"; DBFieldType="bit"; S3APIFieldName="DeviceADUACText"; ConverterFn="GetPasswordRequired"},
		@{DBFieldName="PwdCanChange"; DBFieldType="bit"; S3APIFieldName="DeviceADUACText"; ConverterFn="GetPasswordCanChange"},
		@{DBFieldName="PwdExpires"; DBFieldType="bit"; S3APIFieldName="DeviceADUACText"; ConverterFn="GetPasswordExpires"},
		@{DBFieldName="NormalAccount"; DBFieldType="bit"; S3APIFieldName="DeviceADUACText"; ConverterFn="GetIsNormalAccount"},
		@{DBFieldName="BiosCaption"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceBIOSDescription";},
		@{DBFieldName="BiosManufacturer"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceBIOSManufacturer";},
		@{DBFieldName="BiosReleaseDate"; DBFieldType="datetime"; S3APIFieldName="DeviceBIOSReleaseDate";},
		@{DBFieldName="BiosSerialNumber"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceBIOSSerialNumber";},
		@{DBFieldName="BiosMsdmTable"; DBFieldType="int"; S3APIFieldName="DeviceMSDMLicenseOEMID"; ConverterFn="GetMsdmTableExists"},
		@{DBFieldName="BiosVersion"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceBIOSVersion";},
		@{DBFieldName="BaseBoardManufacturer"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceBaseBoardManufacturer";},
		@{DBFieldName="BaseBoardProduct"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceBaseBoardProduct";},
		@{DBFieldName="BaseBoardSerialNumber"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceBaseBoardSerialNumber";},
		@{DBFieldName="BaseBoardVersion"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceBaseBoardVersion";},
		@{DBFieldName="OsArchitectureBits"; DBFieldType="tinyint"; S3APIFieldName="DeviceProcessorArchitecture";},
		@{DBFieldName="OsFamily"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceOSFamily";},
		@{DBFieldName="OsInstallDate"; DBFieldType="datetime"; S3APIFieldName="DeviceOSInstallationDate";},
		@{DBFieldName="OsLanguage"; DBFieldType="bigint"; S3APIFieldName="DeviceOSLanguage";},
		@{DBFieldName="OsManufacturer"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceOSVendor";},
		@{DBFieldName="OsName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceOS";},
		@{DBFieldName="OsRegisteredUser"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceOSRegisteredUser";},
		@{DBFieldName="OsSerialNumber"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceOSSerialNumber";},
		@{DBFieldName="OsServicePackMajorVersion"; DBFieldType="nvarchar"; DBFieldWidth=640; S3APIFieldName="DeviceOSServicePack";}, 
		@{DBFieldName="OsServicePackMinorVersion"; DBFieldType="nvarchar"; DBFieldWidth=640; S3APIFieldName="DeviceOSSPMinorVersion";},
		@{DBFieldName="OsSku"; DBFieldType="bigint"; S3APIFieldName="DeviceOSSKU";},
		@{DBFieldName="OsVersion"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceOSVersion";},
		@{DBFieldName="AdCn"; DBFieldType="nvarchar"; DBFieldWidth=64; S3APIFieldName="DeviceADCommonName";},
		@{DBFieldName="AdCompany"; DBFieldType="nvarchar"; DBFieldWidth=64; S3APIFieldName="DeviceADCompany";},
		@{DBFieldName="AdCreated"; DBFieldType="datetime"; S3APIFieldName="DeviceADWhenCreated";},
		@{DBFieldName="AdDescription"; DBFieldType="nvarchar"; DBFieldWidth=1024; S3APIFieldName="DeviceADDescription";},
		@{DBFieldName="AdDistinguishedName"; DBFieldType="nvarchar"; DBFieldWidth=4000; S3APIFieldName="DeviceADPath";},
		@{DBFieldName="AdLastLogon"; DBFieldType="datetime"; S3APIFieldName="DeviceADLastLogonDate";},
		@{DBFieldName="AdLocation"; DBFieldType="nvarchar"; DBFieldWidth=1024; S3APIFieldName="DeviceADLocation";},
		@{DBFieldName="AdLogonCount"; DBFieldType="int"; S3APIFieldName="DeviceADLogonCount";},
		@{DBFieldName="AdModified"; DBFieldType="datetime"; S3APIFieldName="DeviceADWhenChanged";},
		@{DBFieldName="AdCountryCode"; DBFieldType="nvarchar"; DBFieldWidth=64; S3APIFieldName="DeviceADCountryCode";},
		@{DBFieldName="AdPwdLastSet"; DBFieldType="datetime"; S3APIFieldName="DeviceADPasswordLastSet";}
	)
}

$MappingUsers = @{
	EntityName = "User";
	Table = "tblUsers";
	Sam360ReportId = "UniversalInventoryUniversalInventoryUsersTable";
	FieldMappings = @(
		@{DBFieldName="SrcId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="UserAnonymisedUserName";},
		@{DBFieldName="Site"; DBFieldType="varchar"; DBFieldWidth=32; S3APIFieldName="UserADDepartment";},
		@{DBFieldName="Cn"; DBFieldType="nvarchar"; DBFieldWidth=64; S3APIFieldName="UserADCommonName";},
		@{DBFieldName="DisplayName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="UserDisplayName";},
		@{DBFieldName="UserPrincipalName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="UserADUserPrincipalName";},
		@{DBFieldName="Domain"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="UserWindowsUserDomain";},
		@{DBFieldName="SamAccountName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="UserADSAMAccountName";},
		@{DBFieldName="SamAccountType"; DBFieldType="int"; S3APIFieldName="UserADSAMAccountType";},
		@{DBFieldName="Created"; DBFieldType="datetime"; S3APIFieldName="UserADWhenCreated";},
		@{DBFieldName="Modified"; DBFieldType="datetime"; S3APIFieldName="UserADWhenChanged";},
		@{DBFieldName="Enabled"; DBFieldType="bit"; S3APIFieldName="UserExcludeUser"; ConverterFn="GetInverseBoolValue"},
		@{DBFieldName="PwdRequired"; DBFieldType="bit"; S3APIFieldName="UserADUACText"; ConverterFn="GetPasswordRequired"},
		@{DBFieldName="PwdCanChange"; DBFieldType="bit"; S3APIFieldName="UserADUACText"; ConverterFn="GetPasswordCanChange"},
		@{DBFieldName="PwdExpires"; DBFieldType="bit"; S3APIFieldName="UserADUACText"; ConverterFn="GetPasswordExpires"},
		@{DBFieldName="NormalAccount"; DBFieldType="bit"; S3APIFieldName="UserADUACText"; ConverterFn="GetIsNormalAccount"},
		@{DBFieldName="AdDistinguishedName"; DBFieldType="nvarchar"; DBFieldWidth=4000; S3APIFieldName="UserADDistinguishedName";},
		@{DBFieldName="AdLastLogon"; DBFieldType="datetime"; S3APIFieldName="UserADLastLogonDate";},
		@{DBFieldName="LogonCount"; DBFieldType="int"; S3APIFieldName="UserADLogonCount";}
	)
}

$MappingSoftwareInstallations = @{
	EntityName = "Software Installation";
	Table = "tblSoftware";
	Sam360ReportId = "UniversalInventoryUniversalInventorySoftwareTable";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="Publisher"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductInstallationAddRemoveProgramsManufacturer";},
		@{DBFieldName="DiscoveredProduct"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductInstallationAddRemoveProgramsName";},
		@{DBFieldName="DiscoveredVersion"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductInstallationAddRemoveProgramsVersion";},
		@{DBFieldName="InstallLocation"; DBFieldType="nvarchar"; DBFieldWidth=1024; S3APIFieldName="ProductInstallationInstallPath";},
		@{DBFieldName="InstallDate"; DBFieldType="datetime"; S3APIFieldName="ProductInstallationInstallDate";},
		@{DBFieldName="SoftwareCode"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductInstallationProductGUID";}
		#@{DBFieldName="SwidUniqueId"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="";},
		#@{DBFieldName="SwidLicensorId"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="";}
	)
}

$MappingCPUs = @{
	EntityName = "CPU";
	Table = "tblCpus";
	Sam360ReportId = "UniversalInventoryUniversalInventoryCPUsTable";
	PreProcessFn = "PreProcessCPUs";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="DeviceId"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceProcessorDeviceID";},
		@{DBFieldName="DataWidth"; DBFieldType="tinyint"; S3APIFieldName="DeviceProcessorDataWidth";},
		@{DBFieldName="Manufacturer"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceProcessorManufacturer";},
		@{DBFieldName="MaxClockSpeed"; DBFieldType="bigint"; S3APIFieldName="DeviceProcessorSpeed";},
		@{DBFieldName="Name"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceProcessorName";},
		@{DBFieldName="NumberOfCores"; DBFieldType="int"; S3APIFieldName="DeviceCoresPerProcessor";},
		@{DBFieldName="NumberOfLogicalProcessors"; DBFieldType="int"; S3APIFieldName="DeviceLogicalProcessors";},
		@{DBFieldName="HyperThreadingCapable"; DBFieldType="bit"; S3APIFieldName="DeviceProcessorHyperThreading";}
	)
}

$MappingDatabaseServers = @{
	EntityName = "Database Server";
	Table = "tblDatabaseServers";
	Sam360ReportId = "UniversalInventoryUniversalInventoryDatabaseServersTable";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="SrcId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="ProductInstallationGlobalID"; ConverterFn="GetDBServerSrcID"},
		@{DBFieldName="InstallLocation"; DBFieldType="nvarchar"; DBFieldWidth=1024; S3APIFieldName="ProductInstallationInstallPath";},
		@{DBFieldName="ProductName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductName";},
		@{DBFieldName="Version"; DBFieldType="varchar"; DBFieldWidth=64; S3APIFieldName="ProductVersion";},
		@{DBFieldName="Edition"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductEdition";},
		@{DBFieldName="ServicePack"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="SQLServiceSPLevel";},
		@{DBFieldName="IsIntegratedSecurity"; DBFieldType="bit"; S3APIFieldName="SQLServiceLoginMode"; ConverterFn="GetSQLServiceLoginMode"},
		@{DBFieldName="DataPath"; DBFieldType="nvarchar"; DBFieldWidth=1024; S3APIFieldName="SQLServiceDataPath";},
		@{DBFieldName="ServiceName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="SQLServiceServiceName";},
		@{DBFieldName="State"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="SQLServiceState";},
		@{DBFieldName="StartMode"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="SQLServiceStartMode";},
		@{DBFieldName="Clustered"; DBFieldType="bit"; S3APIFieldName="ProductInstallationSQLIsClustered"; ConverterFn="GetBoolFromYesNo"},
		@{DBFieldName="InstanceId"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="SQLServiceServiceName";}, #PlaceHolder
		@{DBFieldName="InstanceName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="SQLServiceInstanceID";}, 
		@{DBFieldName="Sku"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="SQLServiceSKU";},
		@{DBFieldName="SqlServiceType"; DBFieldType="int"; S3APIFieldName="SQLServiceServiceType"; ConverterFn="GetSQLServiceType"},
		@{DBFieldName="Environment"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceEnvironmentType";},
		@{DBFieldName="Isv"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductInstallationSQLISVLicensor";}
	)
}

$MappingVMs = @{
	EntityName = "VM";
	Table = "tblVms";
	Sam360ReportId = "UniversalInventoryUniversalInventoryVmsTable";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="SrcVmHostId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceVirtualisationHostAnonymisedDeviceName";},
		@{DBFieldName="PowerState"; DBFieldType="varchar"; DBFieldWidth=16; S3APIFieldName="DeviceVMState";},
		@{DBFieldName="MemoryReservation"; DBFieldType="int"; S3APIFieldName="DeviceTotalRAM"; ConverterFn="GetMBValueFromGB"},
		@{DBFieldName="CpuReservation"; DBFieldType="int"; S3APIFieldName="DeviceProcessors";},
		@{DBFieldName="CpuAssigned"; DBFieldType="int"; S3APIFieldName="DeviceProcessors";},
		@{DBFieldName="Name"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceName";},
		@{DBFieldName="DnsName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceDomainFQN";}
	)
}

$MappingVMHosts = @{
	EntityName = "VM Host";
	Table = "tblVmHosts";
	Sam360ReportId = "UniversalInventoryUniversalInventoryVmHostsTable";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="SrcId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="FarmName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceClusterName";},
		@{DBFieldName="Name"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceName";},
		@{DBFieldName="MigrationEnabled"; DBFieldType="bit"; S3APIFieldName="DeviceMigrationEnabled";},
		@{DBFieldName="OsVersion"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceOS";},
		@{DBFieldName="NumberOfCores"; DBFieldType="int"; S3APIFieldName="DeviceCores";}
	)
}

$MappingMailServers = @{
	EntityName = "Mail Server";
	Table = "tblMailServers";
	Sam360ReportId = "UniversalInventoryUniversalInventoryMailServersTable";
	FieldMappings = @(
		@{DBFieldName="SrcId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="ServerName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceName";},
		@{DBFieldName="InstallPath"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductInstallationInstallPath";},
		@{DBFieldName="ProductName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductName";},
		@{DBFieldName="Version"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductVersion";},
		@{DBFieldName="Edition"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="ProductEdition";},
		@{DBFieldName="Enterprise"; DBFieldType="bit"; S3APIFieldName="ExchangeServerAttributesEnterpriseFeaturesEnabled";},
		@{DBFieldName="ServicePack"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceOSServicePack";}
	)
}

$MappingActiveSyncDevices = @{
	EntityName = "Active Sync Device";
	Table = "tblActiveSync";
	Sam360ReportId = "UniversalInventoryUniversalInventoryActiveSyncTable";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceID"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="ActiveSyncDeviceAnonymisedDeviceName";},
		@{DBFieldName="ActiveSyncVersion"; DBFieldType="varchar"; DBFieldWidth=16; S3APIFieldName="ActiveSyncDeviceActiveSyncVersion";},
		@{DBFieldName="FirstSyncTime"; DBFieldType="datetime"; S3APIFieldName="ActiveSyncDeviceFirstSyncTime";},
		@{DBFieldName="LastPolicyUpdate"; DBFieldType="datetime"; S3APIFieldName="ActiveSyncDeviceLastPolicyUpdate";},
		@{DBFieldName="LastSyncAttempt"; DBFieldType="datetime"; S3APIFieldName="ActiveSyncDeviceLastSyncAttempt";},
		@{DBFieldName="LastSuccessSync"; DBFieldType="datetime"; S3APIFieldName="ActiveSyncDeviceLastSuccessfulSync";},
		@{DBFieldName="DeviceWipeRequestTime"; DBFieldType="datetime"; S3APIFieldName="ActiveSyncDeviceWipeRequestTime";},
		@{DBFieldName="DeviceWipeSentTime"; DBFieldType="datetime"; S3APIFieldName="ActiveSyncDeviceWipeSentTime";},
		@{DBFieldName="DeviceWipeAckTime"; DBFieldType="datetime"; S3APIFieldName="ActiveSyncDeviceWipeAckTime";},
		@{DBFieldName="DevicePolicyApplicationStatus"; DBFieldType="nvarchar"; DBFieldWidth=64; S3APIFieldName="ActiveSyncDevicePolicyApplicationStatus";}
	)
}

$MappingServices = @{
	EntityName = "Service";
	Table = "tblServices";
	Sam360ReportId = "UniversalInventoryUniversalInventoryServicesTable";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="Name"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="WindowsServiceServiceName";},
		@{DBFieldName="DisplayName"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="WindowsServiceServiceDisplayName";},
		@{DBFieldName="Description"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="WindowsServiceServiceDescription";},
		@{DBFieldName="InstallLocation"; DBFieldType="nvarchar"; DBFieldWidth=1024; S3APIFieldName="WindowsServiceExecutablePathName";},
		@{DBFieldName="Account"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="WindowsServiceStartName";},
		@{DBFieldName="Started"; DBFieldType="bit"; S3APIFieldName="WindowsServiceServiceStarted";},
		@{DBFieldName="StartMode"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="WindowsServiceServiceStartMode";},
		@{DBFieldName="Status"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="WindowsServiceServiceStatus";}
	)
}

$MappingWebBroswers = @{
	EntityName = "Web Browser";
	Table = "tblWebBrowsers";
	Sam360ReportId = "UniversalInventoryUniversalInventoryWebBrowsersTable";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="BrowserName"; DBFieldType="nvarchar"; DBFieldWidth=64; S3APIFieldName="ProductName";}
	)
}

$MappingGroups = @{
	EntityName = "Group";
	Table = "tblGroups";
	Sam360ReportId = "UniversalInventoryUniversalInventoryGroupsTable";
	FieldMappings = @(
		@{DBFieldName="SrcId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="GroupAnonymisedName";},
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="Domain"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="DeviceDomain";},
		@{DBFieldName="Name"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="GroupName";},
		@{DBFieldName="Description"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="GroupDescription";},
		@{DBFieldName="GroupType"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="GroupADGroupType"; ConverterFn="GetADGroupTypeName"},
		@{DBFieldName="IsSecurity"; DBFieldType="nvarchar"; DBFieldWidth=256; S3APIFieldName="GroupADGroupType"; ConverterFn="GetIsSecurityGroup"}
	)
}

$MappingVMEvents = @{
	EntityName = "VM Events";
	Table = "tblVmEvents";
	Sam360ReportId = "UniversalInventoryUniversalInventoryVmEventsTable";
	FieldMappings = @(
		@{DBFieldName="SrcVmId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="EventType"; DBFieldType="nvarchar"; DBFieldWidth=32; S3APIFieldName="GroupName"; ConverterFn="GetVMMoveReason"},
		@{DBFieldName="EventTime"; DBFieldType="datetime"; S3APIFieldName="GroupDescription";}
	)
}

$MappingDeviceLogons = @{
	EntityName = "Device Logons";
	Table = "tblDeviceLogons";
	Sam360ReportId = "UniversalInventoryUniversalInventoryDeviceLogonsTable";
	FieldMappings = @(
		@{DBFieldName="SrcDeviceId"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="DeviceAnonymisedDeviceName";},
		@{DBFieldName="UserDomain"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="UserWindowsUserDomain";},
		@{DBFieldName="UserName"; DBFieldType="nvarchar"; DBFieldWidth=128; S3APIFieldName="UserWindowsUserName";},
		@{DBFieldName="LogonCount"; DBFieldType="int"; S3APIFieldName="SoftwareUsageRecordActivityDate";},
		@{DBFieldName="TotalLogonMinutes"; DBFieldType="int"; S3APIFieldName="SoftwareUsageRecordDurationMins";}
	)
}

function GetSam360Inventory() {

    InitialiseLogFile
    LogEnvironmentDetails
    SetupDateFormats
	
	$DBConnection = GetDBConnectionObject
	if (!$DBConnection) {
		return
	}

    if (!($UserName -and $Password)) {
        $Creds = Get-ConsoleCredential -Message "Sam360 Credentials Required" -DefaultUsername $UserName
        if ($Creds) {
            $UserName = $Creds.UserName
            $Password = $Creds.Password
        }
        if (!($UserName -and $Password)) {
            LogError "User Name and Password are required to authenticate to the Sam360 API server"
            return
        }
    }

    LogProgress -progressDescription "Authenticating to $APIServer"
    $token = GetAPIToken
    if (!$token) {
        return
    }

    if ($ClientOrganisationId -eq "0") {
        # No OrganisationId has been specified - Get the list of available organsations
        LogProgress -progressDescription "Retrieving organisation details"
        $clientOrganisations = RunReport -Token $token -ReportId "OrganisationsAllActiveClientOrganisationsGeneralDetails"
        if ($clientOrganisations -and $clientOrganisations.Count -gt 1) {
            # This user has access to more than 1 organisation - Ask which organsation to export data for
            $global:orgCounter = 0
            $clientOrganisations | % { new-object PSObject -Property $_} | Format-Table @{name="Index";expression={$global:orgCounter;$global:orgCounter+=1}}, OrganisationName, OrganisationDeviceCount
            $orgIndex = QueryUser -Message "$($clientOrganisations.Count) organisations available. Select the required organisation" -Prompt "Index" -DefaultValue "1"
            $orgIndexInt = [int]$orgIndex - 1
            $ClientOrganisationId = $clientOrganisations[$orgIndexInt].OrganisationID
            LogText -Color Green "Organisation '$($clientOrganisations[$orgIndexInt].OrganisationName)' selected"
        }
    }
    
    # Copy Sam360 Data to local database
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingDevices
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingUsers
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingSoftwareInstallations
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingCPUs
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingDatabaseServers
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingVMs
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingVMHosts
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingMailServers
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingActiveSyncDevices
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingServices
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingWebBroswers
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingGroups
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingVMEvents
	TransferSam360Data -Token $token -OrganisationId $ClientOrganisationId -dbConnection $DBConnection -dataMapping $MappingDeviceLogons
	
	ExecuteImportSP -dbConnection $DBConnection
}

function GetAPIToken {

    $creds = @{
        username = $UserName
        password = $Password
        grant_type = "password"
    }

    try {
        $response = Invoke-RestMethod "$APIServer/token" -Method Post -Body $creds
        $token = $response.access_token
        return $token;
    }
    catch {
        LogException -Exception $_.Exception
    }

    return
}

function RunReport($Token, [string]$ReportId = "HardwareAllWindowsDomainsGeneralDetails", [string]$OrganisationId = "0") {
    
    $apiHeaders = @{}
    $apiHeaders.Add("Authorization", "Bearer $Token")

    try {
        $apiResult = Invoke-RestMethod "$APIServer/api/Report/GetRecords?reportId=$ReportId&organisationId=$OrganisationId" -Method GET -Headers $apiHeaders -ContentType "application/json"
        if (!$apiResult) {
            return
        }

        # $result is a 2d String array
        $records = New-Object System.Collections.Generic.List[System.Collections.Hashtable]
        $fieldNames = $apiResult[0]
		#LogProgress ($fieldNames -join ",")
        for ($i1=1; $i1 -lt $apiResult.length; $i1++) {
            $record = @{}
            for ($i2=0; $i2 -lt $apiResult[$i1].length; $i2++) {
                $record[$fieldNames[$i2].Replace("-","")] = $apiResult[$i1][$i2]
            }
            $records.Add($record)
        }

        return $records
    }
    catch {
        LogException -Exception $_.Exception
    }
}

function GetDBConnectionObject() {
	$Connection = $null

	if (!($UIDatabaseServerName)) {
        $UIDatabaseServerName = QueryUser -Message "Universal Inventory Details Required" -Prompt "Server Name" -DefaultValue "localhost\SQLEXPRESS"
		if (!$UIDatabaseServerName){
			return $null
		}
	}

	if (!($UIDatabaseName)) {
        $UIDatabaseName = QueryUser -Prompt "Database Name" -DefaultValue "UI_DB1"
		if (!$UIDatabaseName){
			return $null
		}
	}

	LogText
	
	try {
		$ConnectionString = "Server=$UIDatabaseServerName; Database=$UIDatabaseName; Integrated Security=True; Persist Security Info=False";
		$Connection = New-Object System.Data.SqlClient.SqlConnection; 
	    $Connection.ConnectionString = $ConnectionString; 
		$Connection.Open();
	}
    catch {
        LogException -Exception $_.Exception
    }
	
	return $Connection;
}

function TransferSam360Data($Token, [string]$OrganisationId = "0", [System.Data.SqlClient.SqlConnection] $dbConnection, $dataMapping) {
	# Query Sam360 API
	LogProgress -progressDescription "Querying $($dataMapping.EntityName) details from Sam360"
	$reportData = RunReport -Token $token -OrganisationId $ClientOrganisationId -ReportId $dataMapping.Sam360ReportId
	#$reportData | %{New-Object -Type PSObject -Property $_} | Export-Csv  "c:\\temp\\s3api\\$($dataMapping.Sam360ReportId).csv" -NoTypeInformation -Encoding UTF8
	
	# PreProcess data before saving to DB
	if ($dataMapping.PreProcessFn) {
		$PreProcessFn = $dataMapping.PreProcessFn
		$reportData = & $PreProcessFn -reportData $reportData;
	}
	
	# Save data to local database
    LogProgress -progressDescription "Saving $($dataMapping.EntityName) details to database [$($reportData.Count) Records]"
	SaveData -data $reportData -dbConnection $DBConnection -dataMapping $dataMapping
}

function SaveData([System.Data.SqlClient.SqlConnection] $dbConnection, $dataMapping, $data) {
	# Delete all existing Sam360 data from database
	$sqlQueryDelete = "DELETE FROM [in].[$($dataMapping.Table)] WHERE DataSourceId = 'Sam360'"
	$sqlCommandDelete = New-Object System.Data.SqlClient.SqlCommand($sqlQueryDelete, $dbConnection);
	$result = $sqlCommandDelete.ExecuteScalar()
	
	# Create Insert Command
	$sqlQueryInsert = "INSERT INTO [in].[$($dataMapping.Table)] (DataSourceId"
	$sqlCommandInsert = New-Object System.Data.SqlClient.SqlCommand("", $dbConnection);
	$sqlParamDataSourceId = $sqlCommandInsert.Parameters.Add("ParamDataSourceId", [system.data.SqlDbType]::NVarChar);
	$sqlParamDataSourceId.Value = "Sam360"
	$lstParams = New-Object System.Collections.Generic.List[System.Data.SqlClient.SqlParameter]
	
	$sqlParamCounter = 0;
	foreach ($fieldMapping in $dataMapping.FieldMappings) {
		$sqlParamCounter++;
		
		$fieldType = GetFieldType -fieldTypeName $fieldMapping.DBFieldType
		$sqlParam = $sqlCommandInsert.Parameters.Add("Param$sqlParamCounter", $fieldType);
		$lstParams.Add($sqlParam)
		
		$sqlQueryInsert += ", [$($fieldMapping.DBFieldName)]"
	}
	
	$sqlQueryInsert += ") VALUES (@ParamDataSourceId"
	
	for ($nCounter = 1; $nCounter -le $sqlParamCounter; $nCounter++) {
		$sqlQueryInsert += ", @Param$nCounter"
	}
	
	$sqlQueryInsert += ")"
	$sqlCommandInsert.CommandText = $sqlQueryInsert
	
	# Save each record
	foreach ($record in $data) {
		$sqlParamCounter2 = 0;
		foreach ($fieldMapping in $dataMapping.FieldMappings) {
			$GetFieldValueFn = $fieldMapping.ConverterFn
			if (!$GetFieldValueFn) {
				$GetFieldValueFn = "GetFieldValue"
			}
			$lstParams[$sqlParamCounter2].Value = & $GetFieldValueFn -fieldMapping $fieldMapping -record $record;
			$sqlParamCounter2++;
		}
		$result = $sqlCommandInsert.ExecuteScalar()
	}
}

function ExecuteImportSP([System.Data.SqlClient.SqlConnection] $dbConnection) {
	$strCommandText = "[in].sp_Import"
	$sqlCommand = New-Object System.Data.SqlClient.SqlCommand($strCommandText, $dbConnection);

	$sqlCommand.CommandType = [System.Data.CommandType]'StoredProcedure';
	$outParameter = new-object System.Data.SqlClient.SqlParameter;
	$outParameter.ParameterName = "@lastErrorMsg";
	$outParameter.Direction = [System.Data.ParameterDirection]'Output';
	$outParameter.DbType = [System.Data.DbType]'String';
	$outParameter.Size = 2500;
	$sqlCommand.Parameters.Add($outParameter) >> $null;
	$result = $sqlCommand.ExecuteNonQuery();
	$truth = $sqlCommand.Parameters["@lastErrorMsg"].Value;
	$result
	$truth


	#$result = $sqlCommand.ExecuteScalar()
}

function GetFieldType($fieldTypeName) {
	switch ($fieldTypeName) {
		("varchar") {return [system.data.SqlDbType]::VarChar;}
		("nvarchar") {return [system.data.SqlDbType]::NVarChar;}
		("datetime") {return [system.data.SqlDbType]::DateTime;}
		("bigint") {return [system.data.SqlDbType]::BigInt;}
		("bit") {return [system.data.SqlDbType]::Bit;}
		("int") {return [system.data.SqlDbType]::Int;}
		("tinyint") {return [system.data.SqlDbType]::TinyInt;}
	}
	
	throw "Unknown Data Type: $fieldTypeName"
}

function GetFieldValue($fieldMapping, $record) {
	try {
		if ($fieldMapping.DBFieldType -eq "varchar" -or $fieldMapping.DBFieldType -eq "nvarchar") {
			if ($record[$fieldMapping.S3APIFieldName]){
				if (($record[$fieldMapping.S3APIFieldName]).Length -gt $fieldMapping.DBFieldWidth) {
					return  $record[$fieldMapping.S3APIFieldName].Substring(0,$fieldMapping.DBFieldWidth - 1)
				}
				else {
					return $record[$fieldMapping.S3APIFieldName]
				}
			}
			else {
				return ""
			}
		}
		elseif ($fieldMapping.DBFieldType -eq "bigint") {
			if ($record[$fieldMapping.S3APIFieldName]){
				return [double]::Parse($record[$fieldMapping.S3APIFieldName]) * 1000
			}
			else {
				return [System.DBNull]::Value
			}
		}
		elseif ($fieldMapping.DBFieldType -eq "datetime") {
			if ($record[$fieldMapping.S3APIFieldName]){
				return [DateTime]::Parse($record[$fieldMapping.S3APIFieldName])
			}
			else {
				return [System.DBNull]::Value
			}
		}
		else {
			if ($record[$fieldMapping.S3APIFieldName]){
				return $record[$fieldMapping.S3APIFieldName]
			}
			else {
				return [System.DBNull]::Value
			}
		}
	}
	catch {
		return [System.DBNull]::Value
	}
}

function GetInverseBoolValue ($fieldMapping, $record) {
	if ($record[$fieldMapping.S3APIFieldName] -like "TRUE") {
		return "FALSE";
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "FALSE") {
		return "TRUE";
	}
	else {
		return GetFieldValue -fieldMapping $fieldMapping -record $record
	}
}

function GetPasswordRequired ($fieldMapping, $record) {
	if (!$record[$fieldMapping.S3APIFieldName]) {
		return [System.DBNull]::Value
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "*PASSWD_NOTREQD*") {
		return "FALSE";
	}
	
	return "TRUE";
}

function GetMsdmTableExists ($fieldMapping, $record) {
	if (!$record["DeviceLastUpdate"]) {
		return [System.DBNull]::Value
	}
	elseif ($record[$fieldMapping.S3APIFieldName]) {
		return 1;
	}
	
	return 0;
}

function GetPasswordCanChange ($fieldMapping, $record) {
	if (!$record[$fieldMapping.S3APIFieldName]) {
		return [System.DBNull]::Value
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "*PASSWD_CANT_CHANGE*") {
		return "FALSE";
	}
	
	return "TRUE";
}

function GetPasswordExpires ($fieldMapping, $record) {
	if (!$record[$fieldMapping.S3APIFieldName]) {
		return [System.DBNull]::Value
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "*DONT_EXPIRE_PASSWORD*") {
		return "FALSE";
	}
	
	return "TRUE";
}

function GetIsNormalAccount ($fieldMapping, $record) {
	if (!$record[$fieldMapping.S3APIFieldName]) {
		return [System.DBNull]::Value
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "*NORMAL_ACCOUNT*") {
		return "TRUE";
	}
	
	return "FALSE";
}

function GetSQLServiceLoginMode ($fieldMapping, $record) {
	if ($record[$fieldMapping.S3APIFieldName] -like "Integrated") {
		return "TRUE";
	}
	elseif ($record[$fieldMapping.S3APIFieldName]) {
		return "FALSE";
	}
	else {
		return [System.DBNull]::Value
	}
}

function GetSQLServiceType ($fieldMapping, $record) {
	if ($record[$fieldMapping.S3APIFieldName] -like "SQL Server") {
		return 1;
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "Report Server") {
		return 6;
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "Analysis Server") {
		return 5;
	}
	else {
		return [System.DBNull]::Value
	}
}

function GetMBValueFromGB ($fieldMapping, $record) {
	if ($record[$fieldMapping.S3APIFieldName]){
		return [double]::Parse($record[$fieldMapping.S3APIFieldName]) * 1000
	}
	else {
		return [System.DBNull]::Value
	}
}

function GetVMMoveReason ($fieldMapping, $record) {
	return "Migrated"
}

function GetADGroupTypeName ($fieldMapping, $record) {
	if ($record[$fieldMapping.S3APIFieldName] -like "*Domain*") {
		return "Domain-local";
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "*Global*") {
		return "Global";
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "*Universal*") {
		return "Universal";
	}
	else {
		return [System.DBNull]::Value
	}
}

function GetIsSecurityGroup ($fieldMapping, $record) {
	if (!$record[$fieldMapping.S3APIFieldName]) {
		return [System.DBNull]::Value;
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -like "*Security*") {
		return "TRUE";
	}
	else {
		return "FALSE";
	}
}

function GetDBServerSrcID ($fieldMapping, $record) {
	if (!$record[$fieldMapping.S3APIFieldName]) {
		return [System.DBNull]::Value;
	}
	elseif (!$record["SQLServiceSam360InstanceID"]) {
		return $record[$fieldMapping.S3APIFieldName];
	}
	else {
		return $record[$fieldMapping.S3APIFieldName] + "-SQL" + $record["SQLServiceSam360InstanceID"];
	}
}

function GetBoolFromYesNo ($fieldMapping, $record) {
	if ($record[$fieldMapping.S3APIFieldName] -eq "Yes") {
		return "TRUE";
	}
	elseif ($record[$fieldMapping.S3APIFieldName] -eq "No"){
		return "FALSE";
	}
	else {
		return [System.DBNull]::Value;
	}
}

function PreProcessCPUs($reportData) {
	# Processor Info includes one record per device. We need to copy records where
	# a device has more than one processor
	$updateReportData = New-Object System.Collections.Generic.List[System.Collections.Hashtable]
	foreach ($record in $reportData) {
		$cpuIDs = $record["DeviceProcessorDeviceID"] -split ","
		$cpuCount = [int]$record["DeviceProcessors"]
		for ($n1=0; $n1 -lt $cpuCount; $n1++) {
			$newRecord = $record.Clone()
			if ($n1 -lt $cpuIDs.Length) {
				$newRecord["DeviceProcessorDeviceID"] = $cpuIDs[$n1]
			}
			$updateReportData.Add($newRecord)
		}
	}
	
	return $updateReportData;
}

function LogException($Exception) {
    $errorDescription = ""
    if ($Exception.Response) {
        $reader = New-Object System.IO.StreamReader($Exception.Response.GetResponseStream())
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd() | ConvertFrom-Json
        $errorDescription = $responseBody.error
    }
    else {
        $errorDescription = $Exception.Message
    }

    LogText -Color Red "Sam360 API Error: $errorDescription"
}

GetSam360Inventory