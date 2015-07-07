 ##########################################################################
 #
 # Get-ExchangeDetails
 # SAM Gold Toolkit
 # Original Source: Jon Mulligan (Sam360)
 #	              : Sanjay Ramaswamy https://gallery.technet.microsoft.com/scriptcenter/acdcb192-f226-4517-b3f9-005dce6f4fc3
 #                : Oliver Moazzezi http://www.exchange2010.com/2013/11/calculating-cal-requirements-for.html
 #
 ##########################################################################

 Param(
	[alias("server")]
	$ExchangeServer = $env:computerName,
	[alias("o1")]
	$OutputFile1 = "ExchangeServerDetails" + $ExchangeServer + ".csv",
	[alias("o2")]
	$OutputFile2 = "ExchangeMailBoxes" + $ExchangeServer + ".csv",
	[alias("o3")]
	$OutputFile3 = "ExchangeDevices" + $ExchangeServer + ".csv",
	[alias("o4")]
	$OutputFile4 = "ExchangeCALs" + $ExchangeServer + ".csv",
	[alias("o5")]
	$OutputFile5 = "ExchangeCALDetails" + $ExchangeServer + ".csv",
	$UserName,
	$Password,
	[switch]
	$Office365,
	[switch]
	$Verbose,
	[switch]
	$UseSSL,
	[ValidateSet("2007","2010","2010SP1","2010SP3","2013")]
	$CALScriptVersion,
	[ValidateSet("AllData","ServerData","EntityData","UtilizationData","CALData")] 
	$RequiredData = "AllData",
	[ValidateSet("Both","RemoteSession","SnapIn")] 
	$ConnectionMethod = "Both")

<#
.SYNOPSIS
Retrieves Exchange server, mail box, ActiveSync device and CAL information from an Exchange server

.DESCRIPTION
The Get-ExchangeDetails script queries a single Exchange server and produces up to 5 CSV files
	1)    ExchangeServerDetails.csv - One record per Exchange Server in the farm
    2)    ExchangeMailBoxes.csv - One record per MailBox
	3)    ExchangeDevices.csv - One record per ActiveSync device
	4)    ExchangeCALs.csv - General CAL requirement details 
    5)    ExchangeCALDetails.csv - Lists all servers and MailBoxes that require a license and 
	      the type of license required

    Files are written to current working directory

.PARAMETER Server 
Host name of Exchange server to scan

.PARAMETER Office365
Flag. Query Office365 hosted Exchange environment. If this flag is set, the parameter 'Server' is ignored

.PARAMETER Username
Exchange Server Username

.PARAMETER Password
Exchange Server Password

.PARAMETER Verbose
Display extra progress information on screen

.PARAMETER CALScriptVersion
This script contains multiple embedded scripts in order to determine required CAL count for the server.
The script attempts to pick the correct embedded script based on the edition of the selected Exchange 
server. However, a different embedded script can be selected manually. Allowed options are..
"2007","2010","2010SP1","2010SP3","2013"

.PARAMETER RequiredData
By default the script collects Exchange server, mail box, ActiveSync device and CAL information from the 
selected Exchange server. It's possible to collect subsets of the data. Possible options are...
"AllData","ServerData","EntityData","UtilizationData","CALData"

.EXAMPLE
Get all guest, host and migration info from Hyper-V server 'Defiant'. 
Get-HyperVVMList –HyperVServer Defiant

.NOTES
This script supports Exchange server 2007 onwards. There are some limitations on what data can be
collected from different versions of Exchange server remotely. If the script fails to execute remotely
try
	1)	Specify a username and password (even if they are the details of the current user)
	2)  Execute the script locally on the Exchange Server
#>
	
function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Output "Computer Name:             $($env:COMPUTERNAME)"
	Write-Output "User Name:                 $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:           $($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:           $($host.Version.Major)"
	Write-Output "PowerShell Version:        $($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size:      $($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:               $($PSVersionTable.CLRVersion)"
	Write-Output "Username Parameter:        $UserName"
	Write-Output "Server Parameter:          $Server"
	Write-Output "Required Data:             $RequiredData"
	Write-Output "Connection Method:         $ConnectionMethod"
	Write-Output "CAL Script Version:        $CALScriptVersion"
}

function LogLastException() {
    $currentException = $Error[0].Exception;

    while ($currentException)
    {
        write-output $currentException
        write-output $currentException.Data
        write-output $currentException.HelpLink
        write-output $currentException.HResult
        write-output $currentException.Message
        write-output $currentException.Source
        write-output $currentException.StackTrace
        write-output $currentException.TargetSite

        $currentException = $currentException.InnerException
    }
}

function LogProgress([string]$Activity, [string]$Status, [Int32]$PercentComplete, [switch]$Completed ){
	
	Write-Progress -activity $Activity -Status $Status -percentComplete $PercentComplete -Completed:$Completed
	
	if ($Verbose)
	{
		write-output ""
		$output = Get-Date -Format HH:mm:ss.ff
		$output += " - "
		$output += $Status
		write-output $output
	}
}

function EnvironmentConfigured {
	if (Get-Command "Get-MailboxStatistics" -errorAction SilentlyContinue){
		return $true}
	else {
		return $false}
}

function Get-ExchangeDetails {

	LogProgress -Activity "Exchange Data Export" -Status "Logging environment details" -percentComplete 1
	LogEnvironmentDetails

	# Create the Credentials object if username has been provided
	if ($UserName)
	{
		LogProgress -activity "Exchange Data Export" -Status "Creating Credentials Object" -percentComplete 2

		if ($Password) {
			$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
		}
		else {
			$securePassword = Read-Host 'Password' -AsSecureString
		}
		
		$exchangeCreds = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)
	}

	# Connect to exchange server
	LogProgress -activity "Exchange Data Export" -Status "Connecting..." -percentComplete 3
	if ($Office365)
	{
		$connectionUri = "https://ps.outlook.com/powershell/"
		$authenticationType = "Basic"
	}
	else
	{
		$connectionUri = "http://"
		if ($UseSSL) {
			$connectionUri = "https://"
		}
		$connectionUri += $ExchangeServer + "/powershell/"
		$authenticationType = "Kerberos"
	}
	
	if ($Verbose)
	{
		Write-Output "ConnectionUri: $connectionUri"
		Write-Output "AuthenticationType: $authenticationType"
		Write-Output "UserName: $($exchangeCreds.UserName)"
	}
	
	if ($ConnectionMethod -eq "Both" -or $ConnectionMethod -eq "RemoteSession")
	{
		if ($exchangeCreds)
		{
			$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Authentication $authenticationType -AllowRedirection -Credential $exchangeCreds -WarningAction:silentlycontinue
		}
		else
		{
			$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Authentication $authenticationType -AllowRedirection -WarningAction:silentlycontinue
		}
	}
		
	if ($exchangeSession) {
		LogProgress -activity "Exchange Data Export" -Status "Importing Session" -percentComplete 10	
		Import-PSSession $exchangeSession -AllowClobber -WarningAction:silentlycontinue
	}

	if (!(EnvironmentConfigured) -and !($Office365))
	{
		# Exchange environment not configured
		if ($ConnectionMethod -eq "Both" -or $ConnectionMethod -eq "SnapIn")
		{
			# Load Exchange SnapIns
			LogProgress -activity "Exchange Data Export" -Status "Loading SnapIns" -percentComplete 11
			
			$allSnapIns = get-pssnapin -registered
			if ($Verbose)
			{
				Write-Output "SnapIns"
				$allSnapIns | % { Write-Output "Name: $($_.Name) Version: $($_.PSVersion)"}
			}
			
			$allSnapIns = $allSnapIns | sort -Descending
			
			foreach ($snapIn in $allSnapIns){
				if (($snapIn.name -eq 'Microsoft.Exchange.Management.PowerShell.Admin') -or
					($snapIn.name -eq 'Microsoft.Exchange.Management.PowerShell.E2010') -or
					($snapIn.name -eq 'Microsoft.Exchange.Management.PowerShell.E2013')){
					write-output "Adding SnapIn: $($snapIn.Name)"
					add-PSSnapin -Name $snapIn.name
					
					if (EnvironmentConfigured) {
						break}
				}
			}
		}
	}
	
	if (!(EnvironmentConfigured))
	{
		write-output "Unable to configure Powershell Exchange environment"
		exit
	}   

	# Get the list of Exchange Servers (Not supported in Office365)
	if (!($Office365))
	{
		LogProgress -activity "Exchange Data Export" -Status "Getting server details" -percentComplete 15
		if (Get-Command "Get-ExchangeServer" -errorAction SilentlyContinue){
			$exchangeServers = Get-ExchangeServer -Identity $ExchangeServer
			$exchangeServers | export-csv $OutputFile1 -notypeinformation -Encoding UTF8
			if ($Verbose) {
				Write-Output "Server Count: $($exchangeServers.Count)"
			}
		}
		else {
			Write-Output "Exchange cmdlet Get-ExchangeServer not found. Does current user have sufficient permissions?" 
		}
	}

	if ($RequiredData -eq "EntityData" -or $RequiredData -eq "UtilizationData" -or $RequiredData -eq "AllData")
	{
		#Get the list of mailboxes from Exchange
		LogProgress -activity "Exchange Data Export" -Status "Querying Mailboxes" -percentComplete 20
		$mailBoxes = Get-Mailbox -ResultSize 'Unlimited'
		if ($mailBoxes) 
		{
			if ($Verbose) {
				Write-Output "Mailbox Count: $($mailBoxes.Count)"
				Write-Output  ([string]::Format("{0,-5} {1,-55} {2,-20}","Count","UserPrincipalName","LastLogonTime"))
			}

			$listMailBoxData = New-Object System.Collections.Generic.List[System.Management.Automation.PSObject]
			$countMailBoxes = 1

			foreach ($mailBox in $mailBoxes) {
				
				$mailBoxData = $mailBox | select -Property UserPrincipalName, SamAccountName, DisplayName, 
					WindowsLiveID, ExchangeGuid, PrimarySmtpAddress, ExternalDirectoryObjectId, EmailAddresses, 
					DistinguishedName, Guid, RecipientType, IsMailboxEnabled, WhenMailboxCreated, WhenCreatedUTC, 
					WhenChangedUTC, LastLogonTime, LastLogoffTime

				if ($RequiredData -eq "UtilizationData" -or $RequiredData -eq "AllData")
				{
					Write-Progress -activity "Exchange Data Export" -Status "Querying Mailbox Stats $($mailBoxData.UserPrincipalName)" -percentComplete (20 + ($countMailBoxes/$mailBoxes.Count)*40)
					$mailBoxStatistics = $mailBox | Get-MailboxStatistics -WarningAction:silentlycontinue
					$mailBoxData.LastLogonTime = $mailBoxStatistics.LastLogonTime
					$mailBoxData.LastLogoffTime = $mailBoxStatistics.LastLogoffTime
				}

				if ($Verbose) {
					Write-Output  ([string]::Format("{0,-5} {1,-55} {2,-20}", $countMailBoxes, $mailBoxData.UserPrincipalName, $mailBoxData.LastLogonTime))
				}

				$listMailBoxData.Add($mailBoxData)
				$countMailBoxes++
			}

			$listMailBoxData | export-csv $OutputFile2 -notypeinformation -Encoding UTF8
		}
		
		#Get Device details
		LogProgress -activity "Exchange Data Export" -Status "Querying Device Data" -percentComplete 60
		$activeSyncDevices = Get-ActiveSyncDevice -ResultSize 'Unlimited' -WarningAction:silentlycontinue 
		if ($activeSyncDevices)
		{
			if ($Verbose) {
				Write-Output "Device Count: $($activeSyncDevices.Count)"
				Write-Output  ([string]::Format("{0,-5} {1,-19} {2,-35} {3,-20}","Count","User","DeviceOS","LastSuccessSync"))
			}

			$listActiveSyncDeviceData = New-Object System.Collections.Generic.List[System.Management.Automation.PSObject]
			$countDevices = 1

			foreach ($activeSyncDevice in $activeSyncDevices) {
				
				$activeSyncDeviceData = $activeSyncDevice | select -Property Identity, FriendlyName, Name, 
					DeviceId, Guid, DeviceImei, DeviceTelephoneNumber, DeviceMobileOperator, DeviceOS, DeviceOSLanguage, 
					DeviceType, DeviceUserAgent, DeviceModel, UserDisplayName, OrganizationId, DeviceActiveSyncVersion, 
					FirstSyncTime, WhenCreatedUTC, WhenChangedUTC, LastPingHeartbeat, LastSyncAttemptTime, LastSuccessSync, 
					LastPolicyUpdateTime, DevicePolicyApplied, DevicePolicyApplicationStatus, Status, StatusNote, 
					IsRemoteWipeSupported, DeviceWipeSentTime, DeviceWipeRequestTime, DeviceWipeAckTime

				if ($RequiredData -eq "UtilizationData" -or $RequiredData -eq "AllData"){
					Write-Progress -activity "Exchange Data Export" -Status "Querying Device Stats $($activeSyncDeviceData.FriendlyName)" -percentComplete (60 + ($countDevices/$activeSyncDevices.Count)*38)
					$activeSyncDeviceStatistics = $activeSyncDevice | Get-ActiveSyncDeviceStatistics -WarningAction:silentlycontinue
					if ($activeSyncDeviceStatistics){
						$activeSyncDeviceData.FirstSyncTime = $activeSyncDeviceStatistics.FirstSyncTime
						$activeSyncDeviceData.LastPingHeartbeat = $activeSyncDeviceStatistics.LastPingHeartbeat
						$activeSyncDeviceData.LastSyncAttemptTime = $activeSyncDeviceStatistics.LastSyncAttemptTime
						$activeSyncDeviceData.LastSuccessSync = $activeSyncDeviceStatistics.LastSuccessSync
						$activeSyncDeviceData.LastPolicyUpdateTime = $activeSyncDeviceStatistics.LastPolicyUpdateTime
						$activeSyncDeviceData.DevicePolicyApplied = $activeSyncDeviceStatistics.DevicePolicyApplied
						$activeSyncDeviceData.DevicePolicyApplicationStatus = $activeSyncDeviceStatistics.DevicePolicyApplicationStatus
						$activeSyncDeviceData.Status = $activeSyncDeviceStatistics.Status
						$activeSyncDeviceData.StatusNote = $activeSyncDeviceStatistics.StatusNote
						$activeSyncDeviceData.IsRemoteWipeSupported = $activeSyncDeviceStatistics.IsRemoteWipeSupported
						$activeSyncDeviceData.DeviceWipeSentTime = $activeSyncDeviceStatistics.DeviceWipeSentTime
						$activeSyncDeviceData.DeviceWipeRequestTime = $activeSyncDeviceStatistics.DeviceWipeRequestTime
						$activeSyncDeviceData.DeviceWipeAckTime = $activeSyncDeviceStatistics.DeviceWipeAckTime
					}
				}

				if ($Verbose) {
					$activeDeviceUser = GetUserNameFromDeviceID -DeviceID $activeSyncDeviceData.Identity
					Write-Output  ([string]::Format("{0,-5} {1,-19} {2,-35} {3,-20}", $countDevices, $activeDeviceUser,
						$activeSyncDeviceData.DeviceOS, $activeSyncDeviceData.LastSuccessSync))
				}

				$listActiveSyncDeviceData.Add($activeSyncDeviceData)
				$countDevices++
			}

			$listActiveSyncDeviceData | export-csv $OutputFile3 -notypeinformation -Encoding UTF8
		}
	}
	
	if (($RequiredData -eq "AllData" -or $RequiredData -eq "CALData") -and (!($Office365)))
	{
		if (!$CALScriptVersion) {
			if (Get-Command "Get-ExchangeServerAccessLicenseUser" -errorAction SilentlyContinue){
				$CALScriptVersion = "2013"
			}
			elseif ($exchangeServers) {
				$Version = $exchangeServers.AdminDisplayVersion
				if ($Version -Like "Version 8*") {
					$CALScriptVersion = "2007"
				}
				elseif ($Version -Like "Version 14.1*" -or
						$Version -Like "Version 14.2*" -or
						$Version -Like "Version 14.3*" ) {
					$CALScriptVersion = "2010SP1"
				}
				elseif ($Version -Like "Version 14*") {
					$CALScriptVersion = "2010"
				}
				else {
					$CALScriptVersion = "2013"
				}
			}
			else {
				$CALScriptVersion = "2010"
			}
		}
		
		if ($Verbose) {
			Write-Output "Running CAL script: Version $($CALScriptVersion)"
		}
		
		if ($CALScriptVersion -eq "2007") {
			& $scriptGetCALReqs2007
		}
		elseif ($CALScriptVersion -eq "2010") {
			& $scriptGetCALReqs2010
		}
		elseif ($CALScriptVersion -eq "2010SP1") {
			& $scriptGetCALReqs2010SP1
		}
		elseif ($CALScriptVersion -eq "2010SP3") {
			& $scriptGetCALReqs2010SP3
		}
		else {
			& $scriptGetCALReqs2013
		}
	}

	if ($exchangeSession) {
		LogProgress -activity "Exchange Data Export" -Status "Cleaning Session" -percentComplete 98
		Remove-PSSession -Session $exchangeSession}
		
	LogProgress -activity "Exchange Data Export" -complete -Status "Complete"
}

function GetUserNameFromDeviceID {
    param([string] $DeviceID = "")

	$deviceIDParts = $DeviceID.Split("/\")
	if ($deviceIDParts.length -eq 0){
		return ""
	}

	$indexEASD = [array]::IndexOf($deviceIDParts, "ExchangeActiveSyncDevices")
	if ($indexEASD -gt 0){
		$nameParts = $deviceIDParts[$indexEASD - 1].Split("@")
		if ($nameParts.Length -gt 0){
			return $nameParts[0]
		}
	}

	return $deviceIDParts[0]
}


# Function that outputs Exchange CALs in the organization 
function Output-Report { 
    Write-Output "=========================" 
    Write-Output "Exchange CAL Usage Report" 
    Write-Output "=========================" 
    Write-Output "" 
    Write-Output "Total Users:                                    $TotalMailboxes" 
    Write-Output "Total Standard CALs:                            $TotalStandardCALs" 
    Write-Output "Total Enterprise CALs:                          $TotalEnterpriseCALs" 
	
	$calReport = New-Object -TypeName System.Object
	$calReport | Add-Member -MemberType NoteProperty -Name TotalMailboxes -Value $TotalMailboxes
	$calReport | Add-Member -MemberType NoteProperty -Name TotalStandardCALs -Value $TotalStandardCALs
	$calReport | Add-Member -MemberType NoteProperty -Name TotalEnterpriseCALs -Value $TotalEnterpriseCALs
	$calReport | Add-Member -MemberType NoteProperty -Name UnifiedMessagingUserCount -Value $UMUserCount 
	$calReport | Add-Member -MemberType NoteProperty -Name ManagedCustomFolderUserCount -Value $ManagedCustomFolderUserCount 
	$calReport | Add-Member -MemberType NoteProperty -Name AdvancedActiveSyncPolicyUserCount -Value $AdvancedActiveSyncUserCount 
	$calReport | Add-Member -MemberType NoteProperty -Name ArchivedMailboxUserCount -Value $ArchiveUserCount 
	$calReport | Add-Member -MemberType NoteProperty -Name RetentionPolicyUserCount -Value $RetentionPolicyUserCount
	$calReport | Add-Member -MemberType NoteProperty -Name SearchableUserCount -Value $SearchableMaiboxIDs.Count
	$calReport | Add-Member -MemberType NoteProperty -Name JournalingUserCount -Value $JournalingUserCount
	$calReport | Add-Member -MemberType NoteProperty -Name InfoLeakageProtectionEnabled -Value $InfoLeakageProtectionEnabled
	$calReport | Add-Member -MemberType NoteProperty -Name AdvancedAntispamEnabled -Value $AdvancedAntispamEnabled
	$calReport | export-csv $OutputFile4 -notypeinformation -Encoding UTF8
} 

$scriptGetCALReqs2007 = 
{
# Trap block 
trap {  
    Write-Output "An error has occurred running the script:"  
    Write-Output $_ 
 
    $Global:AdminSessionADSettings.DefaultScope = $OriginalDefaultScope 
 	
	LogLastException
	
    exit 
}  
 
# Function that returns true if the incoming argument is a help request 
function IsHelpRequest 
{ 
    param($argument) 
    return ($argument -eq "-?" -or $argument -eq "-help"); 
 
} 
 
# Function that displays the help related to this script following 
# the same format provided by get-help or <cmdletcall> -? 
function Usage 
{ 
@" 
 
NAME: 
`tReportExchangeCALs.ps1 
 
SYNOPSIS: 
`tReports Exchange 2007 client access licenses (CALs) of this organization in Enterprise or Standard categories. 
 
SYNTAX: 
`tReportExchangeCALs.ps1 
 
PARAMETERS: 
 
USAGE: 
`t.\ReportExchangeCALs.ps1 
 
"@ 
} 
 
# Function that resets AdminSessionADSettings.DefaultScope to original value and exits the script 
function Exit-Script 
{ 
    $Global:AdminSessionADSettings.DefaultScope = $OriginalDefaultScope 
 
    exit 
} 
 
######################## 
## Script starts here ## 
######################## 
 
# Check for Usage Statement Request 
$args | foreach { if (IsHelpRequest $_) { Usage; Exit-Script; } } 
 
# Introduction message 
Write-Output "Report Exchange 2007 client access licenses (CALs) in use in the organization"  
Write-Output "It will take some time if there are a large amount of users......" 
Write-Output "" 
 
# Report all recipients in the org. 
$OriginalDefaultScope = $Global:AdminSessionADSettings.DefaultScope 
$Global:AdminSessionADSettings.DefaultScope = $Null 
 
$TotalMailboxes = 0 
$TotalEnterpriseCALs = 0 
$UMUserCount = 0 
$ManagedCustomFolderUserCount = 0 
$AdvancedActiveSyncUserCount = 0 
$AdvancedAntispamUserCount = 0 
$AdvancedAntispamEnabled = $False 
$OrgWideJournalingEnabled = $False 
$AllMailboxIDs = @{} 
$EnterpriseCALMailboxIDs = @{} 
$JournalingUserCount = 0 
$JournalingMailboxIDs = @{} 
$JournalingDGMailboxMemberIDs = @{} 
$TotalStandardCALs = 0 
$VisitedGroups = @{} 
$DGStack = new-object System.Collections.Stack 
 
# Bool variable for outputing progress information running this script. 
$EnableProgressOutput = $True 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Progress:" 
} 
 
################ 
## Debug code ## 
################ 
 
# Bool variable for output hash table information for debugging purpose. 
$EnableOutputCounts = $False 
 
function Output-Counts 
{ 
    if ($EnableOutputCounts -eq $False) { 
        return 
    } 
 
    Write-Output "Hash Table Name                                 Count" 
    Write-Output "---------------                                 -----" 
    Write-Output "AllMailboxIDs:                                 " $AllMailboxIDs.Count 
    Write-Output "EnterpriseCALMailboxIDs:                       " $EnterpriseCALMailboxIDs.Count 
    Write-Output "JournalingMailboxIDs:                          " $JournalingMailboxIDs.Count 
    Write-Output "JournalingDGMailboxMemberIDs:                  " $JournalingDGMailboxMemberIDs.Count 
    Write-Output "VisitedGroups:                                 " $VisitedGroups.Count 
    Write-Output "" 
    Write-Output "" 
} 
 
function Merge-Hashtables 
{ 
    $Table1 = $args[0] 
    $Table2 = $args[1] 
    $Result = @{} 
     
    if ($null -ne $Table1) 
    { 
        $Result += $Table1 
    } 
 
    if ($null -ne $Table2) 
    { 
        foreach ($entry in $Table2.GetEnumerator()) 
        { 
            $Result[$entry.Key] = $entry.Value 
        } 
    } 
 
    $Result 
} 
 

 
################# 
## Total Users ## 
################# 
 
# Note!!!  
# Only user, shared and linked mailboxes are counted.  
# Resource mailboxes and legacy mailboxes are NOT counted. 
 
Get-Mailbox -ResultSize 'Unlimited' -Filter { (RecipientTypeDetails -eq 'UserMailbox') -or 
                                              (RecipientTypeDetails -eq 'SharedMailbox') -or 
                                              (RecipientTypeDetails -eq 'LinkedMailbox') } | foreach { 
    $Mailbox = $_ 
     
 
    if ($Mailbox.ExchangeVersion.ToString().Contains(" (8.")) { 
        $AllMailboxIDs[$Mailbox.Identity] = $null 
        $Script:TotalMailboxes++ 
    } 
} 
 
if ($TotalMailboxes -eq 0) { 
    # No mailboxes in the org. Just output the report and exit 
    Output-Report 
     
    Exit-Script 
} 
 
######################### 
## Total Standard CALs ## 
######################### 
 
# All users are counted as Standard CALs 
$TotalStandardCALs = $TotalMailboxes 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Total Standard CALs calculated:                 $TotalStandardCALs" 
} 
 
############################# 
## Per-org Enterprise CALs ## 
############################# 
 
# If advanced anti-spam is turned on, all mailboxes are counted as Enterprise CALs 
Get-TransportServer | foreach { 
    # If advanced anti-spam is turned on any Hub/Edge server, all mailboxes in the org are counted as Exchange CALs 
     
    $AntispamUpdates = Get-AntispamUpdates $_ 
 
    if (($AntispamUpdates.SpamSignatureUpdatesEnabled -eq $True) -or 
        ($AntispamUpdates.IPReputationUpdatesEnabled -eq $True) -or 
        ($AntispamUpdates.UpdateMode -eq "Automatic")) { 
 
        $AdvancedAntispamEnabled = $True 
        $AdvancedAntispamUserCount = $TotalMailboxes     
        $TotalEnterpriseCALs = $TotalMailboxes 
 
        ## Progress output ...... 
        if ($EnableProgressOutput -eq $True) { 
            Write-Output "Advanced Anti-spam Enabled:                     True" 
            Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
            Write-Output "" 
        } 
 
        # All mailboxes are counted as Enterprise CALs, report and exit. 
        Output-Counts 
 
        Output-Report 
 
        Exit-Script 
    } 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Advanced Anti-spam Enabled:                     False" 
} 
 
 
############################## 
## Per-user Enterprise CALs ## 
############################## 
 
# 
# Calculate Enterprise CAL users using UM, MRM Managed Custom Folder, and advanced ActiveSync policy settings 
# 
$AllMailboxIDs.Keys | foreach {   
    $Mailbox = Get-Mailbox $_ 
     
     # UM usage classifies the user as an Enterprise CAL    
    if ($Mailbox.UMEnabled) 
    { 
        $UMUserCount++ 
        $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
    } 
 
    # MRM Managed Custom Folder usage classifies the user as an Enterprise CAL 
    if ($Mailbox.ManagedFolderMailboxPolicy -ne $null) 
    {      
        $ManagedFolderLinks = (Get-ManagedFolderMailboxPolicy $Mailbox.ManagedFolderMailboxPolicy).ManagedFolderLinks 
        foreach ($FolderLink in $ManagedFolderLinks) { 
            $ManagedFolder = Get-ManagedFolder $FolderLink 
 
            # Managed Custom Folders require an Enterprise CAL 
            If ($ManagedFolder.FolderType -eq "ManagedCustomFolder")  
            { 
                $ManagedCustomFolderUserCount++ 
                $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
 
                break 
            } 
        } 
    } 
 
 
    # Advanced ActiveSync policies classify the user as an Enterprise CAL 
    $CASMailbox = Get-CASMailbox $_ 
    if ($CASMailbox.ActiveSyncEnabled -and ($CASMailbox.ActiveSyncMailboxPolicy -ne $null)) 
    { 
        $ASPolicy = Get-ActiveSyncMailboxPolicy $CASMailbox.ActiveSyncMailboxPolicy 
 
        if (($ASPolicy.AllowDesktopSync -eq $False) -or  
            ($ASPolicy.AllowStorageCard -eq $False) -or 
            ($ASPolicy.AllowCamera -eq $False) -or 
            ($ASPolicy.AllowTextMessaging -eq $False) -or 
            ($ASPolicy.AllowWiFi -eq $False) -or 
            ($ASPolicy.AllowBluetooth -ne "Allow") -or 
            ($ASPolicy.AllowIrDA -eq $False) -or 
            ($ASPolicy.AllowInternetSharing -eq $False) -or 
            ($ASPolicy.AllowRemoteDesktop -eq $False) -or 
            ($ASPolicy.AllowPOPIMAPEmail -eq $False) -or 
            ($ASPolicy.AllowConsumerEmail -eq $False) -or 
            ($ASPolicy.AllowBrowser -eq $False) -or 
            ($ASPolicy.AllowUnsignedApplications -eq $False) -or 
            ($ASPolicy.AllowUnsignedInstallationPackages -eq $False) -or 
            ($ASPolicy.ApprovedApplicationList -ne $null) -or 
            ($ASPolicy.UnapprovedInROMApplicationList -ne $null)) { 
 
            $AdvancedActiveSyncUserCount++ 
            $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
        } 
    } 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Unified Messaging Users calculated:             $UMUserCount" 
    Write-Output "Managed Custom Folder Users calculated:         $ManagedCustomFolderUserCount" 
    Write-Output "Advanced ActiveSync Policy Users calculated:    $AdvancedActiveSyncUserCount" 
} 
 
# 
# Calculate Enterprise CAL users using Journaling 
# 
 
# Help function for function Get-JournalingGroupMailboxMember to traverse members of a DG/DDG/group  
function Traverse-GroupMember 
{ 
    $GroupMember = $args[0] 
     
    if( $GroupMember -eq $null ) 
    { 
        return 
    } 
 
    # Note!!!  
    # Only user, shared and linked mailboxes are counted.  
    # Resource mailboxes and legacy mailboxes are NOT counted. 
    if ( ($GroupMember.RecipientTypeDetails -eq 'UserMailbox') -or 
         ($GroupMember.RecipientTypeDetails -eq 'SharedMailbox') -or 
         ($GroupMember.RecipientTypeDetails -eq 'LinkedMailbox') ) { 
        # Journal one mailbox 
        if ($GroupMember.ExchangeVersion.ToString().Contains(" (8.")) { 
            $JournalingMailboxIDs[$GroupMember.Identity] = $null 
        } 
    } elseif ( ($GroupMember.RecipientType -eq "Group") -or ($GroupMember.RecipientType -like "Dynamic*Group") -or ($GroupMember.RecipientType -like "Mail*Group") ) { 
        # Push this DG/DDG/group into the stack. 
        $DGStack.Push(@($GroupMember.Identity, $GroupMember.RecipientType)) 
    } 
} 
 
# Function that returns all mailbox members including duplicates recursively from a DG/DDG 
function Get-JournalingGroupMailboxMember 
{ 
    # Skip this DG/DDG if it was already enumerated. 
    if ( $VisitedGroups.ContainsKey($args[0]) ) { 
        return 
    } 
     
    $DGStack.Push(@($args[0],$args[1])) 
    while ( $DGStack.Count -ne 0 ) { 
        $StackElement = $DGStack.Pop() 
         
        $GroupIdentity = $StackElement[0] 
        $GroupRecipientType = $StackElement[1] 
 
        if ( $VisitedGroups.ContainsKey($GroupIdentity) ) { 
            # Skip this this DG/DDG if it was already enumerated. 
            continue 
        } 
         
        # Check the members of the current DG/DDG/group in the stack. 
 
        if ( ($GroupRecipientType -like "Mail*Group") -or ($GroupRecipientType -eq "Group" ) ) { 
            $varGroup = Get-Group $GroupIdentity -ErrorAction SilentlyContinue 
            if ( $varGroup -eq $Null ) 
            { 
                $errorMessage = "Invalid group/distribution group/dynamic distribution group: " + $GroupIdentity 
                write-error $errorMessage 
                return 
            } 
             
            $varGroup.members | foreach {     
                # Count users and groups which could be mailboxes. 
                $varGroupMember = Get-User $_ -ErrorAction SilentlyContinue  
                if ( $varGroupMember -eq $Null ) { 
                    $varGroupMember = Get-Group $_ -ErrorAction SilentlyContinue                   
                } 
 
 
                if ( $varGroupMember -ne $Null ) { 
                    Traverse-GroupMember $varGroupMember 
                } 
            } 
        } else { 
            # The current stack element is a DDG. 
            $varGroup = Get-DynamicDistributionGroup $GroupIdentity -ErrorAction SilentlyContinue 
            if ( $varGroup -eq $Null ) 
            { 
                $errorMessage = "Invalid group/distribution group/dynamic distribution group: " + $GroupIdentity 
                write-error $errorMessage 
                return 
            } 
 
            Get-Recipient -RecipientPreviewFilter $varGroup.LdapRecipientFilter -OrganizationalUnit $varGroup.RecipientContainer -ResultSize 'Unlimited' | foreach { 
                Traverse-GroupMember $_ 
            } 
        }  
 
        # Mark this DG/DDG as visited as it's enumerated. 
        $VisitedGroups[$GroupIdentity] = $null 
    }     
} 
 
 
# Check all journaling mailboxes for all journaling rules. 
foreach ($JournalRule in Get-JournalRule){ 
    # There are journal rules in the org. 
 
    if ( $JournalRule.Recipient -eq $Null ) { 
        # One journaling rule journals the whole org (all mailboxes) 
        $OrgWideJournalingEnabled = $True 
        $JournalingUserCount = $TotalMailboxes 
        $TotalEnterpriseCALs = $TotalMailboxes 
 
        break 
    } else { 
        $JournalRecipient = Get-Recipient $JournalRule.Recipient.Local -ErrorAction SilentlyContinue 
 
        if ( $JournalRecipient -ne $Null ) { 
            # Note!!! 
            # Only user, shared and linked mailboxes are counted.  
            # Resource mailboxes and legacy mailboxes are NOT counted. 
            if ( ($JournalRecipient.RecipientTypeDetails -eq 'UserMailbox') -or 
                 ($JournalRecipient.RecipientTypeDetails -eq 'SharedMailbox') -or 
                 ($JournalRecipient.RecipientTypeDetails -eq 'LinkedMailbox') ) { 
 
                # Journal a mailbox 
                if ($JournalRecipient.ExchangeVersion.ToString().Contains(" (8.")) { 
                    $JournalingMailboxIDs[$JournalRecipient.Identity] = $null 
                } 
            } elseif ( ($JournalRecipient.RecipientType -like "Mail*Group") -or ($JournalRecipient.RecipientType -like "Dynamic*Group") ) { 
                # Journal a DG or DDG. 
                # Get all mailbox members for the current journal DG/DDG and add to $JournalingDGMailboxMemberIDs 
                Get-JournalingGroupMailboxMember $JournalRecipient.Identity $JournalRecipient.RecipientType 
                Output-Counts 
            } 
        } 
    } 
} 
 
if ( !$OrgWideJournalingEnabled ) { 
    # No journaling rules journaling the entire org. 
    # Get all journaling mailboxes 
    $JournalingMailboxIDs = Merge-Hashtables $JournalingDGMailboxMemberIDs $JournalingMailboxIDs 
    $JournalingUserCount = $JournalingMailboxIDs.Count 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Journaling Users calculated:                    $JournalingUserCount" 
} 
 
# 
# Calculate Enterprise CALs 
# 
if ( !$OrgWideJournalingEnabled ) { 
    # Calculate Enterprise CALs as not all mailboxes are Enterprise CALs 
 
    $EnterpriseCALMailboxIDs = Merge-Hashtables $JournalingMailboxIDs $EnterpriseCALMailboxIDs 
    $TotalEnterpriseCALs = $EnterpriseCALMailboxIDs.Count 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
    Write-Output "" 
} 
 
 
################### 
## Output Report ## 
################### 
 
Output-Counts 
 
Output-Report 
 
$Global:AdminSessionADSettings.DefaultScope = $OriginalDefaultScope
}

$scriptGetCALReqs2010 = 
{
# Trap block 
trap {  
    Write-Output "An error has occurred running the script:"  
    Write-Output $_ 
 
    Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 
 
 	LogLastException
	
    exit 
}  
 
# Function that returns true if the incoming argument is a help request 
function IsHelpRequest 
{ 
    param($argument) 
    return ($argument -eq "-?" -or $argument -eq "-help"); 
} 
 
# Function that displays the help related to this script following 
# the same format provided by get-help or <cmdletcall> -? 
function Usage 
{ 
@" 
 
NAME: 
`tReportExchangeCALs.ps1 
 
SYNOPSIS: 
`tReports Exchange 2010 client access licenses (CALs) of this organization in Enterprise or Standard categories. 
 
SYNTAX: 
`tReportExchangeCALs.ps1 
 
PARAMETERS: 
 
USAGE: 
`t.\ReportExchangeCALs.ps1 
 
"@ 
} 
 
# Function that resets AdminSessionADSettings.DefaultScope to original value and exits the script 
function Exit-Script 
{ 
    Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 
 
    exit 
} 
 
######################## 
## Script starts here ## 
######################## 
 
$OriginalADServerSetting = Get-ADServerSettings 
 
# Check for Usage Statement Request 
$args | foreach { if (IsHelpRequest $_) { Usage; Exit-Script; } } 
 
# Introduction message 
Write-Output "Report Exchange 2010 client access licenses (CALs) in use in the organization"  
Write-Output "It will take some time if there are a large amount of users......" 
Write-Output "" 
 
Set-ADServerSettings -ViewEntireForest $true 
 
$TotalMailboxes = 0 
$TotalEnterpriseCALs = 0 
$UMUserCount = 0 
$ManagedCustomFolderUserCount = 0 
$AdvancedActiveSyncUserCount = 0 
$ArchiveUserCount = 0 
$RetentionPolicyUserCount = 0 
$AdvancedAntispamUserCount = 0 
$AdvancedAntispamEnabled = $False 
$OrgWideJournalingEnabled = $False 
$AllMailboxIDs = @{} 
$EnterpriseCALMailboxIDs = @{} 
$JournalingUserCount = 0 
$JournalingMailboxIDs = @{} 
$JournalingDGMailboxMemberIDs = @{} 
$DiscoveryConsoleRoles = @{} 
$DiscoveryConsoleRoleAssignees = @() 
$DiscoveryConsoleRoleAssignments = @() 
$SearchableMaiboxIDs = @{} 
$TotalStandardCALs = 0 
$VisitedGroups = @{} 
$DGStack = new-object System.Collections.Stack 
$UserMailboxFilter = "(RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox') -or (RecipientTypeDetails -eq 'LinkedMailbox')" 
# Bool variable for outputing progress information running this script. 
$EnableProgressOutput = $True 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Progress:" 
} 
 
################ 
## Debug code ## 
################ 
 
# Bool variable for output hash table information for debugging purpose. 
$EnableOutputCounts = $False 
 
function Output-Counts 
{ 
    if ($EnableOutputCounts -eq $False) { 
        return 
    } 
 
    Write-Output "Hash Table Name                                 Count" 
    Write-Output "---------------                                 -----" 
    Write-Output "AllMailboxIDs:                                 " $AllMailboxIDs.Count 
    Write-Output "EnterpriseCALMailboxIDs:                       " $EnterpriseCALMailboxIDs.Count 
    Write-Output "JournalingMailboxIDs:                          " $JournalingMailboxIDs.Count 
    Write-Output "JournalingDGMailboxMemberIDs:                  " $JournalingDGMailboxMemberIDs.Count 
    Write-Output "VisitedGroups:                                 " $VisitedGroups.Count 
    Write-Output "" 
    Write-Output "" 
} 
 
function Merge-Hashtables 
{ 
    $Table1 = $args[0] 
    $Table2 = $args[1] 
    $Result = @{} 
     
    if ($null -ne $Table1) 
    { 
        $Result += $Table1 
    } 
 
    if ($null -ne $Table2) 
    { 
        foreach ($entry in $Table2.GetEnumerator()) 
        { 
            $Result[$entry.Key] = $entry.Value 
        } 
    } 
 
    $Result 
} 
 
################# 
## Total Users ## 
################# 
 
# Note!!!  
# Only user, shared and linked mailboxes are counted.  
# Resource mailboxes and legacy mailboxes are NOT counted. 
 
Get-Recipient -ResultSize 'Unlimited' -Filter $UserMailboxFilter | foreach { 
    $Mailbox = $_ 
    if ($Mailbox.ExchangeVersion.ToString().Contains("(14.")) { 
        $AllMailboxIDs[$Mailbox.Identity] = $null 
        $TotalMailboxes++ 
    } 
} 
 
if ($TotalMailboxes -eq 0) { 
    # No mailboxes in the org. Just output the report and exit 
    Output-Report 
     
    Exit-Script 
} 
 
######################### 
## Total Standard CALs ## 
######################### 
 
# All users are counted as Standard CALs 
$TotalStandardCALs = $TotalMailboxes 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Total Standard CALs calculated:                 $TotalStandardCALs" 
} 
 
############################# 
## Per-org Enterprise CALs ## 
############################# 
 
# If advanced anti-spam is turned on, all mailboxes are counted as Enterprise CALs 
foreach ($TransportServer in (Get-TransportServer)) { 
    # If advanced anti-spam is turned on any Hub/Edge server, all mailboxes in the org are counted as Exchange CALs 
    if (!(get-exchangeserver $TransportServer).IsEdgeServer) { 
        $AntispamUpdates = Get-AntispamUpdates $TransportServer 
 
        if (($AntispamUpdates.SpamSignatureUpdatesEnabled -eq $True) -or 
            ($AntispamUpdates.IPReputationUpdatesEnabled -eq $True) -or 
            ($AntispamUpdates.UpdateMode -eq "Automatic")) { 
 
            $AdvancedAntispamEnabled = $True 
            $AdvancedAntispamUserCount = $TotalMailboxes     
            $TotalEnterpriseCALs = $TotalMailboxes 
             
            ## Progress output ...... 
            if ($EnableProgressOutput -eq $True) { 
                Write-Output "Advanced Anti-spam Enabled:                     True" 
                Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
                Write-Output "" 
            } 
 
            # All mailboxes are counted as Enterprise CALs, report and exit. 
            Output-Counts 
             
            Output-Report 
 
            Exit-Script 
        } 
    } 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Advanced Anti-spam Enabled:                     False" 
} 
 
# If Info Leakage Protection is enabled on any transport rule, all mailboxes in the org are counted as Enterprise CALs 
Get-TransportRule | foreach { 
    if ($_.ApplyRightsProtectionTemplate -ne $null) { 
        $TotalEnterpriseCALs = $TotalMailboxes 
         
        ## Progress output ...... 
        if ($EnableProgressOutput -eq $True) { 
            Write-Output "Info Leakage Protection Enabled:                True" 
            Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
            Write-Output "" 
        } 
		
		$InfoLeakageProtectionEnabled = $true
 
        # All mailboxes are counted as Enterprise CALs, report and exit. 
        Output-Counts 
         
        Output-Report 
 
        Exit-Script 
    } 
} 

$InfoLeakageProtectionEnabled = $false 

## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Info Leakage Protection Enabled:                False" 
} 
 
############################## 
## Per-user Enterprise CALs ## 
############################## 
 
# 
# Calculate Enterprise CAL users using UM, MRM Managed Custom Folder, and advanced ActiveSync policy settings 
# 
$AllMailboxIDs.Keys | foreach {   
    $Mailbox = Get-Mailbox $_ 
     
     # UM usage classifies the user as an Enterprise CAL    
    if ($Mailbox.UMEnabled) 
    { 
        $UMUserCount++ 
        $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
    } 
     
    # Archive Mailbox classifies the user as an Enterprise CAL 
    if ($Mailbox.ArchiveGuid -ne [System.Guid]::Empty) { 
        $ArchiveUserCount++ 
        $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
    } 
     
    # Retention Policy classifies the user as an Enterprise CAL 
    if ($Mailbox.RetentionPolicy -ne $null) { 
        $RetentionPolicyUserCount++ 
        $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
    } 
 
    # MRM Managed Custom Folder usage classifies the user as an Enterprise CAL 
    if ($Mailbox.ManagedFolderMailboxPolicy -ne $null) 
    {      
        $ManagedFolderLinks = (Get-ManagedFolderMailboxPolicy $Mailbox.ManagedFolderMailboxPolicy).ManagedFolderLinks 
        foreach ($FolderLink in $ManagedFolderLinks) { 
            $ManagedFolder = Get-ManagedFolder $FolderLink 
 
            # Managed Custom Folders require an Enterprise CAL 
            If ($ManagedFolder.FolderType -eq "ManagedCustomFolder")  
            { 
                $ManagedCustomFolderUserCount++ 
                $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
 
                break 
            } 
        } 
    } 
 
    # Advanced ActiveSync policies classify the user as an Enterprise CAL 
    $CASMailbox = Get-CASMailbox $_ 
    if ($CASMailbox.ActiveSyncEnabled -and ($CASMailbox.ActiveSyncMailboxPolicy -ne $null)) 
    { 
        $ASPolicy = Get-ActiveSyncMailboxPolicy $CASMailbox.ActiveSyncMailboxPolicy 
 
    if (($ASPolicy.AllowDesktopSync -eq $False) -or  
            ($ASPolicy.AllowStorageCard -eq $False) -or 
            ($ASPolicy.AllowCamera -eq $False) -or 
            ($ASPolicy.AllowTextMessaging -eq $False) -or 
            ($ASPolicy.AllowWiFi -eq $False) -or 
            ($ASPolicy.AllowBluetooth -ne "Allow") -or 
            ($ASPolicy.AllowIrDA -eq $False) -or 
            ($ASPolicy.AllowInternetSharing -eq $False) -or 
            ($ASPolicy.AllowRemoteDesktop -eq $False) -or 
            ($ASPolicy.AllowPOPIMAPEmail -eq $False) -or 
            ($ASPolicy.AllowConsumerEmail -eq $False) -or 
            ($ASPolicy.AllowBrowser -eq $False) -or 
            ($ASPolicy.AllowUnsignedApplications -eq $False) -or 
            ($ASPolicy.AllowUnsignedInstallationPackages -eq $False) -or 
            ($ASPolicy.ApprovedApplicationList -ne $null) -or 
            ($ASPolicy.UnapprovedInROMApplicationList -ne $null)) { 
 
            $AdvancedActiveSyncUserCount++ 
            $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
        } 
    } 
     
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Unified Messaging Users calculated:             $UMUserCount" 
    Write-Output "Managed Custom Folder Users calculated:         $ManagedCustomFolderUserCount" 
    Write-Output "Advanced ActiveSync Policy Users calculated:    $AdvancedActiveSyncUserCount" 
    Write-Output "Archived Mailbox Users calculated:              $ArchiveUserCount" 
    Write-Output "Retention Policy Users calculated:              $RetentionPolicyUserCount" 
} 
 
 
# 
# Calculate Enterprise CAL for e-Discovery 
# 
 
# Get all e-discovery management roles which can perform e-discovery tasks 
("*-mailboxsearch", "search-mailbox") | %{Get-ManagementRole -cmdlet $_} | Sort-Object -Unique | %{$DiscoveryConsoleRoles[$_.Identity] = $_} 
 
# Get all e-discovery management role assigment on users 
foreach ($Role in $DiscoveryConsoleRoles.Values) { 
    foreach ($RoleAssignment in @($Role | Get-ManagementRoleAssignment -Delegating $false -Enabled $true)) { 
            $EffectiveAssignees=@() 
            foreach ($EffectiveUserRoleAssignment in (Get-ManagementRoleAssignment -Identity $RoleAssignment.Identity -GetEffectiveUsers)) { 
                $EffectiveAssignees+=$EffectiveUserRoleAssignment.User 
            } 
            foreach ($EffectiveAssignee in $EffectiveAssignees) { 
                $Assignee = Get-User $EffectiveAssignee -ErrorAction SilentlyContinue 
                $error.Clear() 
                if ($Assignee -ne $null) { 
                    $DiscoveryConsoleRoleAssignees += $Assignee 
                    $DiscoveryConsoleRoleAssignments += $RoleAssignment 
                 } 
            } 
    } 
} 
 
# Get excluded mailboxes 
$ExcludedMailboxes = @{} 
 
$ManagementScopes = @{} 
Get-ManagementScope -Exclusive:$true | where {$_.ScopeRestrictionType -eq "RecipientScope"} | foreach { 
    $ManagementScopes[$_.Identity] = $_ 
    [Microsoft.Exchange.Management.Tasks.GetManagementScope]::StampQueryFilterOnManagementScope($_) 
} 
foreach ($ManagementScope in $ManagementScopes.Values) { 
    $Filter = $UserMailboxFilter 
    if (-not([System.String]::IsNullOrEmpty($ManagementScope.RecipientFilter))) { 
        $Filter = "(" + $ManagementScope.RecipientFilter + ") -and (" + $Filter + ")" 
    } 
    Get-Recipient -ResultSize 'Unlimited'-OrganizationalUnit $ManagementScope.RecipientRoot -Filter $Filter | foreach { 
        if ($_.ExchangeVersion.ToString().Contains("(14.")) { 
            $ExcludedMailboxes[$_.Identity] = $true 
        } 
    } 
} 
 
# Get all searched mailboxes in e-discovery 
for ($i=0; $i -lt $DiscoveryConsoleRoleAssignments.Count; $i++) { 
    $RoleAssignment=$DiscoveryConsoleRoleAssignments[$i] 
    $ManagementScope = $null 
    if (($RoleAssignment.CustomRecipientWriteScope -ne $null) -and ($RoleAssignment.RecipientWriteScope -eq "CustomRecipientScope" -or $RoleAssignment.RecipientWriteScope -eq "ExclusiveRecipientScope")) { 
        $ManagementScope = $ManagementScopes[$RoleAssignment.CustomRecipientWriteScope] 
        if ($ManagementScope -eq $null) { 
            $ManagementScope = Get-ManagementScope $RoleAssignment.CustomRecipientWriteScope 
            [Microsoft.Exchange.Management.Tasks.GetManagementScope]::StampQueryFilterOnManagementScope($ManagementScope) 
            $ManagementScopes[$RoleAssignment.CustomRecipientWriteScope] = $ManagementScope 
        } 
    } 
    $ADScope = [Microsoft.Exchange.Management.RbacTasks.GetManagementRoleAssignment]::GetRecipientWriteADScope( 
        $RoleAssignment,  
        $DiscoveryConsoleRoleAssignees[$i],  
        $ManagementScope) 
    if ($ADScope -ne $null) { 
        $Filter = $UserMailboxFilter 
        $ScopeFilter = $ADScope.GetFilterString() 
        if (-not([System.String]::IsNullOrEmpty($ScopeFilter))) { 
            $Filter = "(" + $ScopeFilter + ") -and (" + $Filter + ")" 
        } 
        Get-Recipient -ResultSize 'Unlimited'-OrganizationalUnit $ADScope.Root -Filter $Filter | foreach { 
            if ($_.ExchangeVersion.ToString().Contains("(14.")) { 
                if ($RoleAssignment.RecipientWriteScope -eq [Microsoft.Exchange.Data.Directory.SystemConfiguration.RecipientWriteScopeType]::ExclusiveRecipientScope) { 
                    $EnterpriseCALMailboxIDs[$_.Identity] = $null 
                    $SearchableMaiboxIDs[$_.Identity] = $null 
                } 
                else { 
                    if (-not($ExcludedMailboxes[$_.Identity] -eq $true)) { 
                        $EnterpriseCALMailboxIDs[$_.Identity] = $null 
                        $SearchableMaiboxIDs[$_.Identity] = $null 
                    } 
                } 
            } 
        } 
    } 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Searchable Users calculated:                   "$SearchableMaiboxIDs.Count 
} 
 
# 
# Calculate Enterprise CAL users using Journaling 
# 
 
# Help function for function Get-JournalingGroupMailboxMember to traverse members of a DG/DDG/group  
function Traverse-GroupMember 
{ 
    $GroupMember = $args[0] 
     
    if( $GroupMember -eq $null ) 
    { 
        return 
    } 
 
    # Note!!!  
    # Only user, shared and linked mailboxes are counted.  
    # Resource mailboxes and legacy mailboxes are NOT counted. 
    if ( ($GroupMember.RecipientTypeDetails -eq 'UserMailbox') -or 
          ($GroupMember.RecipientTypeDetails -eq 'SharedMailbox') -or 
          ($GroupMember.RecipientTypeDetails -eq 'LinkedMailbox') ) { 
        # Journal one mailbox 
        if ($GroupMember.ExchangeVersion.ToString().Contains("(14.")) { 
            $JournalingMailboxIDs[$GroupMember.Identity] = $null 
        } 
    } elseif ( ($GroupMember.RecipientType -eq "Group") -or ($GroupMember.RecipientType -like "Dynamic*Group") -or ($GroupMember.RecipientType -like "Mail*Group") ) { 
        # Push this DG/DDG/group into the stack. 
        $DGStack.Push(@($GroupMember.Identity, $GroupMember.RecipientType)) 
    } 
} 
 
# Function that returns all mailbox members including duplicates recursively from a DG/DDG 
function Get-JournalingGroupMailboxMember 
{ 
    # Skip this DG/DDG if it was already enumerated. 
    if ( $VisitedGroups.ContainsKey($args[0]) ) { 
        return 
    } 
     
    $DGStack.Push(@($args[0],$args[1])) 
    while ( $DGStack.Count -ne 0 ) { 
        $StackElement = $DGStack.Pop() 
         
        $GroupIdentity = $StackElement[0] 
        $GroupRecipientType = $StackElement[1] 
 
        if ( $VisitedGroups.ContainsKey($GroupIdentity) ) { 
            # Skip this this DG/DDG if it was already enumerated. 
            continue 
        } 
         
        # Check the members of the current DG/DDG/group in the stack. 
 
        if ( ($GroupRecipientType -like "Mail*Group") -or ($GroupRecipientType -eq "Group" ) ) { 
            $varGroup = Get-Group $GroupIdentity -ErrorAction SilentlyContinue 
            if ( $varGroup -eq $Null ) 
            { 
                $errorMessage = "Invalid group/distribution group/dynamic distribution group: " + $GroupIdentity 
                write-error $errorMessage 
                return 
            } 
             
            $varGroup.members | foreach {     
                # Count users and groups which could be mailboxes. 
                $varGroupMember = Get-User $_ -ErrorAction SilentlyContinue  
                if ( $varGroupMember -eq $Null ) { 
                    $varGroupMember = Get-Group $_ -ErrorAction SilentlyContinue                   
                } 
 
 
                if ( $varGroupMember -ne $Null ) { 
                    Traverse-GroupMember $varGroupMember 
                } 
            } 
        } else { 
            # The current stack element is a DDG. 
            $varGroup = Get-DynamicDistributionGroup $GroupIdentity -ErrorAction SilentlyContinue 
            if ( $varGroup -eq $Null ) 
            { 
                $errorMessage = "Invalid group/distribution group/dynamic distribution group: " + $GroupIdentity 
                write-error $errorMessage 
                return 
            } 
 
            Get-Recipient -RecipientPreviewFilter $varGroup.LdapRecipientFilter -OrganizationalUnit $varGroup.RecipientContainer -ResultSize 'Unlimited' | foreach { 
                Traverse-GroupMember $_ 
            } 
        }  
 
        # Mark this DG/DDG as visited as it's enumerated. 
        $VisitedGroups[$GroupIdentity] = $null 
    }     
} 
 
# Check all journaling mailboxes for all journaling rules. 
foreach ($JournalRule in Get-JournalRule){ 
    # There are journal rules in the org. 
 
    if ( $JournalRule.Recipient -eq $Null ) { 
        # One journaling rule journals the whole org (all mailboxes) 
        $OrgWideJournalingEnabled = $True 
        $JournalingUserCount = $TotalMailboxes 
        $TotalEnterpriseCALs = $TotalMailboxes 
 
        break 
    } else { 
        $JournalRecipient = Get-Recipient $JournalRule.Recipient.Local -ErrorAction SilentlyContinue 
 
        if ( $JournalRecipient -ne $Null ) { 
            # Note!!! 
            # Only user, shared and linked mailboxes are counted.  
            # Resource mailboxes and legacy mailboxes are NOT counted. 
            if (($JournalRecipient.RecipientTypeDetails -eq 'UserMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'SharedMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'LinkedMailbox') ) { 
 
                # Journal a mailbox 
                if ($JournalRecipient.ExchangeVersion.ToString().Contains("(14.")) { 
                    $JournalingMailboxIDs[$JournalRecipient.Identity] = $null 
                } 
            } elseif ( ($JournalRecipient.RecipientType -like "Mail*Group") -or ($JournalRecipient.RecipientType -like "Dynamic*Group") ) { 
                # Journal a DG or DDG. 
                # Get all mailbox members for the current journal DG/DDG and add to $JournalingDGMailboxMemberIDs 
                Get-JournalingGroupMailboxMember $JournalRecipient.Identity $JournalRecipient.RecipientType 
                Output-Counts 
            } 
        } 
    } 
} 
 
if ( !$OrgWideJournalingEnabled ) { 
    # No journaling rules journaling the entire org. 
    # Get all journaling mailboxes 
    $JournalingMailboxIDs = Merge-Hashtables $JournalingDGMailboxMemberIDs $JournalingMailboxIDs 
    $JournalingUserCount = $JournalingMailboxIDs.Count 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Journaling Users calculated:                    $JournalingUserCount" 
} 
 
 
# 
# Calculate Enterprise CALs 
# 
if ( !$OrgWideJournalingEnabled ) { 
    # Calculate Enterprise CALs as not all mailboxes are Enterprise CALs 
    $EnterpriseCALMailboxIDs = Merge-Hashtables $JournalingMailboxIDs $EnterpriseCALMailboxIDs 
    $TotalEnterpriseCALs = $EnterpriseCALMailboxIDs.Count 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
    Write-Output "" 
} 
 
################### 
## Output Report ## 
################### 
 
Output-Counts 
 
Output-Report 
 
Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 
 
}

$scriptGetCALReqs2010SP1 = 
{
# Trap block 
trap {  
    Write-Output "An error has occurred running the script:"  
    Write-Output $_ 
 
    Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 
 
 	LogLastException
	
    exit 
}  
 
# Function that returns true if the incoming argument is a help request 
function IsHelpRequest 
{ 
    param($argument) 
    return ($argument -eq "-?" -or $argument -eq "-help"); 
} 
 
# Function that displays the help related to this script following 
# the same format provided by get-help or <cmdletcall> -? 
function Usage 
{ 
" "
}
 
# Function that resets AdminSessionADSettings.DefaultScope to original value and exits the script 
function Exit-Script 
{ 
    Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 
 
    exit 
} 
 
######################## 
## Script starts here ## 
######################## 
 
$OriginalADServerSetting = Get-ADServerSettings 
 
# Check for Usage Statement Request 
$args | foreach { if (IsHelpRequest $_) { Usage; Exit-Script; } } 
 
# Introduction message 
Write-Output "Report Exchange 2010 client access licenses (CALs) in use in the organization"  
Write-Output "It will take some time if there are a large amount of users......" 
Write-Output "" 
 
Set-ADServerSettings -ViewEntireForest $true 
 
$TotalMailboxes = 0 
$TotalEnterpriseCALs = 0 
$UMUserCount = 0 
$ManagedCustomFolderUserCount = 0 
$AdvancedActiveSyncUserCount = 0 
$ArchiveUserCount = 0 
$RetentionPolicyUserCount = 0 
$OrgWideJournalingEnabled = $False 
$AllMailboxIDs = @{} 
$AllVersionMailboxIDs = @{} 
$EnterpriseCALMailboxIDs = @{} 
$JournalingUserCount = 0 
$JournalingMailboxIDs = @{} 
$JournalingDGMailboxMemberIDs = @{} 
$DiscoveryConsoleRoles = @{} 
$DiscoveryConsoleRoleAssignees = @() 
$DiscoveryConsoleRoleAssignments = @() 
$SearchableMaiboxIDs = @{} 
$TotalStandardCALs = 0 
$VisitedGroups = @{} 
$DGStack = new-object System.Collections.Stack 
$UserMailboxFilter = "(RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox') -or (RecipientTypeDetails -eq 'LinkedMailbox')" 
# Bool variable for outputing progress information running this script. 
$EnableProgressOutput = $True 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Progress:" 
} 
 
################ 
## Debug code ## 
################ 
 
# Bool variable for output hash table information for debugging purpose. 
$EnableOutputCounts = $False 
 
function Output-Counts 
{ 
    if ($EnableOutputCounts -eq $False) { 
        return 
    } 
 
    Write-Output "Hash Table Name                                 Count" 
    Write-Output "---------------                                 -----" 
    Write-Output "AllMailboxIDs:                                 " $AllMailboxIDs.Count 
    Write-Output "EnterpriseCALMailboxIDs:                       " $EnterpriseCALMailboxIDs.Count 
    Write-Output "JournalingMailboxIDs:                          " $JournalingMailboxIDs.Count 
    Write-Output "JournalingDGMailboxMemberIDs:                  " $JournalingDGMailboxMemberIDs.Count 
    Write-Output "VisitedGroups:                                 " $VisitedGroups.Count 
    Write-Output "" 
    Write-Output "" 
} 
 
function Merge-Hashtables 
{ 
    $Table1 = $args[0] 
    $Table2 = $args[1] 
    $Result = @{} 
     
    if ($null -ne $Table1) 
    { 
        $Result += $Table1 
    } 
 
    if ($null -ne $Table2) 
    { 
        foreach ($entry in $Table2.GetEnumerator()) 
        { 
            $Result[$entry.Key] = $entry.Value 
        } 
    } 
 
    $Result 
} 
 
################# 
## Total Users ## 
################# 
 
# Note!!!  
# Only user, shared and linked mailboxes are counted.  
# Resource mailboxes and legacy mailboxes are NOT counted. 
 
Get-Recipient -ResultSize 'Unlimited' -Filter $UserMailboxFilter | foreach { 
    $Mailbox = $_ 
    #if ($Mailbox.ExchangeVersion.ToString().Contains("(14.")) { 
	if ($Mailbox.ExchangeVersion.ToString().Contains("(14.")) { 
        $AllMailboxIDs[$Mailbox.Identity] = $null 
        $TotalMailboxes++ 
    } 
    $AllVersionMailboxIDs[$Mailbox.Identity] = $null 
} 
 
if ($TotalMailboxes -eq 0) { 
    # No mailboxes in the org. Just output the report and exit 
    Output-Report 
     
    Exit-Script 
} 
 
######################### 
## Total Standard CALs ## 
######################### 
 
# All users are counted as Standard CALs 
$TotalStandardCALs = $TotalMailboxes 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Total Standard CALs calculated:                 $TotalStandardCALs" 
} 
 
############################# 
## Per-org Enterprise CALs ## 
############################# 
 
# If Info Leakage Protection is enabled on any transport rule, all mailboxes in the org are counted as Enterprise CALs 
Get-TransportRule | foreach { 
    if ($_.ApplyRightsProtectionTemplate -ne $null) { 
        $TotalEnterpriseCALs = $TotalMailboxes 
         
        ## Progress output ...... 
        if ($EnableProgressOutput -eq $True) { 
            Write-Output "Info Leakage Protection Enabled:                True" 
            Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
            Write-Output "" 
        } 
		
		$InfoLeakageProtectionEnabled = $true
 
        # All mailboxes are counted as Enterprise CALs, report and exit. 
        Output-Counts 
         
        Output-Report 
 
        Exit-Script 
    } 
} 

$InfoLeakageProtectionEnabled = $false

## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Info Leakage Protection Enabled:                False" 
} 
 
############################## 
## Per-user Enterprise CALs ## 
############################## 
 
# 
# Calculate Enterprise CAL users using UM, MRM Managed Custom Folder, and advanced ActiveSync policy settings 
# 
 
$ManagedFolderMailboxPolicyWithCustomedFolder = @{} 
$mailboxPolicies = Get-ManagedFolderMailboxPolicy  
$mailboxPolicies | foreach { 
    foreach ($FolderId in $_.ManagedFolderLinks) 
    { 
        $ManagedFolder = Get-ManagedFolder $FolderId 
        if ($ManagedFolder.FolderType -eq "ManagedCustomFolder") 
        { 
            $ManagedFolderMailboxPolicyWithCustomedFolder[$_.Identity] = $null 
            break 
        } 
    } 
} 
 
$RetentionPolicyWithPersonalTag = @{} 
$RetentionPolicyWithPersonalTagNonArchive = @{} 
 
$retentionPolies = Get-RetentionPolicy 
$retentionPolies | foreach { 
    foreach ($PolicyTagID in $_.RetentionPolicyTagLinks) 
    { 
        $RetentionPolicyTag = Get-RetentionPolicyTag $PolicyTagID 
        if ($RetentionPolicyTag.Type -eq "Personal") 
        { 
            $RetentionPolicyWithPersonalTag[$_.Identity] = $null 
 
            if ($RetentionPolicyTag.RetentionAction -ne "MoveToArchive") 
            { 
                $RetentionPolicyWithPersonalTagNonArchive[$_.Identity] = $null 
                break; 
            } 
        } 
    } 
} 
 
$ActiveSyncMailboxPolicyWithECALFeature = @{} 
 
$activeSyncMailboxPolicies = Get-ActiveSyncMailboxPolicy 
$activeSyncMailboxPolicies | foreach { 
    $ASPolicy = $_ 
    if (($ASPolicy.AllowDesktopSync -eq $False) -or  
                ($ASPolicy.AllowStorageCard -eq $False) -or 
                ($ASPolicy.AllowCamera -eq $False) -or 
                ($ASPolicy.AllowTextMessaging -eq $False) -or 
                ($ASPolicy.AllowWiFi -eq $False) -or 
                ($ASPolicy.AllowBluetooth -ne "Allow") -or 
                ($ASPolicy.AllowIrDA -eq $False) -or 
                ($ASPolicy.AllowInternetSharing -eq $False) -or 
                ($ASPolicy.AllowRemoteDesktop -eq $False) -or 
                ($ASPolicy.AllowPOPIMAPEmail -eq $False) -or 
                ($ASPolicy.AllowConsumerEmail -eq $False) -or 
                ($ASPolicy.AllowBrowser -eq $False) -or 
                ($ASPolicy.AllowUnsignedApplications -eq $False) -or 
                ($ASPolicy.AllowUnsignedInstallationPackages -eq $False) -or 
                ($ASPolicy.ApprovedApplicationList -ne $null) -or 
                ($ASPolicy.UnapprovedInROMApplicationList -ne $null)) 
                { 
                    $ActiveSyncMailboxPolicyWithECALFeature[$ASPolicy.Identity] = $null 
                } 
} 
 
Get-Recipient -ResultSize 'Unlimited' -Filter $UserMailboxFilter -PropertySet 'ConsoleLargeSet' | foreach {   
    $Mailbox = $_ 
     
    if ($Mailbox.ExchangeVersion.ToString().Contains("(14.")) 
    { 
        # UM usage classifies the user as an Enterprise CAL    
        if ($Mailbox.UMEnabled) 
        { 
            $UMUserCount++ 
            $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
        } 
         
        # LOCAL Archive Mailbox classifies the user as an Enterprise CAL 
        if ($Mailbox.ArchiveState -eq "Local") { 
            $ArchiveUserCount++ 
            $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
        } 
         
        # Retention Policy classifies the user as an Enterprise CAL 
        if (($Mailbox.RetentionPolicy -ne $null) -and $RetentionPolicyWithPersonalTag.Contains($Mailbox.RetentionPolicy)) { 
            # For online archive, we will not consider it as ECAL if it's caused by MoveToAchiveTag 
            if (($Mailbox.ArchiveState -eq "HostedProvisioned") -or ($Mailbox.ArchiveState -eq "HostedPending")) 
            { 
                if ($RetentionPolicyWithPersonalTagNonArchive.Contains($Mailbox.RetentionPolicy)) 
                { 
                    $RetentionPolicyUserCount++ 
                    $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
                } 
            } 
            else 
            { 
                $RetentionPolicyUserCount++ 
                $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
            } 
        } 
 
        # MRM Managed Custom Folder usage classifies the user as an Enterprise CAL 
        if (($Mailbox.ManagedFolderMailboxPolicy -ne $null) -and ($ManagedFolderMailboxPolicyWithCustomedFolder.Contains($Mailbox.ManagedFolderMailboxPolicy))) 
        {     
             $ManagedCustomFolderUserCount++ 
             $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
        } 
    } 
} 
 
# Advanced ActiveSync policies classify the user as an Enterprise CAL 
Get-CASMailbox -ResultSize 'Unlimited' -Filter 'ActiveSyncEnabled -eq $true' | foreach { 
    $CASMailbox = $_ 
 
    if (($CASMailbox.ActiveSyncMailboxPolicy -ne $null) -and $ActiveSyncMailboxPolicyWithECALFeature.Contains($CASMailbox.ActiveSyncMailboxPolicy)) 
    { 
        if ($AllMailboxIDs.Contains($CASMailbox.Identity)) 
        { 
            $AdvancedActiveSyncUserCount++ 
            $EnterpriseCALMailboxIDs[$CASMailbox.Identity] = $null 
        } 
    } 
} 
 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Unified Messaging Users calculated:             $UMUserCount" 
    Write-Output "Managed Custom Folder Users calculated:         $ManagedCustomFolderUserCount" 
    Write-Output "Advanced ActiveSync Policy Users calculated:    $AdvancedActiveSyncUserCount" 
    Write-Output "Archived Mailbox Users calculated:              $ArchiveUserCount" 
    Write-Output "Retention Policy Users calculated:              $RetentionPolicyUserCount" 
} 
 
 
# 
# Calculate Enterprise CAL for e-Discovery 
# 
 
# Get all e-discovery management roles which can perform e-discovery tasks 
$ExcludedMailboxes = @{} 
$ManagementScopes = @{}
if (Get-Command "Get-ManagementRole" -errorAction SilentlyContinue){
	("*-mailboxsearch") | %{Get-ManagementRole -cmdlet $_} | Sort-Object -Unique | %{$DiscoveryConsoleRoles[$_.Identity] = $_}
	
	# Get all e-discovery management role assigment on users 
	foreach ($Role in $DiscoveryConsoleRoles.Values) { 
	    foreach ($RoleAssignment in @($Role | Get-ManagementRoleAssignment -Delegating $false -Enabled $true)) { 
	        $EffectiveAssignees=@() 
	        foreach ($EffectiveUserRoleAssignment in (Get-ManagementRoleAssignment -Identity $RoleAssignment.Identity -GetEffectiveUsers)) { 
	            $EffectiveAssignees+=$EffectiveUserRoleAssignment.User 
	        } 
	        foreach ($EffectiveAssignee in $EffectiveAssignees) { 
	            $Assignee = Get-User $EffectiveAssignee -ErrorAction SilentlyContinue 
	            $error.Clear() 
	            if ($Assignee -ne $null) { 
	                $DiscoveryConsoleRoleAssignees += $Assignee 
	                $DiscoveryConsoleRoleAssignments += $RoleAssignment 
	             } 
	        } 
	    } 
	}
	
	# Get excluded mailboxes 
	Get-ManagementScope -Exclusive:$true | where {$_.ScopeRestrictionType -eq "RecipientScope"} | foreach { 
	    $ManagementScopes[$_.Identity] = $_ 
	    [Microsoft.Exchange.Management.Tasks.GetManagementScope]::StampQueryFilterOnManagementScope($_) 
	} 
	foreach ($ManagementScope in $ManagementScopes.Values) { 
	    $Filter = $UserMailboxFilter 
	    if (-not([System.String]::IsNullOrEmpty($ManagementScope.RecipientFilter))) { 
	        $Filter = "(" + $ManagementScope.RecipientFilter + ") -and (" + $Filter + ")" 
	    } 
	    Get-Recipient -ResultSize 'Unlimited'-OrganizationalUnit $ManagementScope.RecipientRoot -Filter $Filter | foreach { 
	        if ($_.ExchangeVersion.ToString().Contains("(14.")) { 
	            $ExcludedMailboxes[$_.Identity] = $true 
	        } 
	    } 
	} 
	 
	# Get all searched mailboxes in e-discovery 
	for ($i=0; $i -lt $DiscoveryConsoleRoleAssignments.Count; $i++) { 
	    $RoleAssignment=$DiscoveryConsoleRoleAssignments[$i] 
	    $ManagementScope = $null 
	    if (($RoleAssignment.CustomRecipientWriteScope -ne $null) -and ($RoleAssignment.RecipientWriteScope -eq "CustomRecipientScope" -or $RoleAssignment.RecipientWriteScope -eq "ExclusiveRecipientScope")) { 
	        $ManagementScope = $ManagementScopes[$RoleAssignment.CustomRecipientWriteScope] 
	        if ($ManagementScope -eq $null) { 
	            $ManagementScope = Get-ManagementScope $RoleAssignment.CustomRecipientWriteScope 
	            [Microsoft.Exchange.Management.Tasks.GetManagementScope]::StampQueryFilterOnManagementScope($ManagementScope) 
	            $ManagementScopes[$RoleAssignment.CustomRecipientWriteScope] = $ManagementScope 
	        } 
	    } 
	    $ADScope = [Microsoft.Exchange.Management.RbacTasks.GetManagementRoleAssignment]::GetRecipientWriteADScope( 
	        $RoleAssignment,  
	        $DiscoveryConsoleRoleAssignees[$i],  
	        $ManagementScope) 
	    if ($ADScope -ne $null) { 
	        $Filter = $UserMailboxFilter 
	        $ScopeFilter = $ADScope.GetFilterString() 
	        if (-not([System.String]::IsNullOrEmpty($ScopeFilter))) { 
	            $Filter = "(" + $ScopeFilter + ") -and (" + $Filter + ")" 
	        } 
	        Get-Recipient -ResultSize 'Unlimited'-OrganizationalUnit $ADScope.Root -Filter $Filter | foreach { 
	            if ($_.ExchangeVersion.ToString().Contains("(14.")) { 
	                if ($RoleAssignment.RecipientWriteScope -eq [Microsoft.Exchange.Data.Directory.SystemConfiguration.RecipientWriteScopeType]::ExclusiveRecipientScope) { 
	                    $EnterpriseCALMailboxIDs[$_.Identity] = $null 
	                    $SearchableMaiboxIDs[$_.Identity] = $null 
	                } 
	                else { 
	                    if (-not($ExcludedMailboxes[$_.Identity] -eq $true)) { 
	                        $EnterpriseCALMailboxIDs[$_.Identity] = $null 
	                        $SearchableMaiboxIDs[$_.Identity] = $null 
	                    } 
	                } 
	            } 
	        } 
	    } 
	} 
}
else {
	 Write-Output "Warning: Get-ManagementRole cmdlet not available. E-discovery management roles not analysed"
}
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Searchable Users calculated:                   "$SearchableMaiboxIDs.Count 
} 
 
# 
# Calculate Enterprise CAL users using Journaling 
# 
 
# Help function for function Get-JournalingGroupMailboxMember to traverse members of a DG/DDG/group  
function Traverse-GroupMember 
{ 
    $GroupMember = $args[0] 
     
    if( $GroupMember -eq $null ) 
    { 
        return 
    } 
 
    # Note!!!  
    # Only user, shared and linked mailboxes are counted.  
    # Resource mailboxes and legacy mailboxes are NOT counted. 
    if ( ($GroupMember.RecipientTypeDetails -eq 'UserMailbox') -or 
          ($GroupMember.RecipientTypeDetails -eq 'SharedMailbox') -or 
          ($GroupMember.RecipientTypeDetails -eq 'LinkedMailbox') ) { 
        # Journal one mailbox 
        if ($GroupMember.ExchangeVersion.ToString().Contains("(14.")) { 
            $JournalingMailboxIDs[$GroupMember.Identity] = $null 
        } 
    } elseif ( ($GroupMember.RecipientType -eq "Group") -or ($GroupMember.RecipientType -like "Dynamic*Group") -or ($GroupMember.RecipientType -like "Mail*Group") ) { 
        # Push this DG/DDG/group into the stack. 
        $DGStack.Push(@($GroupMember.Identity, $GroupMember.RecipientType)) 
    } 
} 
 
# Function that returns all mailbox members including duplicates recursively from a DG/DDG 
function Get-JournalingGroupMailboxMember 
{ 
    # Skip this DG/DDG if it was already enumerated. 
    if ( $VisitedGroups.ContainsKey($args[0]) ) { 
        return 
    } 
     
    $DGStack.Push(@($args[0],$args[1])) 
    while ( $DGStack.Count -ne 0 ) { 
        $StackElement = $DGStack.Pop() 
         
        $GroupIdentity = $StackElement[0] 
        $GroupRecipientType = $StackElement[1] 
 
        if ( $VisitedGroups.ContainsKey($GroupIdentity) ) { 
            # Skip this this DG/DDG if it was already enumerated. 
            continue 
        } 
         
        # Check the members of the current DG/DDG/group in the stack. 
 
        if ( ($GroupRecipientType -like "Mail*Group") -or ($GroupRecipientType -eq "Group" ) ) { 
            $varGroup = Get-Group $GroupIdentity -ErrorAction SilentlyContinue 
            if ( $varGroup -eq $Null ) 
            { 
                $errorMessage = "Invalid group/distribution group/dynamic distribution group: " + $GroupIdentity 
                write-error $errorMessage 
                return 
            } 
             
            $varGroup.members | foreach {     
                # Count users and groups which could be mailboxes. 
                $varGroupMember = Get-User $_ -ErrorAction SilentlyContinue  
                if ( $varGroupMember -eq $Null ) { 
                    $varGroupMember = Get-Group $_ -ErrorAction SilentlyContinue                   
                } 
 
 
                if ( $varGroupMember -ne $Null ) { 
                    Traverse-GroupMember $varGroupMember 
                } 
            } 
        } else { 
            # The current stack element is a DDG. 
            $varGroup = Get-DynamicDistributionGroup $GroupIdentity -ErrorAction SilentlyContinue 
            if ( $varGroup -eq $Null ) 
            { 
                $errorMessage = "Invalid group/distribution group/dynamic distribution group: " + $GroupIdentity 
                write-error $errorMessage 
                return 
            } 
 
            Get-Recipient -RecipientPreviewFilter $varGroup.LdapRecipientFilter -OrganizationalUnit $varGroup.RecipientContainer -ResultSize 'Unlimited' | foreach { 
                Traverse-GroupMember $_ 
            } 
        }  
 
        # Mark this DG/DDG as visited as it's enumerated. 
        $VisitedGroups[$GroupIdentity] = $null 
    }     
} 
 
# Check all journaling mailboxes(include all version) for all journaling rules, and count E2010 mailbox as Enterprise CALs. 
foreach ($JournalRule in Get-JournalRule){ 
    # There are journal rules in the org. 
 
    if ( $JournalRule.Recipient -eq $Null ) { 
        # One journaling rule journals the whole org (all mailboxes) 
        $OrgWideJournalingEnabled = $True 
        $JournalingUserCount = $AllVersionMailboxIDs.Count 
        $TotalEnterpriseCALs = $TotalMailboxes 
 
        break 
    } else { 
        $JournalRecipient = Get-Recipient -Filter ("((PrimarySmtpAddress -eq '" + $JournalRule.Recipient + "'))") 
 
        if ( $JournalRecipient -ne $Null ) { 
            # Note!!! 
            # Remote mailbox is NOT count here since it's totally different story. 
            if (($JournalRecipient.RecipientTypeDetails -eq 'UserMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'SharedMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'LinkedMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'MailContact') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'PublicFolder') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'LegacyMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'RoomMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'EquipmentMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'MailForestContact') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'MailUser')) { 
 
                # Journal a mailbox 
                if ($JournalRecipient.ExchangeVersion.ToString().Contains("(14.")) { 
                    $JournalingMailboxIDs[$JournalRecipient.Identity] = $null 
                } 
            } elseif ( ($JournalRecipient.RecipientType -like "Mail*Group") -or ($JournalRecipient.RecipientType -like "Dynamic*Group") ) { 
                # Journal a DG or DDG. 
                # Get all mailbox members for the current journal DG/DDG and add to $JournalingDGMailboxMemberIDs 
                Get-JournalingGroupMailboxMember $JournalRecipient.Identity $JournalRecipient.RecipientType 
                Output-Counts 
            } 
        } 
    } 
} 
 
if ( !$OrgWideJournalingEnabled ) { 
    # No journaling rules journaling the entire org. 
    # Get all journaling mailboxes 
    $JournalingMailboxIDs = Merge-Hashtables $JournalingDGMailboxMemberIDs $JournalingMailboxIDs 
    $JournalingUserCount = $JournalingMailboxIDs.Count 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Journaling Users calculated:                    $JournalingUserCount" 
} 
 
 
# 
# Calculate Enterprise CALs 
# 
if ( !$OrgWideJournalingEnabled ) { 
    # Calculate Enterprise CALs as not all mailboxes are Enterprise CALs 
    foreach ($journalingMailboxID in $JournalingMailboxIDs.Keys) { 
        if ($AllMailboxIDs.Contains($journalingMailboxID)) { 
            $EnterpriseCALMailboxIDs[$journalingMailboxID] = $null 
        } 
    } 
    $TotalEnterpriseCALs = $EnterpriseCALMailboxIDs.Count 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
    Write-Output "" 
}

################### 
## Output Report ## 
################### 

Output-Counts 
 
Output-Report 
 
Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 

}


$scriptGetCALReqs2010SP3 = 
{
# Trap block 
trap {  
    Write-Output "An error has occurred running the script:"  
    Write-Output $_ 
 
    Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 
 
 	LogLastException
	
    exit 
}  
 
# Function that returns true if the incoming argument is a help request 
function IsHelpRequest 
{ 
    param($argument) 
    return ($argument -eq "-?" -or $argument -eq "-help"); 
} 
 
# Function that displays the help related to this script following 
# the same format provided by get-help or <cmdletcall> -? 
function Usage 
{ 
@" 
 
NAME: 
`tReportExchangeCALs.ps1 
 
SYNOPSIS: 
`tReports Exchange 2010 SP3 client access licenses (CALs) of this organization in Enterprise or Standard categories. 
 
SYNTAX: 
`tReportExchangeCALs.ps1 
 
PARAMETERS: 
 
USAGE: 
`t.\ReportExchangeCALs.ps1 
 
"@ 
} 
 
# Function that resets AdminSessionADSettings.DefaultScope to original value and exits the script 
function Exit-Script 
{ 
    Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 
 
    exit 
} 
 
######################## 
## Script starts here ## 
######################## 
 
$OriginalADServerSetting = Get-ADServerSettings 
 
# Check for Usage Statement Request 
$args | foreach { if (IsHelpRequest $_) { Usage; Exit-Script; } } 
 
# Introduction message 
Write-Output "Report Exchange 2010 SP3 client access licenses (CALs) in use in the organization"  
Write-Output "It will take some time if there are a large amount of users......" 
Write-Output "" 
 
Set-ADServerSettings -ViewEntireForest $true 
 
$TotalMailboxes = 0 
$TotalEnterpriseCALs = 0 
$UMUserCount = 0 
# Oliver Moazzezi - removed Managed Custom Folder count
#$ManagedCustomFolderUserCount = 0 
$AdvancedActiveSyncUserCount = 0 
$DataLeakPreventionUserCount = 0
$ArchiveUserCount = 0 
$RetentionPolicyUserCount = 0 
$OrgWideJournalingEnabled = $False 
$AllMailboxIDs = @{} 
$AllVersionMailboxIDs = @{} 
$EnterpriseCALMailboxIDs = @{} 
$JournalingUserCount = 0 
$JournalingMailboxIDs = @{} 
$JournalingDGMailboxMemberIDs = @{} 
$DiscoveryConsoleRoles = @{} 
$DiscoveryConsoleRoleAssignees = @() 
$DiscoveryConsoleRoleAssignments = @() 
$SearchableMaiboxIDs = @{} 
$TotalStandardCALs = 0 
$VisitedGroups = @{} 
$DGStack = new-object System.Collections.Stack 
$UserMailboxFilter = "(RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox') -or (RecipientTypeDetails -eq 'LinkedMailbox')" 
# Bool variable for outputing progress information running this script. 
$EnableProgressOutput = $True 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Progress:" 
} 
 
################ 
## Debug code ## 
################ 
 
# Bool variable for output hash table information for debugging purpose. 
$EnableOutputCounts = $False 
 
function Output-Counts 
{ 
    if ($EnableOutputCounts -eq $False) { 
        return 
    } 
 
    Write-Output "Hash Table Name                                 Count" 
    Write-Output "---------------                                 -----" 
    Write-Output "AllMailboxIDs:                                 " $AllMailboxIDs.Count 
    Write-Output "EnterpriseCALMailboxIDs:                       " $EnterpriseCALMailboxIDs.Count 
    Write-Output "JournalingMailboxIDs:                          " $JournalingMailboxIDs.Count 
    Write-Output "JournalingDGMailboxMemberIDs:                  " $JournalingDGMailboxMemberIDs.Count 
    Write-Output "VisitedGroups:                                 " $VisitedGroups.Count 
    Write-Output "" 
    Write-Output "" 
} 
 
function Merge-Hashtables 
{ 
    $Table1 = $args[0] 
    $Table2 = $args[1] 
    $Result = @{} 
     
    if ($null -ne $Table1) 
    { 
        $Result += $Table1 
    } 
 
    if ($null -ne $Table2) 
    { 
        foreach ($entry in $Table2.GetEnumerator()) 
        { 
            $Result[$entry.Key] = $entry.Value 
        } 
    } 
 
    $Result 
} 
 
################# 
## Total Users ## 
################# 
 
# Note!!!  
# Only user, shared and linked mailboxes are counted.  
# Resource mailboxes and legacy mailboxes are NOT counted. 
# Oliver Moazzezi - I have set the equals value to 15, for Exchange 2013
 
Get-Recipient -ResultSize 'Unlimited' -Filter $UserMailboxFilter | foreach { 
    $Mailbox = $_ 
    if ($Mailbox.ExchangeVersion.ToString().Contains("(14.")) { 
        $AllMailboxIDs[$Mailbox.Identity] = $null 
        $TotalMailboxes++ 
    } 
    $AllVersionMailboxIDs[$Mailbox.Identity] = $null 
} 
 
if ($TotalMailboxes -eq 0) { 
    # No mailboxes in the org. Just output the report and exit 
    Output-Report 
     
    Exit-Script 
} 
 
######################### 
## Total Standard CALs ## 
######################### 
 
# All users are counted as Standard CALs 
$TotalStandardCALs = $TotalMailboxes 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Total Standard CALs calculated:                 $TotalStandardCALs" 
} 
 
############################# 
## Per-org Enterprise CALs ## 
############################# 
 
# If Info Leakage Protection is enabled on any transport rule, all mailboxes in the org are counted as Enterprise CALs 
Get-TransportRule | foreach { 
    if ($_.ApplyRightsProtectionTemplate -ne $null) { 
        $TotalEnterpriseCALs = $TotalMailboxes 
         
        ## Progress output ...... 
        if ($EnableProgressOutput -eq $True) { 
            Write-Output "Info Leakage Protection Enabled:                True" 
            Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
            Write-Output "" 
        }
		
		$InfoLeakageProtectionEnabled = $true
 
        # All mailboxes are counted as Enterprise CALs, report and exit. 
        Output-Counts 
         
        Output-Report 
 
        Exit-Script 
    } 
} 

$InfoLeakageProtectionEnabled = $false
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Info Leakage Protection Enabled:                False" 
} 
 
############################## 
## Per-user Enterprise CALs ## 
############################## 
 
# 
# Calculate Enterprise CAL users using UM, MRM Managed Custom Folder, and advanced ActiveSync policy settings 
# 
# Managed Folders are discontinued in Exchange 2013 http://technet.microsoft.com/en-us/library/jj619283%28v=exchg.150%29.aspx 
#$ManagedFolderMailboxPolicyWithCustomedFolder = @{} 
#$mailboxPolicies = Get-ManagedFolderMailboxPolicy  
#$mailboxPolicies | foreach { 
#    foreach ($FolderId in $_.ManagedFolderLinks) 
#    { 
#        $ManagedFolder = Get-ManagedFolder $FolderId 
#        if ($ManagedFolder.FolderType -eq "ManagedCustomFolder") 
#        { 
#            $ManagedFolderMailboxPolicyWithCustomedFolder[$_.Identity] = $null 
#            break 
#        } 
#    } 
#} 
 
$RetentionPolicyWithPersonalTag = @{} 
$RetentionPolicyWithPersonalTagNonArchive = @{} 
 
$retentionPolies = Get-RetentionPolicy 
$retentionPolies | foreach { 
    foreach ($PolicyTagID in $_.RetentionPolicyTagLinks) 
    { 
        $RetentionPolicyTag = Get-RetentionPolicyTag $PolicyTagID 
        if ($RetentionPolicyTag.Type -eq "Personal") 
        { 
            $RetentionPolicyWithPersonalTag[$_.Identity] = $null 
 
            if ($RetentionPolicyTag.RetentionAction -ne "MoveToArchive") 
            { 
                $RetentionPolicyWithPersonalTagNonArchive[$_.Identity] = $null 
                break; 
            } 
        } 
    } 
} 
 
$ActiveSyncMailboxPolicyWithECALFeature = @{} 
 
$activeSyncMailboxPolicies = Get-MobileDeviceMailboxPolicy 
$activeSyncMailboxPolicies | foreach { 
    $ASPolicy = $_ 
    if (($ASPolicy.AllowDesktopSync -eq $False) -or  
                ($ASPolicy.AllowStorageCard -eq $False) -or 
                ($ASPolicy.AllowCamera -eq $False) -or 
                ($ASPolicy.AllowTextMessaging -eq $False) -or 
                ($ASPolicy.AllowWiFi -eq $False) -or 
                ($ASPolicy.AllowBluetooth -ne "Allow") -or 
                ($ASPolicy.AllowIrDA -eq $False) -or 
                ($ASPolicy.AllowInternetSharing -eq $False) -or 
                ($ASPolicy.AllowRemoteDesktop -eq $False) -or 
                ($ASPolicy.AllowPOPIMAPEmail -eq $False) -or 
                ($ASPolicy.AllowConsumerEmail -eq $False) -or 
                ($ASPolicy.AllowBrowser -eq $False) -or 
                ($ASPolicy.AllowUnsignedApplications -eq $False) -or 
                ($ASPolicy.AllowUnsignedInstallationPackages -eq $False) -or 
                ($ASPolicy.ApprovedApplicationList -ne $null) -or 
                ($ASPolicy.UnapprovedInROMApplicationList -ne $null)) 
                { 
                    $ActiveSyncMailboxPolicyWithECALFeature[$ASPolicy.Identity] = $null 
                } 
} 
 
Get-Recipient -ResultSize 'Unlimited' -Filter $UserMailboxFilter -PropertySet 'ConsoleLargeSet' | foreach {   
    $Mailbox = $_ 
     
    if ($Mailbox.ExchangeVersion.ToString().Contains("(14.")) 
    { 
        # UM usage classifies the user as an Enterprise CAL    
        if ($Mailbox.UMEnabled) 
        { 
            $UMUserCount++ 
            $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
        } 
         
        # LOCAL Archive Mailbox classifies the user as an Enterprise CAL 
        if ($Mailbox.ArchiveState -eq "Local") { 
            $ArchiveUserCount++ 
            $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
        } 
         
        # Retention Policy classifies the user as an Enterprise CAL 
        if (($Mailbox.RetentionPolicy -ne $null) -and $RetentionPolicyWithPersonalTag.Contains($Mailbox.RetentionPolicy)) { 
            # For online archive, we will not consider it as ECAL if it's caused by MoveToAchiveTag 
            if (($Mailbox.ArchiveState -eq "HostedProvisioned") -or ($Mailbox.ArchiveState -eq "HostedPending")) 
            { 
                if ($RetentionPolicyWithPersonalTagNonArchive.Contains($Mailbox.RetentionPolicy)) 
                { 
                    $RetentionPolicyUserCount++ 
                    $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
                } 
            } 
            else 
            { 
                $RetentionPolicyUserCount++ 
                $EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
            } 
        } 
# Oliver Moazzezi - Managed Folder Policies not support 
        # MRM Managed Custom Folder usage classifies the user as an Enterprise CAL 
        #if (($Mailbox.ManagedFolderMailboxPolicy -ne $null) -and ($ManagedFolderMailboxPolicyWithCustomedFolder.Contains($Mailbox.ManagedFolderMailboxPolicy))) 
        #{     
             #$ManagedCustomFolderUserCount++ 
             #$EnterpriseCALMailboxIDs[$Mailbox.Identity] = $null 
        #} 
    } 
} 
 
# Advanced ActiveSync policies classify the user as an Enterprise CAL 
# Oliver Moazzezi - modified to use Get-MobileDeviceMailboxPolicy
Get-CASMailbox -ResultSize 'Unlimited' -Filter 'ActiveSyncEnabled -eq $true' | foreach { 
    $CASMailbox = $_ 
 
    if (($CASMailbox.MobileDeviceMailboxPolicy -ne $null) -and $ActiveSyncMailboxPolicyWithECALFeature.Contains($CASMailbox.MobileDeviceSyncMailboxPolicy)) 
    { 
        if ($AllMailboxIDs.Contains($CASMailbox.Identity)) 
        { 
            $AdvancedActiveSyncUserCount++ 
            $EnterpriseCALMailboxIDs[$CASMailbox.Identity] = $null 
        } 
    } 
} 
 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Unified Messaging Users calculated:             $UMUserCount" 
# Oliver Moazzezi - Removed Managed Custom Folder calculations
   #Write-Output "Managed Custom Folder Users calculated:         $ManagedCustomFolderUserCount" 
    Write-Output "Advanced ActiveSync Policy Users calculated:    $AdvancedActiveSyncUserCount" 
    Write-Output "Archived Mailbox Users calculated:              $ArchiveUserCount" 
    Write-Output "Retention Policy Users calculated:              $RetentionPolicyUserCount" 
} 
 
 
# 
# Calculate Enterprise CAL for e-Discovery 
# 
 
# Get all e-discovery management roles which can perform e-discovery tasks 
("*-mailboxsearch") | %{Get-ManagementRole -cmdlet $_} | Sort-Object -Unique | %{$DiscoveryConsoleRoles[$_.Identity] = $_} 
 
# Get all e-discovery management role assigment on users 
foreach ($Role in $DiscoveryConsoleRoles.Values) { 
    foreach ($RoleAssignment in @($Role | Get-ManagementRoleAssignment -Delegating $false -Enabled $true)) { 
            $EffectiveAssignees=@() 
            foreach ($EffectiveUserRoleAssignment in (Get-ManagementRoleAssignment -Identity $RoleAssignment.Identity -GetEffectiveUsers)) { 
                $EffectiveAssignees+=$EffectiveUserRoleAssignment.User 
            } 
            foreach ($EffectiveAssignee in $EffectiveAssignees) { 
                $Assignee = Get-User $EffectiveAssignee -ErrorAction SilentlyContinue 
                $error.Clear() 
                if ($Assignee -ne $null) { 
                    $DiscoveryConsoleRoleAssignees += $Assignee 
                    $DiscoveryConsoleRoleAssignments += $RoleAssignment 
                 } 
            } 
    } 
} 
 
# Get excluded mailboxes 
$ExcludedMailboxes = @{} 
 
$ManagementScopes = @{} 
Get-ManagementScope -Exclusive:$true | where {$_.ScopeRestrictionType -eq "RecipientScope"} | foreach { 
    $ManagementScopes[$_.Identity] = $_ 
    [Microsoft.Exchange.Management.Tasks.GetManagementScope]::StampQueryFilterOnManagementScope($_) 
} 
foreach ($ManagementScope in $ManagementScopes.Values) { 
    $Filter = $UserMailboxFilter 
    if (-not([System.String]::IsNullOrEmpty($ManagementScope.RecipientFilter))) { 
        $Filter = "(" + $ManagementScope.RecipientFilter + ") -and (" + $Filter + ")" 
    } 
    Get-Recipient -ResultSize 'Unlimited'-OrganizationalUnit $ManagementScope.RecipientRoot -Filter $Filter | foreach { 
        if ($_.ExchangeVersion.ToString().Contains("(14.")) { 
            $ExcludedMailboxes[$_.Identity] = $true 
        } 
    } 
} 
 
# Oliver Moazzezi - multi mailbox search is not an Enterprise CAL feature anymore http://office.microsoft.com/en-gb/exchange/microsoft-exchange-server-licensing-licensing-overview-FX103746915.aspx

## Progress output ...... 
#if ($EnableProgressOutput -eq $True) { 
  #  Write-Output "Searchable Users calculated:                    Coming Soon" 
#} 
 
# 
# Calculate Enterprise CAL users using Journaling 
# 
 
# Help function for function Get-JournalingGroupMailboxMember to traverse members of a DG/DDG/group  
function Traverse-GroupMember 
{ 
    $GroupMember = $args[0] 
     
    if( $GroupMember -eq $null ) 
    { 
        return 
    } 
 
    # Note!!!  
    # Only user, shared and linked mailboxes are counted.  
    # Resource mailboxes and legacy mailboxes are NOT counted. 
    if ( ($GroupMember.RecipientTypeDetails -eq 'UserMailbox') -or 
          ($GroupMember.RecipientTypeDetails -eq 'SharedMailbox') -or 
          ($GroupMember.RecipientTypeDetails -eq 'LinkedMailbox') ) { 
        # Journal one mailbox 
        if ($GroupMember.ExchangeVersion.ToString().Contains("(14.")) { 
            $JournalingMailboxIDs[$GroupMember.Identity] = $null 
        } 
    } elseif ( ($GroupMember.RecipientType -eq "Group") -or ($GroupMember.RecipientType -like "Dynamic*Group") -or ($GroupMember.RecipientType -like "Mail*Group") ) { 
        # Push this DG/DDG/group into the stack. 
        $DGStack.Push(@($GroupMember.Identity, $GroupMember.RecipientType)) 
    } 
} 
 
# Function that returns all mailbox members including duplicates recursively from a DG/DDG 
function Get-JournalingGroupMailboxMember 
{ 
    # Skip this DG/DDG if it was already enumerated. 
    if ( $VisitedGroups.ContainsKey($args[0]) ) { 
        return 
    } 
     
    $DGStack.Push(@($args[0],$args[1])) 
    while ( $DGStack.Count -ne 0 ) { 
        $StackElement = $DGStack.Pop() 
         
        $GroupIdentity = $StackElement[0] 
        $GroupRecipientType = $StackElement[1] 
 
        if ( $VisitedGroups.ContainsKey($GroupIdentity) ) { 
            # Skip this this DG/DDG if it was already enumerated. 
            continue 
        } 
         
        # Check the members of the current DG/DDG/group in the stack. 
 
        if ( ($GroupRecipientType -like "Mail*Group") -or ($GroupRecipientType -eq "Group" ) ) { 
            $varGroup = Get-Group $GroupIdentity -ErrorAction SilentlyContinue 
            if ( $varGroup -eq $Null ) 
            { 
                $errorMessage = "Invalid group/distribution group/dynamic distribution group: " + $GroupIdentity 
                write-error $errorMessage 
                return 
            } 
             
            $varGroup.members | foreach {     
                # Count users and groups which could be mailboxes. 
                $varGroupMember = Get-User $_ -ErrorAction SilentlyContinue  
                if ( $varGroupMember -eq $Null ) { 
                    $varGroupMember = Get-Group $_ -ErrorAction SilentlyContinue                   
                } 
 
 
                if ( $varGroupMember -ne $Null ) { 
                    Traverse-GroupMember $varGroupMember 
                } 
            } 
        } else { 
            # The current stack element is a DDG. 
            $varGroup = Get-DynamicDistributionGroup $GroupIdentity -ErrorAction SilentlyContinue 
            if ( $varGroup -eq $Null ) 
            { 
                $errorMessage = "Invalid group/distribution group/dynamic distribution group: " + $GroupIdentity 
                write-error $errorMessage 
                return 
            } 
 
            Get-Recipient -RecipientPreviewFilter $varGroup.LdapRecipientFilter -OrganizationalUnit $varGroup.RecipientContainer -ResultSize 'Unlimited' | foreach { 
                Traverse-GroupMember $_ 
            } 
        }  
 
        # Mark this DG/DDG as visited as it's enumerated. 
        $VisitedGroups[$GroupIdentity] = $null 
    }     
} 
 
# Check all journaling mailboxes(include all version) for all journaling rules, and count E2010 mailbox as Enterprise CALs. 
foreach ($JournalRule in Get-JournalRule){ 
    # There are journal rules in the org. 
 
    if ( $JournalRule.Recipient -eq $Null ) { 
        # One journaling rule journals the whole org (all mailboxes) 
        $OrgWideJournalingEnabled = $True 
        $JournalingUserCount = $AllVersionMailboxIDs.Count 
        $TotalEnterpriseCALs = $TotalMailboxes 
 
        break 
    } else { 
        $JournalRecipient = Get-Recipient -Filter ("((PrimarySmtpAddress -eq '" + $JournalRule.Recipient + "'))") 
 
        if ( $JournalRecipient -ne $Null ) { 
            # Note!!! 
            # Remote mailbox is NOT count here since it's totally different story. 
            if (($JournalRecipient.RecipientTypeDetails -eq 'UserMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'SharedMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'LinkedMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'MailContact') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'PublicFolder') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'LegacyMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'RoomMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'EquipmentMailbox') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'MailForestContact') -or 
                ($JournalRecipient.RecipientTypeDetails -eq 'MailUser')) { 
 
                # Journal a mailbox 
                if ($JournalRecipient.ExchangeVersion.ToString().Contains("(14.")) { 
                    $JournalingMailboxIDs[$JournalRecipient.Identity] = $null 
                } 
            } elseif ( ($JournalRecipient.RecipientType -like "Mail*Group") -or ($JournalRecipient.RecipientType -like "Dynamic*Group") ) { 
                # Journal a DG or DDG. 
                # Get all mailbox members for the current journal DG/DDG and add to $JournalingDGMailboxMemberIDs 
                Get-JournalingGroupMailboxMember $JournalRecipient.Identity $JournalRecipient.RecipientType 
                Output-Counts 
            } 
        } 
    } 
} 
 
if ( !$OrgWideJournalingEnabled ) { 
    # No journaling rules journaling the entire org. 
    # Get all journaling mailboxes 
    $JournalingMailboxIDs = Merge-Hashtables $JournalingDGMailboxMemberIDs $JournalingMailboxIDs 
    $JournalingUserCount = $JournalingMailboxIDs.Count 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Journaling Users calculated:                    $JournalingUserCount" 
    
} 
 
#Oliver Moazzezi - DLP section (coming soon..)
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Data Loss Prevention Users Calculated:          Manual Check. Enterprise CAL required for feature"
} 
 
# 
# Calculate Enterprise CALs 
# 
if ( !$OrgWideJournalingEnabled ) { 
    # Calculate Enterprise CALs as not all mailboxes are Enterprise CALs 
    foreach ($journalingMailboxID in $JournalingMailboxIDs.Keys) { 
        if ($AllMailboxIDs.Contains($journalingMailboxID)) { 
            $EnterpriseCALMailboxIDs[$journalingMailboxID] = $null 
        } 
    } 
    $TotalEnterpriseCALs = $EnterpriseCALMailboxIDs.Count 
} 
 
## Progress output ...... 
if ($EnableProgressOutput -eq $True) { 
    Write-Output "Total Enterprise CALs calculated:               $TotalEnterpriseCALs" 
 
    Write-Output "" 
} 
 
################### 
## Output Report ## 
################### 
 
Output-Counts 
 
Output-Report 
 
Set-ADServerSettings -ViewEntireForest $OriginalADServerSetting.ViewEntireForest -RecipientViewRoot $OriginalADServerSetting.RecipientViewRoot 

}

$scriptGetCALReqs2013 = 
{
	$TotalStandardCALs = Get-ExchangeServerAccessLicenseUser -LicenseName (Get-ExchangeServerAccessLicense | ? {($_.UnitLabel -eq "CAL") -and ($_.LicenseName -like "*Standard*")}).licenseName | measure | select Count
    $TotalEnterpriseCALs = Get-ExchangeServerAccessLicenseUser -LicenseName (Get-ExchangeServerAccessLicense | ? {($_.UnitLabel -eq "CAL") -and ($_.LicenseName -like "*Enterprise*")}).licenseName | measure | select Count
	Output-Report

	$ExchangeLicenses = @()
	$ExchangeLicenseTypes = Get-ExchangeServerAccessLicense
	foreach ($ExchangeLicenseType in $ExchangeLicenseTypes){
		$ExchangeLicenses += Get-ExchangeServerAccessLicenseUser -LicenseName $ExchangeLicenseType.LicenseName 
	}

	if ($ExchangeLicenses){
		$ExchangeLicenses | export-csv $OutputFile5 -notypeinformation -Encoding UTF8
	}
}


try {
	Get-ExchangeDetails
}
catch {
	LogLastException
}

exit


