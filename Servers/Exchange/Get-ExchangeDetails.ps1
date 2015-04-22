 ##########################################################################
 #
 # Get-ExchangeDetails
 # Sam Gold Toolkit
 # Original Source: Sam360
 #
 ##########################################################################

 Param(
	[alias("server")]
	$ExchangeServer = [System.Net.Dns]::GetHostByName(($env:computerName)).HostName,
	[alias("o1")]
	$OutputFile1 = "ExchangeServerDetails.csv",
	[alias("o2")]
	$OutputFile2 = "ExchangeUsers.csv",
	[alias("o3")]
	$OutputFile3 = "ExchangeUserActivity.csv",
	[alias("o4")]
	$OutputFile4 = "ExchangeMailBoxes.csv",
	[alias("o5")]
	$OutputFile5 = "ExchangeMailBoxActivity.csv",
	[alias("o6")]
	$OutputFile6 = "ExchangeDevices.csv",
	[alias("o7")]
	$OutputFile7 = "ExchangeDeviceActivity.csv",
	[alias("o8")]
	$OutputFile8 = "ExchangeRecipients.csv",
	$UserName,
	$Password,
	[switch]
	$Office365,
	[switch]
	$Verbose,
	[switch]
	$ServerDetailsOnly,
	[switch]
	$LocalOnly,
	[switch]
	$UseSSL)
	
function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Output "Computer Name:		$($env:COMPUTERNAME)" #-ForegroundColor Magenta
	Write-Output "User Name:			$($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:		$($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:		$($host.Version.Major)"
	Write-Output "PowerShell Version:	$($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size:	$($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:			$($PSVersionTable.CLRVersion)"
	Write-Output "Username Parameter:	$UserName"
	Write-Output "Server Parameter:		$Server"
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

	#Create the Credentials object if username & password have been provided
	if ($UserName -and $Password)
	{
		LogProgress -activity "Exchange Data Export" -Status "Creating Credentials Object" -percentComplete 2
		$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
		$exchangeCreds = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)
	}

	#Connect to exchange server
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

	if (!($LocalOnly))
	{
		if ($exchangeCreds)
		{
			$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Authentication $authenticationType -AllowRedirection -Credential $exchangeCreds
		}
		else
		{
			$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Authentication $authenticationType -AllowRedirection
		}
		
		if ($exchangeSession) {
			LogProgress -activity "Exchange Data Export" -Status "Importing Session" -percentComplete 10	
			Import-PSSession $exchangeSession
		}
	}

	if (!(EnvironmentConfigured))
	{
		#Exchange environment not configured
		#Load Exchange SnapIns
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

		# . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
		# Connect-ExchangeServer -auto
		# . $env:ExchangeInstallPath\bin\Exchange.ps1
	}
	
	if (!(EnvironmentConfigured))
	{
		write-output "Unable to configure Powershell Exchange environment"
		exit
	}   

	#Get the list of Exchange Servers
	LogProgress -activity "Exchange Data Export" -Status "Getting server details" -percentComplete 13   
	Get-ExchangeServer -Identity $ExchangeServer | export-csv $OutputFile1 -notypeinformation -Encoding UTF8

	if (!($ServerDetailsOnly))
	{
		#Get the list of users from Exchange
		LogProgress -activity "Exchange Data Export" -Status "Querying User Data" -percentComplete 20
		$users = Get-User
		$users | export-csv $OutputFile2 -notypeinformation -Encoding UTF8
		
		#Get user activity info from Exchange
		LogProgress -activity "Exchange Data Export" -Status "Querying User Activity Data" -percentComplete 25
		$userActivity = $users | %{Get-LogonStatistics -identity $_.samaccountname}
		$userActivity | export-csv $OutputFile3 -notypeinformation -Encoding UTF8

		#Get Mailbox details
		LogProgress -activity "Exchange Data Export" -Status "Querying Mailboxes" -percentComplete 40
		$mailBoxes = Get-Mailbox
		$mailBoxes | export-csv $OutputFile4 -notypeinformation -Encoding UTF8
		
		#Get Mailbox activity details
		LogProgress -activity "Exchange Data Export" -Status "Querying Mailbox Activity Data" -percentComplete 50
		$mailBoxStatistics = $mailBoxes | %{Get-MailboxStatistics -identity $_.Id}
		$mailBoxStatistics | export-csv $OutputFile5 -notypeinformation -Encoding UTF8
		
		#Get Device details
		LogProgress -activity "Exchange Data Export" -Status "Querying Device Data" -percentComplete 66
		$activeSyncDevices = Get-ActiveSyncDevice
		$activeSyncDevices | export-csv $OutputFile6 -notypeinformation -Encoding UTF8
		
		#Get Device Activity details
		LogProgress -activity "Exchange Data Export" -Status "Querying Device Data" -percentComplete 75
		$activeSyncDeviceStatistics = $activeSyncDevices | %{Get-ActiveSyncDeviceStatistics -identity $_.Id}
		$activeSyncDeviceStatistics | export-csv $OutputFile7 -notypeinformation -Encoding UTF8
		
		#Get Recipients
		LogProgress -activity "Exchange Data Export" -Status "Querying Recipient Data" -percentComplete 90
		$recipients = Get-Recipient
		$recipients | export-csv $OutputFile8 -notypeinformation -Encoding UTF8
	}

	if ($exchangeSession) {
		LogProgress -activity "Exchange Data Export" -Status "Cleaning Session" -percentComplete 95
		Remove-PSSession -Session $exchangeSession}
		
	LogProgress -activity "Exchange Data Export" -complete -Status "Complete"
}

try {
	Get-ExchangeDetails
}
catch {
	LogLastException
}
