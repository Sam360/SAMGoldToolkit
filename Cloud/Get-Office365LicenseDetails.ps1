#########################################################
#                                                                     
# Get-Office365LicenseDetails.ps1
# SAM Gold Toolkit
#
#########################################################

 <#
.SYNOPSIS
Retrieves Office365 Licensing information for all the user accounts.

.DESCRIPTION
Retrieves licensing information for all the users accounts and outputs this information to a csv file. 
This information includes Username, User details like - Contact and office address, all Products available with company and details of licenses assigned to every user.

.PARAMETER      UserName
Office365 Account Username.

.PARAMETER      Password
Office365 Account Password.

.PARAMETER		InstallComponentsIfRequired
Install the dependencies required for acquiring user information without any user interruption.

.PARAMETER		OutputFile
Output CSV file to store the results. You can specify a specific location to output the csv file.

UserPrincipalName				<user registered post code>
DisplayName						<Name>
SignInName						<signin name>
Title							<user title>
MobilePhone						<user mobile number>
PhoneNumber						<user phone number>
Office							<user office location>
StreetAddress					<user specified address>
City							<user specified city>
State							<user specified state>
PostalCode						<user specified post code>
WhenCreated						<Date & Time>
LastPasswordChangeTimestamp		<Date & Time>
IsBlackberryUser				TRUE | FALSE
LicenseReconciliationNeeded		TRUE | FALSE
PreferredLanguage				en-IE
UsageLocation					IE
OverallProvisioningStatus		Success | PendingActivation | PendingInput
AlternateEmailAddresses			<username>@<company-name>.com
ProxyAddresses					smtp:<username>@<LicenseName>.onmicrosoft.com; SMTP:<username>@<company-name>.com
ProductLicense_1				Success | PendingActivation | PendingInput
ProductLicense_2				Success | PendingActivation | PendingInput
...

#>

Param(
	$Username,
	$Password,
	[switch]
	$InstallComponentsIfRequired,
	[alias("o1")]
	$OutputFile1 = "Office365Licenses.csv",
	[alias("o2")]
	$OutputFile2 = "Office365LicenseAssignments.csv",
	[alias("o3")]
	[string] $LogFile = "Office365QueryLog.txt",
	[switch]
	$Verbose)

function InitialiseLogFile {
	if ($LogFile -and (Test-Path $LogFile)) {
		Remove-Item $LogFile
	}
}

function LogText {
	param(
		[Parameter(Position=0, ValueFromRemainingArguments=$true, ValueFromPipeline=$true)]
		[Object] $Object,
		[System.ConsoleColor]$Color = [System.Console]::ForegroundColor,
		[switch]$NoNewLine = $false  
	)

	# Display text on screen
	Write-Host -Object $Object -ForegroundColor $Color -NoNewline:$NoNewLine

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

function LogProgress([string]$Activity, [string]$Status, [Int32]$PercentComplete, [switch]$Completed ){
	
	Write-Progress -activity $Activity -Status $Status -percentComplete $PercentComplete -Completed:$Completed

	$output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $Status
	LogText $output -Color Green
}

function QueryUser([string]$Message, [string]$Prompt, [switch]$AsSecureString = $false, [string]$DefaultValue){
	if ($Message) {
		LogText $Message -color Yellow
	}

	if ($DefaultValue) {
		$Prompt += " (Default [$DefaultValue])"
	}

	$Prompt += ": "
	LogText $Prompt -color Yellow -NoNewLine
	$strResult = Read-Host -AsSecureString:$AsSecureString

	if(!$strResult) {
		$strResult = $DefaultValue
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

	return new-object Management.Automation.PSCredential $strUsername, $strSecurePassword
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
	LogText -Color Gray " Get-Office365Details.ps1"
	LogText -Color Gray " "

	$OSDetails = Get-WmiObject Win32_OperatingSystem
	LogText -Color Gray "Computer Name:        $($env:COMPUTERNAME)"
	LogText -Color Gray "User Name:            $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	LogText -Color Gray "Windows Version:      $($OSDetails.Caption)($($OSDetails.Version))"
	LogText -Color Gray "PowerShell Host:      $($host.Version.Major)"
	LogText -Color Gray "PowerShell Version:   $($PSVersionTable.PSVersion)"
	LogText -Color Gray "PowerShell Word size: $($([IntPtr]::size) * 8) bit"
	LogText -Color Gray "CLR Version:          $($PSVersionTable.CLRVersion)"
	LogText -Color Gray "Current Date Time:    $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
	LogText -Color Gray "Username Parameter:   $UserName"
	LogText -Color Gray "Output File 1:        $OutputFile1"
	LogText -Color Gray "Output File 2:        $OutputFile2"
	LogText -Color Gray "Log File:             $LogFile"
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

function VerifySignature([string]$msiPath) {
	$sign = Get-AuthenticodeSignature $msiPath

	if ($sign.Status -eq 'Valid') {
		$signDict = ($sign.SignerCertificate.Subject -split ', ') |
         foreach `
             { $signDict = @{} } `
             { $item = $_.Split('='); $signDict[$item[0]] = $item[1] } `
             { $signDict }

		if ($signDict['CN'] -eq 'Microsoft Corporation' -and $signDict['O'] -eq 'Microsoft Corporation' ) {
			return 1
		}
		else {
			return 0
		}
	}
}

function Get-InstalledApps {

    # Create 
    if ([IntPtr]::Size -eq 4) {
        $regpath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    }
    else {
        $regpath = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
    }
    Get-ItemProperty $regpath | .{process{if($_.DisplayName -and $_.UninstallString) { $_ } }} | Select DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString |Sort DisplayName
}

function DependencyInstaller([string]$InstallName, [string]$msiURL, [string]$msiFileName, [string]$downloadPath) {
			
	# Check if dependency has been installed.
    $InstallCheck = Get-InstalledApps | where {$_.DisplayName -like $InstallName}

	if ($InstallCheck -eq $null) {
    
		$msifile = $downloadPath + '\' + $msiFileName

		# Download the required msi
		$webclient = New-Object System.Net.WebClient
		$webclient.DownloadFile($msiURL, $msifile)
		
		# Check if the file exist in the directory and install on the system
		if (Test-Path $msifile) {

			# Check the msi installer signature and then allow installation	
			if (VerifySignature($msifile)) {
				msiexec /i $msifile /qn | Out-Null
			}
			else {
				LogError "MSI verification - msi file signature verification failed"
				return $false
			}
			
			# Check if dependency has been installed.
            $InstallCheck = Get-InstalledApps | where {$_.DisplayName -like $InstallName}
			
			if ($InstallCheck -eq $null) {
				LogError "Installation error - msi could not be installed. Please check if you have admin rights"
				return $false
			}
        }
		else {
			LogError "A problem occured while downloading the msi file. If problem persists, manually install the required msi file ($msiURL)"           
			return $false
		}
	}

    # return true if product is already installed
    # return true if installation succeeds
    return $true
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

function ConfigureOffice365Environment() {
	
	LogProgress -Activity "Office 365 Data Export" -Status "Configuring Environment" -percentComplete 0

	if (Get-Module -ListAvailable -Name 'MSOnline') {
		Import-Module MSOnline
		LogProgress -Activity "Office 365 Data Export" -Status "Modules Loaded" -percentComplete 25

		return $true
	}

	# Required component(s) not installed
	if (! $InstallComponentsIfRequired) {
		LogText 'The required PowerShell Office 365 components are not installed on this device.'
		LogText 'The following components are required'
		LogText ' 1) Microsoft Online Services Sign-In Assistant - http://www.microsoft.com/en-us/download/details.aspx?id=41950'
		LogText ' 2) Windows Azure Active Directory Module for Windows PowerShell (64-bit version) - http://go.microsoft.com/fwlink/p/?linkid=236297'
		LogText ' '
		$response = QueryUser -Prompt "Download and install required components (Y/N)"

		if (!($response -eq 'y' -or $response -eq 'Y')) {
			# User has chosen not to install required component(s)
			LogError 'Required component(s) missing. Script exiting.'
			return $false
		}
	}

	$scriptPath = GetScriptPath
	$scriptFolder = split-path -parent $scriptPath
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	$OSArch = $OSDetails.OSArchitecture	
        
    # Path for MS Online Sign-in Assistant
    if ($OSArch -eq '64-bit') {
    	$AzureAD_msi_url = 'http://go.microsoft.com/fwlink/p/?linkid=236297'
    	$MsolCLI_msi_url = 'https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi'
    }
    else {
    	$AzureAD_msi_url = 'http://go.microsoft.com/fwlink/p/?linkid=236298'
		$MsolCLI_msi_url = 'https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_32.msi'
    }	
	
	# Install MS Online Sign-in Assistant
	LogProgress -Activity "Office 365 Data Export" -Status "Installing Microsoft Online Services Sign-in Assistant" -percentComplete 5
	if (! (DependencyInstaller -InstallName "Microsoft Online Services Sign-in Assistant" `
										-msiURL $MsolCLI_msi_url `
										-msiFileName "msoidcli.msi" `
										-downloadPath $scriptFolder)){
		LogError "Microsoft Online Services Sign-in Assistant could not be installed. Please check if you have admin rights and try again"
		return $false
	}

	# Install MSOnline Administration Module
	LogProgress -Activity "Office 365 Data Export" -Status "Installing Windows Azure Active Directory Component" -percentComplete 15
	if (! (DependencyInstaller -InstallName "Windows Azure Active Directory*" `
										-msiURL $AzureAD_msi_url `
										-msiFileName "AdministrationConfig-en.msi" `
										-downloadPath $scriptFolder)){
		LogError  "MSOnline Administration Module could not be installed. Please check if you have admin rights and try again"
		return $false
	}
                                              
	LogProgress -Activity "Office 365 Data Export" -Status "Loading Module" -percentComplete 20
	if (! (Get-Module -ListAvailable -Name 'MSOnline')) {
		LogError "An error occured while installing a required component. Script exiting."
		return $false
	}

	Import-Module MSOnline
	LogProgress -Activity "Office 365 Data Export" -Status "Modules Loaded" -percentComplete 25

	return $true
}

function ConnectToMsol() {
	LogProgress -Activity "Office 365 Data Export" -Status "Connecting to Office 365 Server" -percentComplete 30

	# Create the Credentials object if username has been provided
	if(!($UserName -and $Password)){
		$credential = Get-ConsoleCredential -DefaultUsername $UserName -Message "Office 365 Administrator Credentials Required"
	}
	else {
		$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
		$credential = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)
	}

	## Add the account user has entered
	Connect-MsolService -Credential $credential -ErrorAction SilentlyContinue -ErrorVariable errConn
    
	if ($errConn -ne $null) {
		LogError "An error occurred connecting to Office 365 online." , $errConn
		return $false
	}

	LogProgress -Activity "Office 365 Data Export" -Status "Connected to Office 365 Server" -percentComplete 40
	return $true
}

function ExportOffice365Data() {

	# Get a list of all licences that exist within the tenant
	LogProgress -Activity 'Office 365 Data Export' -Status 'Querying Available Licenses' -percentComplete 45
	$orgLicenses = Get-MsolAccountSku 
	
	try {
		$test = $orgLicenses.GetType()
	}
	catch {
		LogError "An error occurred when retrieving data from Office 365", "Ensure valid credentials were provided"
		return $false
	}
	
	$orgLicensesInfo = $orgLicenses | Select AccountName, SkuPartNumber, `
		@{Name="SkuName"; Expression={LookupSkuName($_.SkuPartNumber)}}, SkuId, TargetClass, ActiveUnits, `
		ConsumedUnits, LockedOutUnits, SuspendedUnits, WarningUnits
	$orgLicensesInfo | Export-Csv  $OutputFile1 -Encoding UTF8 -NoTypeInformation

    if ($Verbose) {
        LogText ($orgLicensesInfo | Format-Table -property SkuName, SkuPartNumber, ActiveUnits, ConsumedUnits -autosize | Out-String)
    }

	# Get list of all the users
	LogProgress -Activity 'Office 365 Data Export' -Status 'Querying User List' -percentComplete 55
	$users = Get-MsolUser -all | where {$_.isLicensed -eq "True"}

	# Get license assignments
	LogProgress -Activity 'Office 365 Data Export' -Status 'Querying License Assignments' -percentComplete 75
	$result = @()
	foreach ($user in $users) {
		
		# Get specific user details
		$userDetails = $user | Select -Property UserPrincipalName, DisplayName, SignInName, Title, MobilePhone, PhoneNumber, ObjectId, UserType, Department, Office, StreetAddress, City, State, PostalCode, Country, UsageLocation, WhenCreated, LastPasswordChangeTimestamp, PasswordNeverExpires, IsBlackberryUser, LicenseReconciliationNeeded, PreferredLanguage, OverallProvisioningStatus
	
		$altEmail = $user.AlternateEmailAddresses -join '; '
		$userDetails | Add-Member -MemberType NoteProperty -Name "AlternateEmailAddresses" -Value $altEmail
	
		$ProxyAddrs = $user.ProxyAddresses -join '; '
		$userDetails | Add-Member -MemberType NoteProperty -Name "ProxyAddresses" -Value $ProxyAddrs

		# Get list of User licenses
		$userLicenses = $user.Licenses
        
		$userLicensesSkus = @()
		$userLicensesSkuIDs = @()
		$userLicensesSkuNames = @()
    	foreach ($license in $userLicenses) {
			$userLicensesSkus += $license.AccountSku.SkuPartNumber	
			$userLicensesSkuIDs += $license.AccountSkuId
			$userLicensesSkuNames += LookupSkuName($license.AccountSku.SkuPartNumber);
		}

		$userDetails | Add-Member -MemberType NoteProperty -Name "LicenseSkus" -Value ($userLicensesSkus -join ";")
		$userDetails | Add-Member -MemberType NoteProperty -Name "LicenseNames" -Value ($userLicensesSkuNames -join ";")
        
		if ($userLicensesSkuIDs -ne $null) {
    		# loop through all licenses subscribed by organisation.
    		foreach ($orgLicense in $orgLicenses) {		
    			$license = $orgLicense.SkuPartNumber

    			# Check if the current license is present in the users license list
    			if ($userLicensesSkuIDs -contains $orgLicense.AccountSkuId) {
    				# Retrieve specific license from users licenses
    				$userLicense = $userLicenses | Where { $_.AccountSku.SkuPartNumber -eq $license }

    				# Add license title
    				$licenseDetails += '(' + $userLicense.AccountSku.AccountName + ')'			
    			
    				# Loop through all the services in the license.
    				foreach ($serviceStatus in $userLicense.ServiceStatus) {
    					$licenseDetails += $serviceStatus.ServicePlan.ServiceName + ':' + $serviceStatus.ProvisioningStatus + ';'
    				}
    			}

    			# Add the license details to the list
    			$userDetails | Add-Member -MemberType NoteProperty -Name $license -Value $licenseDetails
    			$licenseDetails = ''
    		}
		}
        
		$result += $userDetails
	}

	# Export users details to CSV
	$result | Export-Csv  $OutputFile2 -Encoding UTF8 -NoTypeInformation

	return $true
}

function LookupSkuName ([string] $Sku) {
	# Return a friendly name for the Office 365 SKU (if available)
	switch ($Sku) 
	{
		"AAD_BASIC" {return "Azure Active Directory Basic"}
		"AAD_PREMIUM" {return "Azure Active Directory Premium"}
		"ATP_ENTERPRISE" {return "Exchange Online Advanced Threat Protection"}
		"BI_AZURE_P1" {return "Power BI Reporting and Analytics"}
		"BI_AZURE_P2" {return "Power BI Pro"}
		"CRMINSTANCE" {return "Dynamics CRM Online Additional Production Instance"}
		"CRMIUR" {return "CRM for Partners"}
		"CRMPLAN1" {return "Dynamics CRM Online Essential"}
		"CRMPLAN2" {return "Dynamics CRM Online Basic"}
		"CRMSTANDARD" {return "CRM Online"}
		"CRMSTORAGE" {return "Dynamics CRM Online Additional Storage"}
		"CRMTESTINSTANCE" {return "CRM Test Instance"}
		"DESKLESSPACK" {return "Office 365 (Plan K1)"}
		"DESKLESSPACK_GOV" {return "Office 365 (Plan K1) for Government"}
		"DESKLESSPACK_YAMME" {return "Office 365 (Plan K1) with Yammer"}
		"DESKLESSWOFFPACK" {return "Office 365 (Plan K2)"}
		"DESKLESSWOFFPACK_GOV" {return "Office 365 (Plan K2) for Government"}
		"EMS" {return "Enterprise Mobility Suite"}
		"ENTERPRISEPACK" {return "Office 365 (Plan E3)"}
		"ENTERPRISEPACK_B_PILOT" {return "Office 365 (Enterprise Preview)"}
		"ENTERPRISEPACK_FACULTY" {return "Office 365 (Plan A3) for Faculty"}
		"ENTERPRISEPACK_GOV" {return "Office 365 (Plan G3) for Government"}
		"ENTERPRISEPACK_STUDENT" {return "Office 365 (Plan A3) for Students"}
		"ENTERPRISEPACKLRG" {return "Office 365 (Plan E3)"}
		"ENTERPRISEPACKWSCAL" {return "Office 365 (Plan E4)"}
		"ENTERPRISEPREMIUM_NOPSTNCONF" {return "Office 365 (Plan E5) Without PSTN"}
		"ENTERPRISEWITHSCAL" {return "Office 365 (Plan E4)"}
		"ENTERPRISEWITHSCAL_FACULTY" {return "Office 365 (Plan A4) for Faculty"}
		"ENTERPRISEWITHSCAL_GOV" {return "Office 365 (Plan G4) for Government"}
		"ENTERPRISEWITHSCAL_STUDENT" {return "Office 365 (Plan A4) for Students"}
		"EOP_ENTERPRISE" {return "Exchange Online Protection"}
		"EOP_ENTERPRISE_FACULTY" {return "Exchange Online Protection for Faculty"}
		"EQUIVIO_ANALYTICS" {return "Office 365 Advanced eDiscovery"}
		"ESKLESSWOFFPACK_GOV" {return "Office 365 (Plan K2) for Government"}
		"EXCHANGE_ANALYTICS" {return "Delve Analytics"}
		"EXCHANGE_L_STANDARD" {return "Exchange Online (Plan 1)"}
		"EXCHANGE_S_ARCHIVE_ADDON_GOV" {return "Exchange Online Archiving for Government"}
		"EXCHANGE_S_DESKLESS" {return "Exchange Online Kiosk"}
		"EXCHANGE_S_DESKLESS_GOV" {return "Exchange Kiosk for Government"}
		"EXCHANGE_S_ENTERPRISE_GOV" {return "Exchange (Plan G2) for Government"}
		"EXCHANGE_S_STANDARD" {return "Exchange Online (Plan 2)"}
		"EXCHANGE_S_STANDARD_MIDMARKET" {return "Exchange Online (Plan 1)"}
		"EXCHANGEARCHIVE" {return "Exchange Online Archiving"}
		"EXCHANGEARCHIVE_ADDON" {return "Exchange Online Archiving for Exchange Online"}
		"EXCHANGEDESKLESS" {return "Exchange Online Kiosk"}
		"EXCHANGEENTERPRISE" {return "Exchange Online (Plan 2)"}
		"EXCHANGEENTERPRISE_GOV" {return "Office 365 Exchange Online (Plan 2) for Government"}
		"EXCHANGESTANDARD" {return "Exchange Online (Plan 1)"}
		"EXCHANGESTANDARD_GOV" {return "Office 365 Exchange Online (Plan 1) for Government"}
		"EXCHANGESTANDARD_STUDENT" {return "Exchange Online (Plan 1) for Students"}
		"EXCHANGETELCO" {return "Exchange Online POP"}
		"INTUNE_A" {return "Intune for Office 365"}
		"INTUNE_STORAGE" {return "Intune Extra Storage"}
		"LITEPACK" {return "Office 365 (Plan P1)"}
		"LITEPACK_P2" {return "Office 365 Small Business Premium"}
		"LOCKBOX" {return "Customer Lockbox"}
		"LOCKBOX_ENTERPRISE" {return "Customer Lockbox"}
		"MCOEV" {return "Skype for Business Cloud PBX"}
		"MCOIMP" {return "Skype for Business Online (Plan 1)"}
		"MCOLITE" {return "Skype for Business Online (Plan 1)"}
		"MCOPLUSCAL" {return "Skype for Business Plus CAL"}
		"MCOSTANDARD" {return "Skype for Business Online (Plan 2)"}
		"MCOSTANDARD_GOV" {return "Skype for Business (Plan G2) for Government"}
		"MCOSTANDARD_MIDMARKET" {return "Skype for Business Online (Plan 1)"}
		"MCOVOICECONF" {return "Lync Online (Plan 3)"}
		"MCVOICECONF" {return "Skype for Business Online (Plan 3)"}
		"MFA_PREMIUM" {return "Azure Multi-Factor Authentication"}
		"MIDSIZEPACK" {return "Office 365 Midsize Business"}
		"MS-AZR-0145P" {return "Azure"}
		"NBPOSTS" {return "Social Engagement Additional 10K Posts"}
		"NBPROFESSIONALFORCRM" {return "Social Listening Professional"}
		"O365_BUSINESS" {return "Office 365 Business"}
		"O365_BUSINESS_ESSENTIALS" {return "Office 365 Business Essentials"}
		"O365_BUSINESS_PREMIUM" {return "Office 365 Business Premium"}
		"OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ" {return "Office ProPlus"}
		"OFFICESUBSCRIPTION" {return "Office ProPlus"}
		"OFFICESUBSCRIPTION_GOV" {return "Office ProPlus for Government"}
		"OFFICESUBSCRIPTION_STUDENT" {return "Office ProPlus Student Benefit"}
		"ONEDRIVESTANDARD" {return "OneDrive"}
		"POWER_BI_PRO" {return "Power BI Pro"}
		"POWER_BI_STANDALONE" {return "Power BI for Office 365"}
		"POWER_BI_STANDARD" {return "Power BI Standard"}
		"PROJECT_CLIENT_SUBSCRIPTION" {return "Project Pro for Office 365"}
		"PROJECT_ESSENTIALS" {return "Project Lite"}
		"PROJECTCLIENT" {return "Project Pro for Office 365"}
		"PROJECTESSENTIALS" {return "Project Lite"}
		"PROJECTONLINE_PLAN_1" {return "Project Online (Plan 1)"}
		"PROJECTONLINE_PLAN_2" {return "Project Online (Plan 2)"}
		"PROJECTONLINE_PLAN1_FACULTY" {return "Project Online for Faculty"}
		"PROJECTONLINE_PLAN1_STUDENT" {return "Project Online for Students"}
		"PROJECTWORKMANAGEMENT" {return "Office 365 Planner Preview"}
		"RIGHTSMANAGEMENT" {return "Azure Rights Management"}
		"RMS_S_ENTERPRISE" {return "Azure Active Directory Rights Management"}
		"RMS_S_ENTERPRISE_GOV" {return "Windows Azure Active Directory Rights Management for Government"}
		"SHAREPOINT_PROJECT_EDU" {return "Project Online for Education"}
		"SHAREPOINTDESKLESS" {return "SharePoint Online Kiosk"}
		"SHAREPOINTDESKLESS_GOV" {return "SharePoint Online Kiosk for Government"}
		"SHAREPOINTENTERPRISE" {return "SharePoint Online (Plan 2)"}
		"SHAREPOINTENTERPRISE_EDU" {return "SharePoint (Plan 2) for EDU"}
		"SHAREPOINTENTERPRISE_GOV" {return "SharePoint (Plan G2) for Government"}
		"SHAREPOINTENTERPRISE_MIDMARKET" {return "SharePoint Online (Plan 1)"}
		"SHAREPOINTLITE" {return "SharePoint Online (Plan 1)"}
		"SHAREPOINTPARTNER" {return "SharePoint Online Partner Access"}
		"SHAREPOINTSTANDARD" {return "SharePoint Online (Plan 1)"}
		"SHAREPOINTSTORAGE" {return "SharePoint Online Storage"}
		"SHAREPOINTWAC" {return "Office Online"}
		"SHAREPOINTWAC_EDU" {return "Office Online for Education"}
		"SHAREPOINTWAC_GOV" {return "Office Online for Government"}
		"SQL_IS_SSIM" {return "Power BI Information Services"}
		"STANDARD_B_PILOT" {return "Office 365 (Small Business Preview)"}
		"STANDARDPACK" {return "Office 365 (Plan E1)"}
		"STANDARDPACK_FACULTY" {return "Office 365 (Plan A1) for Faculty"}
		"STANDARDPACK_GOV" {return "Office 365 (Plan G1) for Government"}
		"STANDARDPACK_STUDENT" {return "Office 365 (Plan A1) for Students"}
		"STANDARDWOFFPACK" {return "Office 365 (Plan E2)"}
		"STANDARDWOFFPACK_FACULTY" {return "Office 365 Education E1 for Faculty"}
		"STANDARDWOFFPACK_GOV" {return "Office 365 (Plan G2) for Government"}
		"STANDARDWOFFPACK_IW_FACULTY" {return "Office 365 Education for Faculty"}
		"STANDARDWOFFPACK_IW_STUDENT" {return "Office 365 Education for Students"}
		"STANDARDWOFFPACK_STUDENT" {return "Office 365 (Plan A2) for Students"}
		"STANDARDWOFFPACKPACK_FACULTY" {return "Office 365 (Plan A2) for Faculty"}
		"STANDARDWOFFPACKPACK_STUDENT" {return "Office 365 (Plan A2) for Students"}
		"VISIO_CLIENT_SUBSCRIPTION" {return "Visio Pro for Office 365"}
		"VISIOCLIENT" {return "Visio Pro for Office 365"}
		"WACONEDRIVEENTERPRISE" {return "OneDrive for Business (Plan 2)"}
		"WACONEDRIVESTANDARD" {return "OneDrive Pack"}
		"WACSHAREPOINTENT" {return "Office Web Apps with SharePoint (Plan 2)"}
		"WACSHAREPOINTSTD" {return "Office Web Apps with SharePoint (Plan 1)"}
		"YAMMER_ENTERPRISE" {return "Yammer"}
		"YAMMER_ENTERPRISE_STANDALONE" {return "Yammer Enterprise"}
		"YAMMER_MIDSIZE" {return "Yammer"}
	}

	return $Sku
}

function Get-Office365LicenseDetails() {
	
	try {
		InitialiseLogFile
		LogEnvironmentDetails
        SetupDateFormats
		
		if (-not (ConfigureOffice365Environment)) {
			return;
		}

		if (-not (ConnectToMsol)) {
			return;
		}

		if (-not (ExportOffice365Data)) {
			return;
		}

		LogProgress -Activity "Office 365 Data Export" -Status "Export Complete" -percentComplete 100 -Completed $true
		LogText "Results available in $OutputFile1"
		Start-Sleep -s 3
	}
	catch {		
		LogLastException
	}
} 

Get-Office365LicenseDetails