#########################################################
#                                                                     
# MsolUserLicenseDetails
# SAM Gold Toolkit
# Original Source: Akshay Chiddarwar (Sam360)
#
#########################################################

 <#
.SYNOPSIS
Retrieves Office365 Licensing information for all the user accounts.

    Files are written to current working directory

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
	[Parameter(Mandatory=$false)]
	[bool]$InstallComponentsIfRequired,
	[alias("o1")]
	$OutputFile = "Office_365_Licenses.csv",
	[alias("log")]
	[string] $LogFile = "O365LogFile.txt"
)

function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Host "Computer Name:            $($env:COMPUTERNAME)"
	Write-Host "User Name:                $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Host "Windows Version:          $($OSDetails.Caption)($($OSDetails.Version))"
	Write-Host "PowerShell Host:          $($host.Version.Major)"
	Write-Host "PowerShell Version:       $($PSVersionTable.PSVersion)"
	Write-Host "PowerShell Word size:     $($([IntPtr]::size) * 8) bit"
	Write-Host "CLR Version:              $($PSVersionTable.CLRVersion)"
	
	$global:OSArch = $OSDetails.OSArchitecture	
}

function LogLastException() {
	$currentException = $Error[0].Exception;

	while ($currentException)
	{
		Write-Host $currentException
		Write-Host $currentException.Data
		Write-Host $currentException.HelpLink
		Write-Host $currentException.HResult
		Write-Host $currentException.Message
		Write-Host $currentException.Source
		Write-Host $currentException.StackTrace
		Write-Host $currentException.TargetSite

		$currentException = $currentException.InnerException
	}
}

function LogProgress([string]$Activity, [string]$Status, [Int32]$PercentComplete, [switch]$Completed ){
	
	Write-Progress -activity $Activity -Status $Status -percentComplete $PercentComplete -Completed:$Completed
	
	if ($Verbose)
	{
		Write-Host ""
		$output = Get-Date -Format HH:mm:ss.ff
		$output += " - "
		$output += $Status
		Write-Host $output
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

function DependencyInstaller {
	
	# Path for MS Online Sign-in Assistant
	if ($OSArch -eq '64-bit') {
		$AzureAD_msi_url = 'http://go.microsoft.com/fwlink/p/?linkid=236297'
		$MsolCLI_msi_url = 'https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi'
	}
	else {
		$AzureAD_msi_url = 'http://go.microsoft.com/fwlink/p/?linkid=236298'
        $MsolCLI_msi_url = 'https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_32.msi'
	}	
	
	# Check if dependency Msol has been installed.
	$MsolInstallCheck = Get-WmiObject -Class Win32_Product | select Name | where { $_.Name -match “Microsoft Online Services Sign-in Assistant”}

	if ($MsolInstallCheck.Name -eq $null) {
		$msifile_msol = $PSScriptRoot + '\msoidcli.msi'

		# Download the required msi
		$webclient_msol = New-Object System.Net.WebClient
		$webclient_msol.DownloadFile($MsolCLI_msi_url, $msifile_msol)
		
		#Check if the file exist in the directory and install on the system
		if (Test-Path $msifile_msol) {

			# Check the msi installer signature and then allow installation	
			if (VerifySignature($msifile_msol)) {
				msiexec /i $msifile_msol /qn | Out-Null
			}
			else {
				$percent = 100
				LogProgress -Activity "MSI Verification" -Status "Msol msi file signature verification failed" -percentComplete $percent
			
				exit
			}			
			
			# Check if dependency Msol has been installed.
			$MsolInstallCheck = Get-WmiObject -Class Win32_Product | select Name | where { $_.Name -match “Microsoft Online Services Sign-in Assistant”}
			  
			if ($MsolInstallCheck.Name -eq $null) {
				$percent = 100
				LogProgress -Activity "Installation Error" -Status "Msol msi could not be installed. Please check if you have admin rights" -percentComplete $percent
				
				exit
			}
        }
		else {
			$percent = 100
			LogProgress -Activity "Dependency Download error" -Status "Could not download Msol msi file in the script path" -percentComplete $percent
			
			Write-Host 'A problem occured while downloading the msi file. If problem persist then please manually install the required msi file downloadable from (http://www.microsoft.com/en-us/download/details.aspx?id=41950)'
			exit
		}
	}	

    # Path for MSOnline Administration Module
	$msifile_azureAD = $PSScriptRoot + '\AdministrationConfig-en.msi'

    $webclient_azureAD = New-Object System.Net.WebClient
	$webclient_azureAD.DownloadFile($AzureAD_msi_url, $msifile_azureAD)

    # Attempt install only if MS online Sign-in Assistant is installed.
    if (Test-Path $msifile_azureAD) {

		# Check the msi installer signature and then allow installation
		if (VerifySignature($msifile_azureAD)) {
			msiexec /i $msifile_azureAD /qn | Out-Null
		}
		else {
			$percent = 100
			LogProgress -Activity "MSI Verification" -Status "AdminConfig msi file signature verification failed" -percentComplete $percent
				
			exit
		}

		# Check if dependency Msol has been installed.
		$ADInstallCheck = Get-WmiObject -Class Win32_Product | select Name | where { $_.Name -match “Windows Azure Active Directory”}
			  
		if ($ADInstallCheck.Name -eq $null) {
			$percent = 100
			LogProgress -Activity "MSI Verification" -Status "AdminConfig msi could not be installed. Please check if you have admin rights" -percentComplete $percent
				
			exit
		}
    }
	else {
		$percent = 100
		LogProgress -Activity "Dependency Download error" -Status "Could not download AdminConfig msi file in the script path" -percentComplete $percent
			
		Write-Host 'A problem occured while downloading the msi file. If problem persist then please manually install the required msi file downloadable from (http://go.microsoft.com/fwlink/p/?linkid=236297)'
		exit
	}    

    Write-Host 'Installation Done'
}

clear 

$global:OSArch = $null
$percent = 0

try {

	LogEnvironmentDetails
	LogProgress -Activity "Office365 Online services licensing information" -Status "Started" -percentComplete $percent

	# Get required modules loaded
	if (Get-Module -ListAvailable -Name 'MSOnline') {
		$percent += 8
		LogProgress -Activity "Required dependency modules" -Status "Already Installed. Continue Execution" -percentComplete $percent
	}
	else {
		if ($InstallComponentsIfRequired) {
			$percent += 2
			LogProgress -Activity "Required dependency modules" -Status "Download and Installation in progress" -percentComplete $percent
			DependencyInstaller
			$percent += 3
			LogProgress -Activity "Required dependency modules are already installed" -Status "Dependency Installation check" -percentComplete $percent
			if (Get-Module -ListAvailable -Name 'MSOnline') {
				Import-Module MSOnline
                Write-Host 'Installation Done'
			}
			else {
				$percent = 100
				LogProgress -Activity "Error occured while installing one of the dependency" -Status "Dependency Installation Error" -percentComplete $percent
				exit
			}		
		}
		else {
			Write-Host 'The required Dependencies are not installed on your machine.'
			Write-Host 'You can download and install the dependencies from'
			Write-Host 'Microsoft Online Services Sign-In Assistant - http://www.microsoft.com/en-us/download/details.aspx?id=41950'
			Write-Host 'Windows Azure Active Directory Module for Windows PowerShell (64-bit version) - http://go.microsoft.com/fwlink/p/?linkid=236297'
			Write-Host 'OR'
			$response = Read-Host 'If you enter (y/Y) the script install the dependencies?'

			if ($response -eq 'y' -or $response -eq 'Y') {
				$percent += 2
				LogProgress -Activity "Required dependency modules" -Status "Download and Installation in progress" -percentComplete $percent
				DependencyInstaller
				$percent += 3
				LogProgress -Activity "Required dependency modules are already installed" -Status "Dependency Installation check" -percentComplete $percent
				if (Get-Module -ListAvailable -Name 'MSOnline') {
					Import-Module MSOnline
				}
				else {
					$percent = 100
					LogProgress -Activity "Error occured while installing one of the dependency" -Status "Dependency Installation Error" -percentComplete $percent
					exit
				}
			}
			else {
				$percent = 100		
				Write-Host 'Please install the dependencies from above mentioned links and re-run the script.'
				LogProgress -Activity "Required dependency module installation" -Status "User opted manual installation" -percentComplete $percent
				exit
			}
		}
	}

	$percent += 1
	LogProgress -Activity "Connect to MS Online Services using username and password if provided or prompt" -Status "Connecting to Office365 Online" -percentComplete 8

	if ($Username -and $Password) {
		$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force

		## Create MS Credential Object
		$credential = New-Object -TypeName System.Management.Automation.PSCredential ($Username, $securePassword)
				
	}
	else {
		## opens the Microsoft Login UI - Powershell dialog
		$credential = Get-Credential
	}

	## Add the account user has entered
	Connect-MsolService -Credential $credential

	$percent += 4
	LogProgress -Activity 'Retrieving Office365 Licensing and User details' -Status 'Get Office365 account details' -percentComplete $percent

	# Get a list of all licences that exist within the tenant
	$OrgLicenseTypes = Get-MsolAccountSku

	# Get list of all the users
	$users = Get-MsolUser -all | where {$_.isLicensed -eq "True"}

	try {
		$test = $users.GetType()
	}
	catch {
		$percent = 100		
		Write-Host 'Please Connect to Office365 account through terminal or when prompted'
		LogProgress -Activity "Login Failed" -Status "User cancelled login operation" -percentComplete $percent
		exit
	}

	$percent_users = (100 - $percent - 5) / $users.count
	$percent_lictype = ($percent_users / $OrgLicenseTypes.Count) / 1.5

	foreach ($user in $users) {
		$percent += $percent_lictype
		LogProgress -Activity 'User Details' -Status ('Collect information for User: ' + $user.displayname) -percentComplete $percent

		# Get specific user details
		$userDetails = $user | Select -Property UserPrincipalName, DisplayName, SignInName, Title, MobilePhone, PhoneNumber, Office, StreetAddress, City, State, PostalCode, WhenCreated, LastPasswordChangeTimestamp, IsBlackberryUser, LicenseReconciliationNeeded, PreferredLanguage, UsageLocation, OverallProvisioningStatus
	
		$altEmail = $user.AlternateEmailAddresses -join '; '
		$userDetails | Add-Member -MemberType NoteProperty -Name "AlternateEmailAddresses" -Value $altEmail
	
		$ProxyAddrs = $user.ProxyAddresses -join '; '
		$userDetails | Add-Member -MemberType NoteProperty -Name "ProxyAddresses" -Value $ProxyAddrs

		# Get list of User licenses
		$userLicenses = $user.Licenses
		$userLicensesSKU = $userLicenses.AccountSkuId
	
		# loop through all licenses subscribed by organisation.
		foreach ($licenseSKU in $OrgLicenseTypes) {		
			$license = $licenseSKU.SkuPartNumber
			$licenseDetails = ''
		
			$percent += $percent_lictype
			LogProgress -Activity 'Licensing information' -Status ('User License: ' + $license) -percentComplete $percent

			# Check if the current license is present in the users license list
			if ($userLicensesSKU.Contains($licenseSKU.AccountSkuId)) {
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
		}
	
		$percent += $percent_lictype
		LogProgress -Activity 'Licensing information' -Status 'Export User details to Csv file' -percentComplete $percent
        
		# Export users details to CSV
		$userDetails | Export-Csv  $OutputFile -append -Encoding UTF8 -NoTypeInformation	
	}

	$percent = 100
	LogProgress -Activity "Script Completed" -Status ('Results available in - ' + $OutputFile) -percentComplete 100
}

catch {		
    LogLastException
	exit
}