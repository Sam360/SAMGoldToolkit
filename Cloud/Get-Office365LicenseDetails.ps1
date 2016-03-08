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
	$OutputFile1 = "Office365Licenses.csv",
	[alias("log")]
	[string] $LogFile = "Office365QueryLog.txt"
)

function LogText {
	param(
		[Parameter(Position=0, ValueFromRemainingArguments=$true, ValueFromPipeline=$true)]
		[Object] $Object,
		[System.ConsoleColor]$color = [System.Console]::ForegroundColor  
	)

	# Display text on screen
	Write-Host -Object $Object -ForegroundColor $color

	if ($LogFile) {
		$Object | Out-File $LogFile -Encoding utf8 -Append 
	}
}

function InitialiseLogFile {
	if ($LogFile -and (Test-Path $LogFile)) {
		Remove-Item $LogFile
	}
}

function LogProgress([string]$Activity, [string]$Status, [Int32]$PercentComplete, [switch]$Completed ){
	
	Write-Progress -activity $Activity -Status $Status -percentComplete $PercentComplete -Completed:$Completed
	
	if ($Verbose){
		LogText ""
	}

	$output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $Status
	LogText $output -Color Green
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
                                                                          
function LogEnvironmentDetails {
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
	LogText -Color Gray "Log File:             $LogFile"
	LogText -Color Gray ""
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

function DependencyInstaller([string]$InstallName, [string]$msiURL, [string]$msiFileName) {
		
	scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
	
	# Check if dependency has been installed.
    $InstallCheck = Get-InstalledApps | where {$_.DisplayName -like $InstallName}

	if ($InstallCheck -eq $null) {
    
		$msifile = $scriptPath + '\' + $msiFileName

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
				$percent = 100
				LogProgress -Activity "MSI Verification" -Status "msi file signature verification failed" -percentComplete $percent
			
				return $false
			}
			
			# Check if dependency has been installed.
            $InstallCheck = Get-InstalledApps | where {$_.DisplayName -like $InstallName}
			
			if ($InstallCheck -eq $null) {
				$percent = 100
				LogProgress -Activity "Installation Error" -Status "msi could not be installed. Please check if you have admin rights" -percentComplete $percent
				
				return $false
			}
        }
		else {
			$percent = 100
			LogProgress -Activity "Dependency Download error" -Status "Could not download msi file($InstallName) in the script path" -percentComplete $percent
			
            LogError 'A problem occured while downloading the msi file. If problem persist then please manually install the required msi file downloadable from ($msiURL)'            
			return $false
		}
	}

    # return true if product is already installed
    # return true if installation succeeds
    return $true
}

clear 

try {
	InitialiseLogFile
	LogEnvironmentDetails
	LogProgress -Activity "Office365 Online services licensing information" -Status "Started" -percentComplete $percent
	
	# Variables
	$percent = 0
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	$OSArch = $OSDetails.OSArchitecture	

	# Get required modules loaded
	if (Get-Module -ListAvailable -Name 'MSOnline') {
        Import-Module MSOnline
		$percent += 8
		LogProgress -Activity "Required dependency modules" -Status "Modules Loaded" -percentComplete $percent
	}
	else {
    	# Path for MS Online Sign-in Assistant
    	if ($OSArch -eq '64-bit') {
    		$AzureAD_msi_url = 'http://go.microsoft.com/fwlink/p/?linkid=236297'
    		$MsolCLI_msi_url = 'https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi'
    	}
    	else {
    		$AzureAD_msi_url = 'http://go.microsoft.com/fwlink/p/?linkid=236298'
            $MsolCLI_msi_url = 'https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_32.msi'
    	}	

		if ($InstallComponentsIfRequired) {
			$percent += 2
			LogProgress -Activity "Required dependency modules" -Status "Download and Installation in progress" -percentComplete $percent
			
            # Install MS Online Sign-in Assistant
            $resInstall = DependencyInstaller -InstallName "Microsoft Online Services Sign-in Assistant" -msiURL $MsolCLI_msi_url -msiFileName "msoidcli.msi"
			if (! $resInstall) {
				$percent = 100
				LogProgress -Activity "Installation Error" -Status "Microsoft Online Services Sign-in Assistant could not be installed. Please check if you have admin rights and try again" -percentComplete $percent
				
				exit
			}
            
            # Install MSOnline Administration Module
            $resInstall = DependencyInstaller -InstallName "Windows Azure Active Directory*" -msiURL $AzureAD_msi_url -msiFileName "AdministrationConfig-en.msi"
            if (! $resInstall) {
				$percent = 100
				LogProgress -Activity "Installation Error" -Status "MSOnline Administration Module could not be installed. Please check if you have admin rights and try again" -percentComplete $percent
				
				exit
			}
            
            $percent += 3
			LogProgress -Activity "Required dependency modules" -Status "Installed. Verification in progress" -percentComplete $percent
			if (Get-Module -ListAvailable -Name 'MSOnline') {
				Import-Module MSOnline
                
                LogText 'Dependency Installed'
			}
			else {
				$percent = 100
				LogProgress -Activity "Error occured while installing one of the dependency" -Status "Dependency Installation Error" -percentComplete $percent
				exit
			}		
		}
		else {
			LogText 'The required Dependencies are not installed on your machine.'
			LogText 'You can download and install the dependencies from'
			LogText 'Microsoft Online Services Sign-In Assistant - http://www.microsoft.com/en-us/download/details.aspx?id=41950'
			LogText 'Windows Azure Active Directory Module for Windows PowerShell (64-bit version) - http://go.microsoft.com/fwlink/p/?linkid=236297'
			LogText 'OR'
			$response = Read-Host 'If you enter (y/Y) SAM360 management point would install the dependencies?'

			if ($response -eq 'y' -or $response -eq 'Y') {
				$percent += 2
				LogProgress -Activity "Required dependency modules" -Status "Download and Installation in progress" -percentComplete $percent

                # Install MS Online Sign-in Assistant
                $resInstall = DependencyInstaller -InstallName "Microsoft Online Services Sign-in Assistant" -msiURL $MsolCLI_msi_url -msiFileName "msoidcli.msi"
    			if (! $resInstall) {
    				$percent = 100
    				LogProgress -Activity "Installation Error" -Status "Microsoft Online Services Sign-in Assistant could not be installed. Please check if you have admin rights and try again" -percentComplete $percent
    				
    				exit
    			}
                
                # Install MSOnline Administration Module
                $resInstall = DependencyInstaller -InstallName "Windows Azure Active Directory*" -msiURL $AzureAD_msi_url -msiFileName "AdministrationConfig-en.msi"
                if (! $resInstall) {
    				$percent = 100
    				LogProgress -Activity "Installation Error" -Status "MSOnline Administration Module could not be installed. Please check if you have admin rights and try again" -percentComplete $percent
    				
    				exit
    			}
            
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
                LogText 'Please install the dependencies from above mentioned links and re-run the script.'
                
				LogProgress -Activity "Required dependency module installation" -Status "User opted manual installation" -percentComplete $percent
				exit
			}
		}
	}

	$percent += 1
	LogProgress -Activity "Office 365 Data Export" -Status "Office 365 Administrator Credentials Required" -percentComplete 8

	# Create the Credentials object if username has been provided
	if(!($UserName -and $Password)){
		$credential = Get-Credential
	}
	else 
	{
		$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
		$credential = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)
	}

	## Add the account user has entered
	Connect-MsolService -Credential $credential -ErrorAction SilentlyContinue -ErrorVariable errConn
    
    if ($errConn -ne $null) {
        LogError "Authentication - Incorrect login details. Please verify the credentials and try again."
                
    	$percent = 100
    	LogProgress -Activity 'Authentication' -Status 'Incorrect login details. Please verify the credentials and try again.' -percentComplete $percent
                
        exit
    }
    
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
		LogError "Please Connect to Office365 account through terminal or when prompted"
                    
		LogProgress -Activity "Login Failed" -Status "User cancelled login operation" -percentComplete $percent
		exit
	}

	$percent_users = (100 - $percent - 5) / $users.count
	$percent_lictype = ($percent_users / $OrgLicenseTypes.Count) / 1.5
    
    # Store results
    $result = @()
    
	foreach ($user in $users) {
		$percent += $percent_lictype
		LogText "Collect information for User: $($user.displayname)"

		# Get specific user details
		$userDetails = $user | Select -Property UserPrincipalName, DisplayName, SignInName, Title, MobilePhone, PhoneNumber, ObjectId, UserType, Department, Office, StreetAddress, City, State, PostalCode, Country, UsageLocation, WhenCreated, LastPasswordChangeTimestamp, PasswordNeverExpires, IsBlackberryUser, LicenseReconciliationNeeded, PreferredLanguage, OverallProvisioningStatus
	
		$altEmail = $user.AlternateEmailAddresses -join '; '
		$userDetails | Add-Member -MemberType NoteProperty -Name "AlternateEmailAddresses" -Value $altEmail
	
		$ProxyAddrs = $user.ProxyAddresses -join '; '
		$userDetails | Add-Member -MemberType NoteProperty -Name "ProxyAddresses" -Value $ProxyAddrs

		# Get list of User licenses
		$userLicenses = $user.Licenses
        
        $userLicensesSKU = @()
    	foreach ($license in $userLicenses) {	
            $userLicensesSKU += $license.AccountSkuId
	    }
        
        if ($userLicensesSKU -ne $null) {
    		# loop through all licenses subscribed by organisation.
    		foreach ($licenseSKU in $OrgLicenseTypes) {		
    			$license = $licenseSKU.SkuPartNumber
    		
    			$percent += $percent_lictype

    			# Check if the current license is present in the users license list
    			if ($userLicensesSKU -contains $licenseSKU.AccountSkuId) {
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
        
		$percent += $percent_lictype
        $result += $userDetails
	}

	# Export users details to CSV
	$result | Export-Csv  $OutputFile1 -Encoding UTF8 -NoTypeInformation
    
	$percent = 100
	LogProgress -Activity "Script Completed" -Status ('Results available in - ' + $OutputFile1) -percentComplete 100
	Start-Sleep -s 3
}

catch {		
    LogLastException
	exit
}