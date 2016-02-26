 #########################################################
 #                                                                     
 # Get-AzureVMList
 # SAM Gold Toolkit
 # Original Source: Akshay Chiddarwar (Sam360)
 #					Michael Brennan (Sam360)
 #
 #########################################################


 <#
.SYNOPSIS
Retrieves Azure installation data from a added Microsoft account

    Files are written to current working directory

.DESCRIPTION
Retrieves installation information for multiple accounts and outputs this information to a csv file. This information includes SubscriptionId,
SubscriptionName, Environment supported modes, DefaultAccount and more.
The User needs to pass the credentials to execute the script

.PARAMETER      USERNAME
Azure Account Username

.PARAMETER       PASSWORD
Azure Account Password

.PARAMETER OUTPUTFILE
Output CSV file to store the results


EXAMPLE DATA WITH 1 DISK

Subscription Name = BizSpark
Subscription ID = 34dd2b84-xxxx-xxxx-xxxx-3fb6df88a4d0
Default Account = <email>
Environment = AzureCloud
VM Name = <vm-name>
VM IP Address = xxx.xxx.xxx.xxx
DNS Name = http://<vm-name>.cloudapp.net/
VM Status = Started
Availability Set Name = xdc
Virtual Network Name = VNETEast
OS = Windows Server 2012 R2 Datacenter, September 2015
VM Image Name = <id>__Windows-Server-2012-R2-201505.01-en.us-127GB.vhd
Total Disk = 1
Total Disk Size in GB = 30
Disk Location = Japan West;


EXAMPLE DATA WITH 2 DISK

Subscription Name = BizSpark
Subscription ID = 95946428-xxxx-xxxx-xxxx-3fb6df88a4d0
Default Account = <email>
Environment = AzureCloud
VM Name = <vm-name>
VM IP Address = xxx.xxx.xxx.xxx
DNS Name = http://<vm-name>.cloudapp.net/
VM Status = Started
Availability Set Name = xdc
Virtual Network Name = 
OS = Ubuntu Server 15.04
VM Image Name = <id>__Ubuntu-15_04-amd64-server-20150910-en-us-30GB
Total Disk = 2
Total Disk Size in GB = 148
Disk Location = East US 2; East US 2

#>


Param(
	$UserName,
	$Password,
	[alias("o1")]
	$OutputFile = "AzureVMList.csv",
	[alias("log")]
	[string] $LogFile = "AzureLogFile.txt"
)


function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Output "Computer Name:            $($env:COMPUTERNAME)"
	Write-Output "User Name:                $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:          $($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:          $($host.Version.Major)"
	Write-Output "PowerShell Version:       $($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size:     $($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:              $($PSVersionTable.CLRVersion)"
	
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

function LogProgress([string]$Activity, [string]$Status, [Int32]$PercentComplete, [switch]$Completed ) {
	
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

function DependencyInstaller([string]$InstallName, [string]$msiURL, [string]$msiFileName) {
		
	# Check if Azure dependency has been installed.
	$InstallCheck = Get-WmiObject -Class Win32_Product | select Name | where { $_.Name -match $InstallName}

	if ($InstallCheck.Name -eq $null) {
		$msifile = $PSScriptRoot + '\' + $msiFileName

		# Download the required msi
		$webclient = New-Object System.Net.WebClient
		$webclient.DownloadFile($msiURL, $msifile)
		
		#Check if the file exist in the directory and install on the system
		if (Test-Path $msifile) {

			# Check the msi installer signature and then allow installation	
			if (VerifySignature($msifile)) {
				msiexec /i $msifile /qn | Out-Null
			}
			else {
				$percent = 100
				LogProgress -Activity "MSI Verification" -Status "Msol msi file signature verification failed" -percentComplete $percent
			
				return $false
			}
			
			# Check if dependency Msol has been installed.
			$InstallCheck = Get-WmiObject -Class Win32_Product | select Name | where { $_.Name -match $InstallName}
			
			if ($InstallCheck.Name -eq $null) {
				$percent = 100
				LogProgress -Activity "Installation Error" -Status "Msol msi could not be installed. Please check if you have admin rights" -percentComplete $percent
				
				return $false
			}
        }
		else {
			$percent = 100
			LogProgress -Activity "Dependency Download error" -Status "Could not download Msol msi file in the script path" -percentComplete $percent
			
			Write-Host 'A problem occured while downloading the msi file. If problem persist then please manually install the required msi file downloadable from ($msiURL)'
			return $false
		}
	}

    # return true if product is already installed
    # return true if installation succeeds
    return $true
}

function InstallUsingWebPI ([string]$InstallName) {
    [reflection.assembly]::LoadWithPartialName("Microsoft.Web.PlatformInstaller") | Out-Null
 
    $ProductManager = New-Object Microsoft.Web.PlatformInstaller.ProductManager
    $ProductManager.Load()
    $product = $ProductManager.Products | Where { $_.ProductId -eq "WindowsAzurePowershell" }
 
    $InstallManager = New-Object Microsoft.Web.PlatformInstaller.InstallManager
 
    $Language = $ProductManager.GetLanguage("en")
    $installertouse = $product.GetInstaller($Language)
 
    $installer = New-Object 'System.Collections.Generic.List[Microsoft.Web.PlatformInstaller.Installer]'
    $installer.Add($installertouse)
    $InstallManager.Load($installer)
 
    $failureReason=$null
    foreach ($installerContext in $InstallManager.InstallerContexts) {
        $InstallManager.DownloadInstallerFile($installerContext, [ref]$failureReason)
    }

    $installstatus = $InstallManager.StartInstallation()

    $installProgress = $InstallManager.InstallerContexts.InstallationState

    Write-Host ("B4 while: " + $installProgress)
    while($installProgress -ne "InstallCompleted") {
        $installProgress = $InstallManager.InstallerContexts.InstallationState
        Write-Host ("Just Enter: " + $installProgress)
        #Check if installation is still in progress or it has come across any error
        if ( ($installProgress -eq "Installing") -or ($installProgress -eq "Downloaded") -or ($installProgress -eq "Downloading") ) {
            Write-Host ("B4 Sleep: " + $InstallManager.InstallerContexts.InstallationState)
            # Wait for 10 seconds
            Start-Sleep -s 5
            $installProgress = $InstallManager.InstallerContexts.InstallationState

            Write-Host ("After Sleep: " + $InstallManager.InstallerContexts.InstallationState)
        }
        else {
            break
            return $false
        }
    }

    $installstatus
    Write-Host ("Exit While: " + $InstallManager.InstallerContexts.InstallationState)
    return $true

}

function Get-AzureVMList{	
	LogEnvironmentDetails
	LogProgress -Activity "Azure VM List Export" -Status "Logging environment details" -percentComplete 2

    ## Check if Azure cmdlets are installed.
    ## Install dependencies if they don't exist.
    $AzureModule = Get-Module -Name Azure
    if($AzureModule) {
        Import-Module Azure
    }
    else {
        if($OSArch -eq "64-bit") {
            $msi_url_wpi = "http://download.microsoft.com/download/C/F/F/CFF3A0B8-99D4-41A2-AE1A-496C08BEB904/WebPlatformInstaller_amd64_en-US.msi"
        }
        else {
            $msi_url_wpi = "http://download.microsoft.com/download/C/F/F/CFF3A0B8-99D4-41A2-AE1A-496C08BEB904/WebPlatformInstaller_x86_en-US.msi"
        }
        $exec = DependencyInstaller -InstallName "Microsoft Web Platform Installer*" -msiURL $msi_url_wpi -msiFileName "wpilauncher.exe"

        if ($exec) {
            $installStatus = InstallUsingWebPI -InstallName "WindowsAzurePowerShell"
            if ($installStatus) {
                try {
                    .$profile.AllUsersAllHosts
                    $sptPath = $PSScriptRoot
                    powershell -noexit "& ""$sptPath\Get-AzureVMList.ps1"""
                    Import-Module Azure                    
                }
                catch {
		            LogLastException
		            LogProgress -activity "Azure VM List Export" -Status "Problem in loading azure dependencies." -percentComplete 100
                    exit
                }
            }
            else {
		        LogProgress -activity "Azure VM List Export" -Status "Problem in installing azure dependencies." -percentComplete 100
                exit
            }
        }
        else {
		    LogProgress -activity "Azure VM List Export" -Status "Problem in installing dependencies." -percentComplete 100
            exit
        }
    }

    #Clear-AzureProfile -Force
    $LoginFlag = $false

	Try {
		
		## Process the user credentials passed through terminal
		if ($UserName -and $Password) {
			$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force

			## Convert to Azure account aceptable		
			$cred = New-Object -TypeName System.Management.Automation.PSCredential ($UserName, $securePassword)
				
			LogProgress -activity "Azure VM List Export" -Status "Add Credentials from Terminal" -percentComplete 5
			## Add the account user has entered
			$AzureAccount = Add-AzureAccount -Credential $cred
            $LoginFlag = $true
		}
		else {
            $loggedAccount = Get-AzureAccount

		    if ($loggedAccount.count -eq 0) {
			    LogProgress -activity "Azure VM List Export" -Status "Request Credentials from User" -percentComplete 5
			    $AzureAccount = Add-AzureAccount
                $LoginFlag = $true            
			}            				
		}
	    
        if($LoginFlag) {
		    LogProgress -activity "Azure VM List Export" -Status "Checking if login was successful" -percentComplete 6
		    ## Check if the account is logged in successfully
		    $loggedAccount = Get-AzureAccount

		    if ($loggedAccount.count -eq 0) {
			    Write-Output "Login failed."
			    Exit
		    }					
		}

		LogProgress -activity "Azure VM List Export" -Status "Credentials accepted, Continue" -percentComplete 8
			
		## Get the Subscription List
		LogProgress -activity "Azure VM List Export" -Status "Get Subscription Details" -percentComplete 10
		$subscriptions = Get-AzureSubscription -EV AzSubscriptionError -EA Stop

		$subscriptionCount = $subscriptions.count
		$tempValue = [int](85 / $subscriptionCount) / 5
		$subIntervalActivity = $tempValue * 2
		$subInterval = $tempValue * 2
		$percent = 11
		
        #$ListAzureVMImages = Get-AzureVMImage
        
		LogProgress -activity "Azure VM List Export" -Status "Loop through multiple subscription if any." -percentComplete $percent

		## Array for storing csv file details
		$results = @()

		## Loop through each subscription
		foreach ($subsciption in $subscriptions) {

			## Get the Subscription Basic
			$SubscriptionId = $subsciption.SubscriptionId
			$SubscriptionName = $subsciption.SubscriptionName
			$DefaultAccount = $subsciption.DefaultAccount
			$Environment = $subsciption.Environment

			## Set Default Subscription to get all its details
			Select-AzureSubscription -SubscriptionId $SubscriptionId

			## Check if the session is expired
			## Throws an exception if session is expired
			try {			
				## $ADUser = Get-AzureADUser
				$percent += [int]$subIntervalActivity
				LogProgress -activity "Azure VM List Collection" -Status "Get VM List for Account Validation for Subscription: $SubscriptionName ($SubscriptionId)" -percentComplete $percent
				$vmList = Get-AzureVM -EV vmListError -EA SilentlyContinue
				$vmListtype = $vmList.GetType()
			}			
			catch {
				## Remove the saved user account credentials, if session has expired
				#Remove-AzureAccount -Name $loggedAccount.Id -Force
				Clear-AzureProfile -Force

				Write-Output "====> Your Azure credentials might have not setup or expired. Continuing with another subscription if available."
				LogProgress -activity "Azure VM List Collection **Error**" -Status "Session Expired. Exit" -percentComplete $percent
				Continue
			}

			## Interval decided upon previous completion divided by Number of VM's and
			## further divided by 5 main process
			$vmCount = $vmList.count
			$completePercentInterval = ($subInterval / $vmCount) / 5

			foreach ($vm in $vmList) {
				$vmLocation = $vm.Location

                $vmEndpoints = Get-AzureEndpoint -VM $vm
                if ($vmEndpoints) {
                    foreach ($ed in $vmEndpoints) {
                        $vmEndpoint += ("" + $ed.Name + ":" + $ed.Port + ":" + $ed.Protocol + ";")
                    }
                }
                
				$details = @{
					"Subscription Name" = $SubscriptionName
					"Subscription ID" = $SubscriptionId
					"Default Account" = $DefaultAccount
					Environment = $Environment
				    "Deployment Name" = $vm.DeploymentName
					"VM Name" = $vm.Name
                    "Label" = $vm.Label
                    "Host Name" = $vm.HostName
                    "Service Name" = $vm.ServiceName
                    "Availability Set" = $vm.AvailabilitySetName
					"DNS Name" = $vm.DNSName
				    "Instance Name" = $vm.InstanceName
				    "Instance Size" = $vm.InstanceSize
                    "Power State" = $vm.PowerState
					"VM Status" = $vm.Status
					"VM IP Address" = $vm.IpAddress
                    "Public IP Address" = $vm.PublicIPAddress
                    "Public IP Name" = $vm.PublicIPName
					"Virtual Network Name" = $vm.VirtualNetworkName
					OS = $vm.VM.OSVirtualHardDisk.OS
					"VM Image Name" = $vm.VM.OSVirtualHardDisk.SourceImageName
				}
				$results += New-Object PSObject -Property $details
			}
			
			$percent += [int]$subIntervalActivity
			LogProgress -activity "Azure VM List Export" -Status "Start CSV Export for Subscription - $SubscriptionId" -percentComplete $percent
		}
        
		$results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        Write-Host "Export Completed"

		LogProgress -activity "Azure VM List Export" -Status "CSV Export Complete" -percentComplete 100

	} catch{
		LogLastException

        ## Some error occured. Clear the logged in user session.
        Clear-AzureProfile -Force
	}
}


# Call the Get-AzureVMSubscriptionDetails Function to 
#	- Load Account Details
#	- Export CSV

Get-AzureVMList
