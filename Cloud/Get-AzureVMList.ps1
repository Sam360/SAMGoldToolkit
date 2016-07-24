 #########################################################
 #                                                                     
 # Get-AzureVMList
 #
 #########################################################

 <#
.SYNOPSIS
Retrieves Azure installation data from a added Microsoft account

.DESCRIPTION
Retrieves installation information from classic and resource manager(ARM) Azure VM for multiple accounts and outputs this information to a csv file. 
This information includes SubscriptionId,SubscriptionName, Environment supported modes, DefaultAccount and more.
The User needs to pass the credentials to execute the script

.PARAMETER      USERNAME
Enter only Organisational(work) or Student Azure Account Username

.PARAMETER       PASSWORD
Azure Account Password

.PARAMETER OUTPUTFILE1
Output CSV file to store the results

.PARAMETER OUTPUTFILE2
Output CSV file to store the results

EXAMPLE DATA WITH 1 DISK

Subscription Name = BizSpark
Subscription ID = 34dd2b84-xxxx-xxxx-xxxx-3fbd0
Default Account = [email]
Environment = AzureCloud
VM Name = [vm-name]
VM IP Address = xxx.xxx.xxx.xxx
DNS Name = http://[vm-name].cloudapp.net/
VM Status = Started
Availability Set Name = xdc
Virtual Network Name = VNETEast
OS = Windows Server 2012 R2 Datacenter, September 2015
VM Image Name = [id]__Windows-Server-2012-R2-201505.01-en.us-127GB.vhd
Total Disk = 1
Total Disk Size in GB = 30
Disk Location = Japan West;


EXAMPLE DATA WITH 2 DISK

Subscription Name = BizSpark
Subscription ID = 95946428-xxxx-xxxx-xxxx-3fbd0
Default Account = [email]
Environment = AzureCloud
VM Name = [vm-name]
VM IP Address = xxx.xxx.xxx.xxx
DNS Name = http://[vm-name].cloudapp.net/
VM Status = Started
Availability Set Name = xdc
Virtual Network Name = 
OS = Ubuntu Server 15.04
VM Image Name = [id]__Ubuntu-15_04-amd64-server-20150910-en-us-30GB
Total Disk = 2
Total Disk Size in GB = 148
Disk Location = East US 2; East US 2

EXAMPLE DATA FOR RM VM

SubscriptionId = 2cb877e2-xxxx-xxxx-xxxx-404b
SubscriptionName = MSDN
Resource Group Name = [resource-group-name]
VM Name = [vm-name]
License Type = 
Location = eastus
Availability Set = [set-name]
Instance Size = Standard_A1
Admin Username = AdminName
VM Provisioning State = Succeeded
Creation Method = FromImage
Publisher = MicrosoftWindowsServer
OS = WindowsServer
VM Image Name = 2012-R2-Datacenter
VM Image Version = latest
VM Private IP Address = xxx.xxx.xxx.xxx
VM Private IP Allocation Method = [Dynamic/Static]

#>

Param(
	$AzureUserName,
	$AzurePassword,
	[alias("o1")]
	$OutputFile1 = "AzureClassicVMList.csv",
	[alias("o2")]
	$OutputFile2 = "AzureRMVMList.csv",
	[alias("log")]
	[string] $LogFile = "AzureLogFile.txt",
	[switch]
	$Verbose
)

function InitialiseLogFile {
	if ($LogFile -and (Test-Path $LogFile)) {
		Remove-Item $LogFile
	}
}

function LogText {
	param(
		[Parameter(Position=0, ValueFromRemainingArguments=$true, ValueFromPipeline=$true)]
		[Object] $Object,
		[System.ConsoleColor]$color = [System.Console]::ForegroundColor,
		[switch]$noNewLine = $false 
	)

	# Display text on screen
	Write-Host -Object $Object -ForegroundColor $color -NoNewline:$noNewLine

	if ($LogFile) {
		$Object | Out-File $LogFile -Encoding utf8 -Append 
	}
}

function LogProgress([string]$Activity, [string]$Status, [Int32]$PercentComplete, [switch]$Completed ){
	
	Write-Progress -activity $Activity -Status $Status -percentComplete $PercentComplete -Completed:$Completed

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
	LogText -Color Gray " "
	LogText -Color Gray "   _____         __  __    _____       _     _   _______          _ _    _ _   "
	LogText -Color Gray "  / ____|  /\   |  \/  |  / ____|     | |   | | |__   __|        | | |  (_) |  "
	LogText -Color Gray " | (___   /  \  | \  / | | |  __  ___ | | __| |    | | ___   ___ | | | ___| |_ "
	LogText -Color Gray "  \___ \ / /\ \ | |\/| | | | |_ |/ _ \| |/ _`` |    | |/ _ \ / _ \| | |/ / | __|"
	LogText -Color Gray "  ____) / ____ \| |  | | | |__| | (_) | | (_| |    | | (_) | (_) | |   <| | |_ "
	LogText -Color Gray " |_____/_/    \_\_|  |_|  \_____|\___/|_|\__,_|    |_|\___/ \___/|_|_|\_\_|\__|"
	LogText -Color Gray " "
	LogText -Color Gray " Get-AzureVMList.ps1"
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
	LogText -Color Gray "Username Parameter:   $AzureUserName"
	LogText -Color Gray "Username Parameter:   $InstallDependency"
	LogText -Color Gray "Output File 1:        $OutputFile1"
	LogText -Color Gray "Output File 2:        $OutputFile2"
	LogText -Color Gray "Log File:             $LogFile"
	LogText -Color Gray ""
}

function Get-CurrentWMFVersion {
    $currVersion_WMF = $PSVersionTable.WSManStackVersion | select Major,Minor
    $currVersionStr_WMF = ([string]$currVersion_WMF.Major + "." + [string]$currVersion_WMF.Minor)
	LogText ("Current WMF Version - " + $currVersionStr_WMF)
	return $currVersionStr_WMF
}

function Get-AzureVMList {	
	InitialiseLogFile
	LogEnvironmentDetails

    $percent = 0	
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	$OSArch = $OSDetails.OSArchitecture	
	
    ## Check if Azure cmdlets are installed.
    ## Install dependencies if they don't exist.
    $AzureModule = Get-Module -ListAvailable -Name Azure
    
    $percent = $percent + 1
	LogProgress -activity "Azure Data Export" -Status "Configuring Environment" -percentComplete $percent
	
    if($AzureModule) {
		## Check in case of existing Azure Installation
		if ($AzureModule.Version.Major -lt 1) {
			## Needs to update the Azure Module.
			$percent = 100
			LogProgress -activity "Azure Data Export" -Status "An old version of Azure PowerShell is detected. Please install the latest version of Azure Powershell from (http://aka.ms/webpi-azps) and re-run the script." -percentComplete $percent
	
			exit
		}
		
		$percent = $percent + 1
			
        Import-Module Azure
    }
    else {
		$WMFVersion = Get-CurrentWMFVersion

		LogText "The required Azure PowerShell components are not installed on this device."
		LogText "The following components are required "
		if ($WMFVersion -lt "3.0") {
			LogText " * Windows Management Framework 3.0 or greater - https://www.microsoft.com/en-us/download/details.aspx?id=34595"
		}		
		LogText " * Microsoft Web Platform Installer with Azure Powershell - http://aka.ms/webpi-azps"
		LogText ""

		$percent = 100
		LogProgress -activity "Azure Data Export" -Status "Please re-run the script once the above mentioned Azure PowerShell dependencies are installed. Script Exiting" -percentComplete $percent
	
		exit
    }

	try {
		##
        ## Get Classic Azure VM details
        ##
        $LoginFlag = $false
        # Clear the logged in user session.
        Clear-AzureProfile -Force

		## Process the user credentials passed through terminal
		if ($AzureUserName -and $AzurePassword) {
			$securePassword = ConvertTo-SecureString $AzurePassword -AsPlainText -Force

			## Convert to Azure account aceptable		
			$cred = New-Object -TypeName System.Management.Automation.PSCredential ($AzureUserName, $securePassword)
			
            $percent = $percent + 2
			LogProgress -activity "Azure Data Export" -Status "Processing credentials from command line" -percentComplete $percent

			## Add the account user has entered
			$AzureAccount = Add-AzureAccount -Credential $cred -ErrorAction SilentlyContinue -ErrorVariable errAddAccount
            if ($errAddAccount.count -eq 0) {
                $LoginFlag = $true
            }
            else {
                LogError -errorDescription "An error occured while trying to login into the azure account. Please check your credentials or try again later."
                Exit
            }
		}
		else {
            $loggedAccount = Get-AzureAccount

		    if ($loggedAccount.count -eq 0) {
                $percent = $percent + 2
			    LogProgress -activity "Azure Data Export" -Status "Azure Classic Credentials Required" -percentComplete $percent
			    $AzureAccount = Add-AzureAccount -ErrorAction SilentlyContinue -ErrorVariable errAddAccount

                if ($errAddAccount.count -eq 0) {
                    $LoginFlag = $true
                }
                else {
                    LogError -errorDescription "An error occured while trying to login into the azure account. Please check your credentials or try again later."
                    Exit
                }
			}            				
		}
	    
        if($LoginFlag) {
		    ## Check if the account is logged in successfully
		    $loggedAccount = Get-AzureAccount

		    if ($loggedAccount.count -eq 0) {
			    LogError "Login failed."
			    Exit
		    }					
		}

		## Get the Subscription List
        $percent = $percent + 2
		LogProgress -activity "Azure Data Export" -Status "Getting Subscription Details" -percentComplete $percent			
		$subscriptions = Get-AzureSubscription -EV AzSubscriptionError -EA Stop

		$subscriptionCount = $subscriptions.count
		$tempValue = [int](40 / $subscriptionCount)
		$subInterval = [int]($tempValue / 3)

        if ($Verbose) {
            LogText "$subscriptionCount subscription(s) found"
            LogText ($subscriptions | Format-Table -property SubscriptionName, SubscriptionID -autosize | Out-String)
        }
		
        $percent = $percent + 1		
		
		## Array for storing csv file details
		$results = @()

		## Loop through each subscription
		foreach ($subscription in $subscriptions) {
	        $percent += $subInterval
	        
	        ## Get Subscription basic details
			$SubscriptionId = $subscription.SubscriptionId
			$SubscriptionName = $subscription.SubscriptionName

		    ## Check if user has access to the subscription
            try {
	            ## Set Default Subscription to get all its details
	            $selectedSub = Select-AzureSubscription -SubscriptionId $SubscriptionId
            }
            catch {
		        LogError -errorDescription "The Loggedin user does not has access for subscription - $SubscriptionName ($SubscriptionId)." -color Red
		        Continue
            }

			## Check if the session is expired
			## Throws an exception if session is expired
			try {
				$percent += [int]$subIntervalActivity
				LogProgress -activity "Azure Data Export" -Status "Querying subscription - $SubscriptionName ($SubscriptionId)" -percentComplete $percent

				$vmList = Get-AzureVM -EV vmListError -EA SilentlyContinue
                if ($vmList -eq $null) {
                    Continue
                }

				$vmListtype = $vmList.GetType()
			}			
			catch {
				## Remove the saved user account credentials, if session has expired
				# Remove-AzureAccount -Name $loggedAccount.Id -Force
				Clear-AzureProfile -Force
				LogError -errorDescription "An error occurred querying VMs for subscription $SubscriptionName ($SubscriptionId). Credentials may have expired."
				Continue
			}

			## Interval decided upon previous completion divided by Number of VM's and
			## further divided by 5 main process
			$vmCount = $vmList.count
	        $vmPercentInterval = [int]($subInterval / $vmCount)

	        foreach ($vm in $vmList) {
	            $percent += $vmPercentInterval  

                if ($Verbose) {
                    LogText "Querying VM $($vm.Name)"
                }              
	            
                $vmEndpoints = Get-AzureEndpoint -VM $vm
                if ($vmEndpoints) {
                    foreach ($ed in $vmEndpoints) {
                        $vmEndpoint += ("" + $ed.Name + ":" + $ed.Port + ":" + $ed.Protocol + ";")
                    }
                }
                
                $vmDetails = $subscription | Select -Property SubscriptionId, SubscriptionName, DefaultAccount, Environment
                              
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Deployment Name" -Value $vm.DeploymentName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Name" -Value $vm.Name
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Label" -Value $vm.Label
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Host Name" -Value $vm.HostName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Service Name" -Value $vm.ServiceName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Availability Set" -Value $vm.AvailabilitySetName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "DNS Name" -Value $vm.DNSName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Instance Name" -Value $vm.InstanceName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Instance Size" -Value $vm.InstanceSize
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Power State" -Value $vm.PowerState
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Status" -Value $vm.Status
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM IP Address" -Value $vm.IpAddress
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Public IP Address" -Value $vm.Location
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Public IP Name" -Value $vm.PublicIPName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Virtual Network Name" -Value $vm.VirtualNetworkName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "OS" -Value $vm.VM.OSVirtualHardDisk.OS
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Image Name" -Value $vm.VM.OSVirtualHardDisk.SourceImageName
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Endpoints" -Value $vmEndpoints
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Location" -Value $vm.Location

		        $results += $vmDetails
			}
		}
        
	    $percent += 1
	    
		$results | Export-Csv -Path $OutputFile1 -NoTypeInformation -Encoding UTF8
		
		# Clear the logged in user session.
		Clear-AzureProfile -Force
        
	    $percent += 1
	    LogProgress -activity "Azure Data Export" -Status "Azure Classic VM List Export Completed" -percentComplete $percent


        ##
        ## Get AzureRM(ARM) VM details
        ##
        
        $LoginFlag = $false

		## Process the user credentials passed through terminal
		if ($AzureUserName -and $AzurePassword) {
			$securePassword = ConvertTo-SecureString $AzurePassword -AsPlainText -Force

			## Convert to Azure account aceptable		
			$cred = New-Object -TypeName System.Management.Automation.PSCredential ($AzureUserName, $securePassword)
				
            $percent = $percent + 2
			LogProgress -activity "Azure Data Export" -Status "Add Credentials from Terminal" -percentComplete $percent
			
            ## Login the account user has entered
			$AzureAccount = Add-AzureRmAccount -Credential $cred -ErrorAction SilentlyContinue -ErrorVariable errAddAccount
            if ($errAddAccount.count -eq 0) {
                $LoginFlag = $true
            }
            else {
                LogError -errorDescription "An error occured while trying to login into the AzureRM account. Please check your credentials or try again later."
                exit
            }
		}
		else {
            $percent = $percent + 2
			LogProgress -activity "Azure Data Export" -Status "AzureRM Credentials Required" -percentComplete $percent
			$AzureAccount = Add-AzureRmAccount -ErrorAction SilentlyContinue -ErrorVariable errAddAccount
            
            if ($errAddAccount.count -eq 0) {
                $LoginFlag = $true
            }
            else {
                LogError -errorDescription "An error occured while trying to login into the AzureRM account. Please check your credentials or try again later."
                exit
            }
		}
	    
        if($LoginFlag -eq $false) {
			LogError "Login failed!"
			Exit		
		}

		## Get the Subscription List
        $percent = $percent + 2
		LogProgress -activity "Azure Data Export" -Status "Credentials accepted, Get ARM Subscription Details" -percentComplete $percent
		$subscriptions = Get-AzureRmSubscription -EV AzSubscriptionError -EA Stop

		$subscriptionCount = $subscriptions.count
		$tempValue = [int](40 / $subscriptionCount)
		$subInterval = [int]($tempValue / 3)
		
        $percent = $percent + 1
		LogProgress -activity "Azure Data Export" -Status "Loop through multiple subscription if any." -percentComplete $percent
        ## Array for storing csv file details
        $results = @()

        ## Loop through each subscription
        foreach ($subscription in $subscriptions) {
	        $percent += $subInterval
	        LogProgress -activity "Azure Data Export" -Status "Start data collection for subscription - $SubscriptionName ($SubscriptionId)" -percentComplete $percent

	        ## Get Subscription basic details
	        $SubscriptionId = $subscription.SubscriptionId
	        $SubscriptionName = $subscription.SubscriptionName

		    ## Check if user has access to the subscription
            try {
	            ## Set Default Subscription to get all its details
	            $selectedSub = Select-AzureRmSubscription -SubscriptionId $SubscriptionId
            }
            catch {
		        LogError -errorDescription "The Loggedin user does not has access for ARM subscription - $SubscriptionName ($SubscriptionId)." -color Red
		        Continue
            }

	        ## Check if the session is expired
	        ## Throws an exception if session is expired
	        try {
				$percent += $subInterval
				LogProgress -activity "Azure Data Export" -Status "Get ARM VM List for Account Validation for subscription - $SubscriptionName ($SubscriptionId)" -percentComplete $percent
				
		        $vmList = Get-AzureRmVM -EV vmListError -EA SilentlyContinue
                if ($vmList -eq $null) {
                    LogText  " The subscription - $SubscriptionName ($SubscriptionId), does not have any ARM VM"
                    Continue
                }

		        $vmListtype = $vmList.GetType()
	        }			
	        catch {
		        ## Remove the saved user account credentials, if session has expired
		        LogError -errorDescription "An error occurred querying VMs for subscription - $SubscriptionName ($SubscriptionId). Credentials may have expired." -color Red
		        Continue
	        }

	        ## Interval decided upon previous completion divided by Number of VM's and
	        ## further divided by 5 main process
	        $vmCount = $vmList.count
	        $vmPercentInterval = [int]($subInterval / $vmCount)

	        foreach ($vm in $vmList) { 
	            $percent += $vmPercentInterval
	            LogProgress -activity "Azure Data Export" -Status ("Start data collection for ARM VM - " + $vm.Name) -percentComplete $percent
                
                $rgn = $vm.ResourceGroupName

                $vmDetails = $subscription | Select -Property SubscriptionId, SubscriptionName

                ## Parse VM OS Profile
                $ospt = $vm.OSProfile
				
                ## Parse VM OS Details
                $spt = $vm.StorageProfile
        
                ## Parse VM Hardware Details
                $hpt = $vm.HardwareProfile
              
                ## Parse VM Network Details
                $vmNIId = ($vmObj.NetworkInterfaceIDs -split "/")[-1]
                $vmNI = Get-AzureRmNetworkInterface -Name $vmNIId -ResourceGroupName $rgn

		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Resource Group Name" -Value $rgn
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Name" -Value $vm.Name
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "License Type" -Value $vm.LicenseType
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Location" -Value $vm.Location
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Availability Set" -Value $vm.AvailabilitySetReference
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Instance Size" -Value $hpt.vmSize
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Admin Username" -Value $ospt.adminUsername
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Provisioning State" -Value $vm.ProvisioningState
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Creation Method" -Value $spt.osDisk.createOption
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $spt.imageReference.publisher
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "OS" -Value $spt.imageReference.offer
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Image Name" -Value $spt.imageReference.sku
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Image Version" -Value $spt.imageReference.version
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Private IP Address" -Value $vmNI.IpConfigurations.PrivateIpAddress
		        $vmDetails | Add-Member -MemberType NoteProperty -Name "VM Private IP Allocation Method" -Value $vmNI.IpConfigurations.PrivateIpAllocationMethod

		        $results += $vmDetails
	        }
            			
        }       
        
	    $percent += 2
	    LogProgress -activity "Azure Data Export" -Status "Start Exporting data to csv file $OutputFile2" -percentComplete $percent
                
		$results | Export-Csv -Path $OutputFile2 -NoTypeInformation -Encoding UTF8
        LogText  "AzureRM VM Export Completed"
         
	    $percent = 100
	    LogProgress -activity "Azure Data Export" -Status "Completed" -percentComplete $percent                
		
	} 
    catch {
		LogLastException

        ## An error occured. Clear the logged in user session.
        Clear-AzureProfile -Force
	}
}

## Call the Get-AzureVMSubscriptionDetails Function to 
##	- Load Account Details
##	- Export CSV

Get-AzureVMList