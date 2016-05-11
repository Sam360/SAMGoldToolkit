##########################################################################
# 
# Get-SharePointLicenseDetails
# SAM Gold Toolkit
#
##########################################################################
 
 Param(
	$Username = "$($env:USERDOMAIN)\$($env:USERNAME)",
	$Password,
	[alias("server")]
    $SharePointServer,
	[switch]
	$Headless = $false,
    [alias("i1")]
    $InputADGroupsFile = "ADGroups.csv",
    [alias("i2")]
    $InputADUsersFile = "ADUsers.csv",
    [alias("o1")]
    $OutputFile1 = "SharePointSites.csv",
	[alias("o2")]
    $OutputFile2 = "SharePointUserGroups.csv",
	[alias("o3")]
    $OutputFile3 = "SharePointUserCALs.csv",
	[alias("log")]
	[string] $LogFile = "SPLogFile.txt"
)

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

function InitialiseLogFile {
	if ($LogFile -and (Test-Path $LogFile)) {
		Remove-Item $LogFile
	}
}

function LogProgress($progressDescription){
	if ($Verbose){
		LogText ""
	}

	$output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
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
	LogText -Color Gray " Get-SharePointLicenseDetails.ps1"
	LogText -Color Gray " "

	$OSDetails = Get-WmiObject Win32_OperatingSystem
	$ScriptPath = GetScriptPath
	LogText -Color Gray "Computer Name:        $($env:COMPUTERNAME)"
	LogText -Color Gray "User Name:            $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	LogText -Color Gray "Windows Version:      $($OSDetails.Caption)($($OSDetails.Version))"
	LogText -Color Gray "PowerShell Host:      $($host.Version.Major)"
	LogText -Color Gray "PowerShell Version:   $($PSVersionTable.PSVersion)"
	LogText -Color Gray "PowerShell Word size: $($([IntPtr]::size) * 8) bit"
	LogText -Color Gray "CLR Version:          $($PSVersionTable.CLRVersion)"
	LogText -Color Gray "Current Date Time:    $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
	LogText -Color Gray "Script Path:          $ScriptPath"
	LogText -Color Gray "Server Parameter:     $SharePointServer"
	LogText -Color Gray "Output File 1:        $OutputFile1"
    LogText -Color Gray "Output File 2:        $OutputFile2"
    LogText -Color Gray "Output File 3:        $OutputFile3"
    LogText -Color Gray "AD Groups File:       $InputADGroupsFile"
    LogText -Color Gray "AD Users File:        $InputADUsersFile"
	LogText -Color Gray "Log File:             $LogFile"
	LogText -Color Gray ""
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

function EnvironmentConfigured {
	if (Get-Command "Get-SPSite" -errorAction SilentlyContinue){
		return $true
    }
	else {
		return $false
    }
}

function IsLocalComputer {
    param([string]$ComputerName)
    
	if (!$ComputerName) {
		return $true}
	
	if ($env:computerName -eq $ComputerName) {
		return $true}
	
	if ($env:computerName -eq $ComputerName.Split(".")[0]) {
		return $true}
	
	return $false
}

function PutFile {
	param([string]$LocalFilePath,
		[string]$RemoteFilePath,
		[System.Management.Automation.Runspaces.PSSession]$RemoteSession)
	
	LogText "Copying file $LocalFilePath to $SharePointServer"
	try {
		$fileData = Get-Content -Path $LocalFilePath -ErrorAction SilentlyContinue}
	catch {}
	Invoke-Command -Session $RemoteSession -ScriptBlock { 
		param($fileData, $FilePath)
		Set-Content -Path $FilePath -Value $fileData} -ArgumentList $fileData, $RemoteFilePath
}

function GetFile {
	param([string]$RemoteFilePath,
		[string]$LocalFilePath,
		[System.Management.Automation.Runspaces.PSSession]$RemoteSession)
	
	LogText "Copying file $RemoteFilePath from $SharePointServer"
	try {
		$fileData = Invoke-Command -Session $RemoteSession -ScriptBlock { 
			param($FilePath)
			Get-Content -Path $FilePath} -ArgumentList $RemoteFilePath
	}
	catch {
		LogError "Unable to retrieve file ($RemoteFilePath) from device ($SharePointServer)"
	}
	
	Set-Content -Path $LocalFilePath -Value $fileData
}

function DeleteRemoteFile {
	param([string]$RemoteFilePath,
		[System.Management.Automation.Runspaces.PSSession]$RemoteSession)
	
	LogText "Deleting file $RemoteFilePath from $SharePointServer"
	try {
		$fileData = Invoke-Command -Session $RemoteSession -ScriptBlock { 
			param($FilePath)
			if ($FilePath -and (Test-Path $FilePath)) {
				Remove-Item $FilePath
			}} -ArgumentList $RemoteFilePath
	}
	catch {
		LogError "Unable to delete file ($RemoteFilePath) from device ($SharePointServer)"
	}
}

function GetRemoteDefaultFolder {
	param([System.Management.Automation.Runspaces.PSSession]$RemoteSession)
	
	try {
		return Invoke-Command -Session $RemoteSession -ScriptBlock { 
			(Get-Location).Path
		}
	}
	catch {
		LogError "Unable to query default folder on device ($SharePointServer)"
	}
}

function GetScriptPath
{
	if($PSCommandPath){
		return $PSCommandPath; }
		
	if($MyInvocation.ScriptName){
		return $MyInvocation.ScriptName }
		
	if($script:MyInvocation.MyCommand.Path){
		return $script:MyInvocation.MyCommand.Path }

	$ScriptDir = (Get-Location).Path
    return "$ScriptDir\Get-SharePointLicenseDetails.ps1"
}

function Set-SPUserCAL($bIsPremium, $member, $htAllUsers) {
    if ($bIsPremium) {
        $htAllUsers.Set_Item($member.Key, "Premium")
    }
    else {
        if ($member.Value -ne "Premium") {
            $htAllUsers.Set_Item($member.Key, "Standard")
        }
    }
}

function Set-ADGroupMembersCAL($ADGroupName, $htADGroups, $bIsPremium, $htAllUsers) {
    ## This function is to set the AD Group members CAL which is given permission to access the site

    ## Get the Group members
    $grpMembersAD = $htADGroups.GetEnumerator() | where {$_.Key -eq $ADGroupName} | select value
    
    foreach ($grpMemberAD in $grpMembersAD) {
        ## Get member FQDN
        $member = $htAllUsers.GetEnumerator() | where {$_.Key -eq $grpMemberAD}
        Write-Host ("Ad grp mem as site user" + $grpMemberAD)
        Set-SPUserCAL -bIsPremium $bIsPremium -member $member -htAllUsers $htAllUsers
    }
}

function Get-SharePointLicenseDetails {
	try {
		InitialiseLogFile
		LogEnvironmentDetails

		if (!$SharePointServer) {
			# Target server was not specified on the command line. Query user.
			$SharePointServer = QueryUser -Prompt "SharePoint Server" -DefaultValue "$($env:computerName)"
		}
		
		$result = @()
		
		if (IsLocalComputer($SharePointServer)) {
			$result = QuerySharePointInfo
		}
		else{
			# Execuing SharePoint script over a remote session requires CredSSP authentication option
			# to avoid double hop authentication issues. This is generally blocked by Group Policy.
			# We use PSExec and user credentials to avoid this issue.
			
			# Ensure PSExec is available
			$ScriptPath = GetScriptPath
			$ScriptFolder = split-path -parent $ScriptPath
			$PSExecPath = "$ScriptFolder\PSExec.exe"
			if (!(Test-Path $PSExecPath)) {
				LogError "PSExec.exe not found in folder $ScriptDir",
					"PSexec is required to query remote SharePoint server"
                return
			}
			
			# Credentials are required to avoid double hop authentication problem
			if(!($UserName -and $Password)){
				$Creds = Get-ConsoleCredential -Message "SharePoint Server Credentials Required (Username Format [Domain\Username])" -DefaultUsername $UserName
				if ($Creds) {
					$UserName = $Creds.UserName
					$Password = $Creds.Password
				}
			}
			
			LogProgress "Connecting to remote SharePoint server ($SharePointServer)"
			if(!($UserName -and $Password)){
				LogError "Warning: SharePoint script generally requires username and password to be specified.", 
					"Script will continue, but may fail due to insufficient privileges."
				$session = New-PSSession -ComputerName $SharePointServer -ErrorAction SilentlyContinue -ErrorVariable strConnectionError
			}
			else {
				$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        		$PSCreds = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)

				$session = New-PSSession -ComputerName $SharePointServer –Credential $PSCreds -ErrorAction SilentlyContinue -ErrorVariable strConnectionError
			}

			
			if(-not($session)) {
                LogError "Unable to connect to server $SharePointServer", $strConnectionError
                return
            }
			
			# Move required files to remote computer
			LogProgress "Moving required files to $SharePointServer"
			# Move script file
			$LocalScriptPath = GetScriptPath
			$RemoteDefaultFolder = GetRemoteDefaultFolder -RemoteSession $session
			$RemoteScriptPath = "$RemoteDefaultFolder\Get-SharePointLicenseDetails.ps1"
			PutFile -LocalFilePath $LocalScriptPath -RemoteFilePath $RemoteScriptPath -RemoteSession $session
			
			# Move AD files (If they exist)
			PutFile -LocalFilePath $InputADGroupsFile -RemoteFilePath "ADGroups.csv" -RemoteSession $session
			PutFile -LocalFilePath $InputADUsersFile -RemoteFilePath "ADUsers.csv" -RemoteSession $session
			
			# Reset remote log file
			PutFile -LocalFilePath "" -RemoteFilePath "SPLogFile.txt" -RemoteSession $session
			
			# Execute the script remotely
			LogProgress "Executing script remotely using PSExec"
			if($UserName -and $Password){
				Start-Process $PSExecPath -ArgumentList "\\$SharePointServer -h -accepteula -w ""$RemoteDefaultFolder"" -u $UserName -p $Password powershell.exe -ExecutionPolicy Bypass -File ""$RemoteScriptPath"" -server $SharePointServer -headless" -Wait -NoNewWindow -WorkingDirectory $RemoteDefaultFolder -ErrorVariable errPS
			}
			else {
				Start-Process $PSExecPath -ArgumentList "\\$SharePointServer -h -accepteula -w ""$RemoteDefaultFolder"" powershell.exe -ExecutionPolicy Bypass -File ""$RemoteScriptPath"" -server $SharePointServer -headless" -Wait -NoNewWindow -WorkingDirectory $RemoteDefaultFolder -ErrorVariable errPS
			}

			# Copy the results back to this device
			LogProgress "Collecting output files from $SharePointServer"
			GetFile -RemoteFilePath "SharePointSites.csv" -LocalFilePath $OutputFile1 -RemoteSession $session
			GetFile -RemoteFilePath "SharePointUserGroups.csv" -LocalFilePath $OutputFile2 -RemoteSession $session
			GetFile -RemoteFilePath "SharePointUserCALs.csv" -LocalFilePath $OutputFile3 -RemoteSession $session
			GetFile -RemoteFilePath "SPLogFile.txt" -LocalFilePath "RemoteSPLogFile.txt" -RemoteSession $session
			
			Get-Content "RemoteSPLogFile.txt" > $LogFile
			if ($errPS) {
				LogText $errPS
			}
			
			# Clean up the remote device
			LogProgress "Cleaning environment on $SharePointServer"
			DeleteRemoteFile -RemoteSession $session -RemoteFilePath "SharePointSites.csv"
			DeleteRemoteFile -RemoteSession $session -RemoteFilePath "SharePointUserGroups.csv"
			DeleteRemoteFile -RemoteSession $session -RemoteFilePath "SharePointUserCALs.csv"
			DeleteRemoteFile -RemoteSession $session -RemoteFilePath "SPLogFile.txt"
			DeleteRemoteFile -RemoteSession $session -RemoteFilePath "ADGroups.csv"
			DeleteRemoteFile -RemoteSession $session -RemoteFilePath "ADUsers.csv"
			DeleteRemoteFile -RemoteSession $session -RemoteFilePath $RemoteScriptPath
		}
	}
	catch {
		LogError "An exception occurred querying SharePoint server"
        LogLastException
	}
	
	# Kill session
    if ($session) {
        Remove-PsSession $session
    }
}

function QuerySharePointInfo {
	try {
        if (!(EnvironmentConfigured)) {
            Add-PSSnapin Microsoft.Sharepoint.Powershell
		    if (!(EnvironmentConfigured)) {
			    LogError "SharePoint Environment could not be configured. Please enter a SharePoint server name and try again"
			    return $false
		    }
		}

        ## Define variables for user groups export
        $arrUserGroupDetails = @()
        $htResultSPUserGroups = @{}
        $arrSPADGroups = @()
        $arrResultUserGroups = @()
        
        $resultUserCAL = @()
        $result = @()
        
        $sites = Get-SPWebApplication | Get-SPSite -Limit All | Get-SPWeb -Limit All | select `
                    @{Name="WebApp"; Expression = {$_.Site.WebApplication}}, `
                    @{Name="SiteDet"; Expression = {$_.Site}} -Unique
       
        foreach ($site in $sites) {
            
            $waFeatureNames = ""
            $scFeatureNames = ""

            $waURL = $site.WebApp.Url
            $scURL = $site.SiteDet.Url
            $siteWebs = $site.SiteDet.AllWebs
            
            foreach ($siteWeb in $siteWebs) {                
                
                $details = New-Object PSObject

                $userGroups = ""
                $userLogins = ""

                $siteURL = $siteWeb.URL
                LogText ("Querying details for site - " + $siteURL)

                ##
                ## Web Application Details
                ##
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationUrl" -Value $waURL
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationName" -Value $site.WebApp.Name
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationDisplayName" -Value $site.WebApp.DisplayName
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationId" -Value $site.WebApp.Id
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationFarmName" -Value $site.WebApp.Farm
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationStatus" -Value $site.WebApp.Status
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationVersion" -Value $site.WebApp.Version
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationApplicationPoolName" -Value $site.WebApp.ApplicationPool

                ##
                ## Site Collection Details
                ##
                $details | Add-Member -MemberType NoteProperty -Name "SCUrl" -Value $scURL
                $details | Add-Member -MemberType NoteProperty -Name "SCHostName" -Value $site.SiteDet.HostName
                $details | Add-Member -MemberType NoteProperty -Name "SCWebApplication" -Value $site.SiteDet.WebApplication
                $details | Add-Member -MemberType NoteProperty -Name "SCID" -Value $site.SiteDet.ID
                $details | Add-Member -MemberType NoteProperty -Name "SCSchemaVersion" -Value $site.SiteDet.SchemaVersion
                $details | Add-Member -MemberType NoteProperty -Name "SCArchived" -Value $site.SiteDet.Archived
                $details | Add-Member -MemberType NoteProperty -Name "SCCreation Date" -Value $site.SiteDet.CertificationDate
                $details | Add-Member -MemberType NoteProperty -Name "SCExpirationDate" -Value $site.SiteDet.ExpirationDate
                $details | Add-Member -MemberType NoteProperty -Name "SCLastContentModifiedDate" -Value $site.SiteDet.LastContentModifiedDate
                $details | Add-Member -MemberType NoteProperty -Name "SCLastSecurityModifiedDate" -Value $site.SiteDet.LastSecurityModifiedDate
            
                ##
                ## Site Details
                ##
                $details | Add-Member -MemberType NoteProperty -Name "SiteUrl" -Value $siteURL
                $details | Add-Member -MemberType NoteProperty -Name "SiteTitle" -Value $siteWeb.Title
                $details | Add-Member -MemberType NoteProperty -Name "SiteName" -Value $siteWeb.Name
                $details | Add-Member -MemberType NoteProperty -Name "SiteID" -Value $siteWeb.ID
                $details | Add-Member -MemberType NoteProperty -Name "SiteDescription" -Value $siteWeb.Description
                $details | Add-Member -MemberType NoteProperty -Name "SiteAuthor" -Value $siteWeb.Author
                $details | Add-Member -MemberType NoteProperty -Name "SiteParentWeb" -Value $siteWeb.ParentWeb
                $details | Add-Member -MemberType NoteProperty -Name "SiteParentWebId" -Value $siteWeb.ParentWebId
                $details | Add-Member -MemberType NoteProperty -Name "SiteIsAppWeb" -Value $siteWeb.IsAppWeb
                $details | Add-Member -MemberType NoteProperty -Name "SiteIsRootWeb" -Value $siteWeb.IsRootWeb
                $details | Add-Member -MemberType NoteProperty -Name "SiteHasUniqueRoleDefinitions" -Value $siteWeb.HasUniqueRoleDefinitions
                $details | Add-Member -MemberType NoteProperty -Name "SiteAllowAnonymousAccess" -Value $siteWeb.AllowAnonymousAccess
                $details | Add-Member -MemberType NoteProperty -Name "SiteWebTemplate" -Value $siteWeb.WebTemplate
                $details | Add-Member -MemberType NoteProperty -Name "SiteUIVersion" -Value $siteWeb.UIVersion
                $details | Add-Member -MemberType NoteProperty -Name "SiteCreation Date" -Value $siteWeb.Created
                $details | Add-Member -MemberType NoteProperty -Name "SiteLastItemModifiedDate" -Value $siteWeb.LastItemModifiedDate
                

                $bWAPremium = $false
                $bSCPremium = $false
                $bSitePremium = $false

                ## Collect Web Application premium features    
                $waPremiumFeatures = Get-SPFeature "PremiumWebApplication" -WebApplication $waURL -ErrorAction SilentlyContinue -ErrorVariable errGetSPFeature
                foreach ($wapremiumfeature in $waPremiumFeatures) {
                    $wapremiumfeatureids = $wapremiumfeature.ActivationDependencies | select FeatureId
                    
                    $waFeatureNames = ""
                    foreach ($fid in $wapremiumfeatureids) {
                        $waFeatureNames += (Get-SPFeature -Id $fid.FeatureId).DisplayName + ", "
                    }

                    $bWAPremium = $true
                }

                ## Collect Site Collection premium features
                $scPremiumFeatures = Get-SPFeature "PremiumSite" -Web $scURL -ErrorAction SilentlyContinue
		        foreach ($scPremiumFeature in $scPremiumFeatures) {
                    $scpremiumfeatureids = $scPremiumFeature.ActivationDependencies | select FeatureId
                    
                    $scFeatureNames = ""
                    foreach ($fid in $scpremiumfeatureids) {
                        $scFeatureNames += (Get-SPFeature -Id $fid.FeatureId).DisplayName + ", "
                    }

                    $bSCPremium = $true
		        }

                ## Collect Site premium features
                $sitePremiumFeatures = Get-SPFeature "PremiumWeb" -Web $siteURL -ErrorAction SilentlyContinue
		        foreach ($sitePremiumFeature in $sitePremiumFeatures) {
                    $sitepremiumfeatureids = $sitePremiumFeature.ActivationDependencies | select FeatureId
                    
                    $siteFeatureNames = ""
                    foreach ($fid in $sitepremiumfeatureids) {
                        $siteFeatureNames += (Get-SPFeature -Id $fid.FeatureId).DisplayName + ", "
                    }

                    $bSitePremium = $true
		        }
                $details | Add-Member -MemberType NoteProperty -Name "IsPremiumFeatureEnabled" -Value $bSitePremium
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationFeatureNames" -Value $waFeatureNames
                $details | Add-Member -MemberType NoteProperty -Name "SCFeatureNames" -Value $scFeatureNames
                $details | Add-Member -MemberType NoteProperty -Name "SiteFeatureNames" -Value $siteFeatureNames
                
                ## Collect Site users details
                $users = $siteWeb.Users
                if ($users) {
                    $ADUserAsGroupNames = ""
                    foreach ($user in $users) {
                        ## Get list of Domain Group from users of SP groups.
                        if ($user.IsDomainGroup) {
                            $ADUserAsGroupNames += ($user.Name + ";")

                            ## Add AD group to global array
                            if ($arrSPADGroups -notcontains $user.Name) {
                                $arrSPADGroups += $user.Name
                            }
                        }
                        else {
                            $userLogins += ($user.Name + '[' + $user.UserLogin + ']') + ";"
                        }
                    }
                }

                ## Get list of Site Groups
                $siteUserGroups = $siteWeb.Groups 
                foreach ($siteUserGroup in $siteUserGroups) {
                    ## List of groups in site
                    $userGroups += ( $siteUserGroup.LoginName + ";" )

                    ## Select specific details from the SP group object
                    $arrUserGroupDetails = $siteUserGroup | select Name, LoginName, Owner, Description, ParentWeb, DistributionGroupEmail

                    ## Get Group users
                    $grpusers = $siteUserGroup.Users
                    $grpUsersList = ""                        
                    $ADGroupNames = ""

                    foreach ($grpuser in $grpusers) {
                        ## Get list of Domain Group from users of SP groups.
                        if ($grpuser.IsDomainGroup) {
                            ## Add Domain group to array. We add Name as UserLogin has SharePoint reference ID for the group
                            $ADGroupNames += ($grpuser.Name + ";")

                            ## Add AD group to global array
                            if ($arrSPADGroups -notcontains $grpuser.Name) {
                                $arrSPADGroups += $grpuser.Name
                            }
                            #$grpADUsersList += $grpuser.UserLogin + "(" + $grpuser.Name + "):ADGroup" + ";"
                        }
                        else {
                            ## We only add the domain name for user and remove the initial reference by SP
                            $grpUsersList += $grpuser.UserLogin + ";"
                        }
                    }
                    $arrUserGroupDetails | Add-Member -MemberType NoteProperty -Name "ADGroups" -Value $ADGroupNames
                    $arrUserGroupDetails | Add-Member -MemberType NoteProperty -Name "UsersList" -Value $grpUsersList
                    
                    ## Check the SP Groups array
                    if (!$htResultSPUserGroups.Get_Item($siteUserGroup.LoginName) ) {
                        $htResultSPUserGroups.Add($siteUserGroup.LoginName, $arrUserGroupDetails)
                    }
                }

                ## Add user logins and user groups for the site
                $details | Add-Member -MemberType NoteProperty -Name "UsersList" -Value $userLogins
                $details | Add-Member -MemberType NoteProperty -Name "SPGroups" -Value $userGroups
                $details | Add-Member -MemberType NoteProperty -Name "ADGroups" -Value ($ADUserAsGroupNames + $ADGroupNames)
                
                ## Populate result array
                $result += $details
            }
        }
        
        ## Export Web Applications detail
        $result | Export-Csv -Path $OutputFile1 -NoTypeInformation -Encoding UTF8
        LogText "SharePoint License details exported"

        ##
        ## CAL calculation process begins
        ##

        ## Check if AD Groups and Users file is specified.
        if ( ( ($InputADGroupsFile -ne $null) -and (Test-Path -Path $InputADGroupsFile) ) -and `
             ( ($InputADUsersFile -ne $null) -and (Test-Path -Path $InputADUsersFile) ) ) {        
            ## Load AD Users and Groups csv file
            $importedADGroups = Import-Csv -Path $InputADGroupsFile -Encoding UTF8
            
            ## Load AD users  csv file
            $importedADUsers = Import-Csv -Path $InputADUsersFile -Encoding UTF8

            ## Hash table for AD Groups and its users
            $htADGroups = @{}            
            
            ## Groups Hash Table        
            $htAllGroups = @{}
            foreach ($importedADGroup in $importedADGroups) {
                $htAllGroups.Add($importedADGroup.distinguishedname, $importedADGroup)
            }
            
            ## Users Hash Table        
            $htAllUsers = @{}
            $requiredCAL = ""
            foreach ($importedADUser in $importedADUsers) {
                $htAllUsers.Add($importedADUser.distinguishedname, $requiredCAL)
            }
                
            ## Get details for each AD Group
            foreach ($ADGroup in $arrSPADGroups) {
                $arrGroupUserList = @()

                if ($ADGroup -ne "Everyone" -and $ADGroup -ne "NT AUTHORITY\authenticated users") {
                    ## Process AD groups that are used in SP
                    $userGrpSplit = $ADGroup -split "\\"
                    $getADGroupByName = $importedADGroups | where { (($_.distinguishedname -split ",").ToLower() -like ("*" + $userGrpSplit[1])) }
                    $grpName = $getADGroupByName.Name
                                
                    ## Groups - Everyone or NT system are classified as AD groups since they are all authenticated users
                    ## $grpName will hold only values for actual AD groups
                    ## If the group name is Everyone or NT system then they will be ignored
                    if ($grpName) {
                        ## Get all the users within the group
                        $grpMembers = $getADGroupByName.Members -split ";"

                        ## Check if any of the user is a AD Group
                        foreach ($grpMember in $grpMembers) {
                            $htChildGroupLists = @{}
                        
                            ## Check if the user is a group
                            $grpParent = $importedADGroups | where { $_.distinguishedname -like $grpMember }
                            if ($grpParent) {
                                ## The user is a group
                                $htChildGroupLists.Add($grpParent.name, $grpParent.distinguishedname)
                            }
                            else {
                                ## Add the user
                                $arrGroupUserList += $grpMember
                            }

                            while ($htChildGroupLists.count -ne 0) {
                                ## 
                                $htChildGroups = @{}
                                foreach ($arrChildGroupList in $htChildGroupLists.GetEnumerator()) {
                                
                                    ## Get the Child group details
                                    $hashKey = $arrChildGroupList.Key
                                    $hashValue = $arrChildGroupList.Value

                                    ## Check if the user is a group
                                    $grpChild = $importedADGroups | where { $_.distinguishedname -like $hashValue }
                                    $grpChildMembers = $grpChild.Members -split ";"

                                    foreach($grpChildMember in $grpChildMembers) {  
                                        ## Check if the user is a group
                                        $grpParent = $importedADGroups | where { $_.distinguishedname -like $grpChildMember }                                  
                                        if ($grpParent) {
                                            ## The user is a group
                                            $htChildGroups.Add($grpParent.name, $grpParent.distinguishedname)
                                        }
                                        else {
                                            ## Add the user
                                            $arrGroupUserList += $grpChildMember
                                        }
                                    }
                                }

                                ## Remove old groups from the list and replace with the child groups to avoid looping for same group
                                $htChildGroupLists = $htChildGroups
                            }
                        }

                        ## Add the group and it's members to the array
                        $htADGroups.Add($grpName, $arrGroupUserList)                    
                    }
                }
            } ## End-of-loop

            ## Iterate all SharePoint group(s) to merge AD group(s) user(s) with SP user(s)
            $htAllGroups = @{}
        
            ## Convert Hashes to ArrayList
            foreach ($htResultSPUserGroup in $htResultSPUserGroups.GetEnumerator()) {
                $arrADGroupUsers = @()

                ## Get details of current group
                $grpDetails = $htResultSPUserGroup.Value

                ## Check if the group has an AD Group
                if ($grpDetails.ADGroups) {                
                    $arrADGroups = $grpDetails.ADGroups -split ";"

                    foreach ($grpAD in ($arrADGroups -split "\\")[1]) {
                        $arrADGroupUsers += $htADGroups.GetEnumerator() | where {$_.Key -eq $grpAD}
                    }
                }
                
                ## Find the Distinguished username for the user directly referencing in SP
                $usersList = $htResultSPUserGroup.Value.UsersList -split ";"
                        
                foreach ($user in $usersList) {
                    if ($user -ne "" -and $user -ne "SHAREPOINT\System") {
                        $userSplit = ($user.Split("|")[1]) -split "\\"
                                
                        ## Process AD users that are used in SP
                        $getADUserByName = $importedADUsers | where { (($_.samaccountname).ToLower() -like ("*" + $userSplit[1])) }
                        if ($arrADGroupUsers -notcontains $getADUserByName.distinguishedname) {
                            $arrADGroupUsers += $getADUserByName.distinguishedname
                        }
                    }
                }
                
                $htResultSPUserGroup.Value.UsersList = ($arrADGroupUsers -join ";")

                ## Add the values to All Groups Hash table
                $htAllGroups.Add($htResultSPUserGroup.Key, $htResultSPUserGroup.Value)

                ## Default Action - Add the users to the list as array            
                ## Array for Exporting SP Groups
                $arrResultUserGroups += $htResultSPUserGroup.Value            
            }
            
            ## Export all user groups
            $arrResultUserGroups | Export-Csv -Path $OutputFile2 -NoTypeInformation -Encoding UTF8
		    LogText "User Groups Exported"

            ##
            ## Calculate CAL
            ##

            ## Iterate through sites 
            foreach ($site in $sites) {
            
                $waURL = $site.WebApp.Url
                $scURL = $site.SiteDet.Url
                $siteWebs = $site.SiteDet.AllWebs
                            
                foreach ($siteWeb in $siteWebs) { 
            
                    $bIsPremium = $false

                    $siteURL = $siteWeb.URL
                    
                    ## Get list of Site Groups
                    $siteGroups = $siteWeb.Groups
                    $siteUsers = $siteWeb.Users

                    $sitePremiumFeatures = Get-SPFeature "PremiumWeb" -Web $siteURL -ErrorAction SilentlyContinue
                    if ($sitePremiumFeatures) {
                        ## URL has premium features
                        $bIsPremium = $true
                    }
                    
                    LogText ("Calculating CAL for site - " + $siteURL)

                    ## Get the groups the user is in from AllSPGroups
                    ## Update the Users list with CAL information required
                    foreach ($siteGroup in $siteGroups) {
                        $objGroupDetails = $htAllGroups.GetEnumerator() | where {$_.Key -eq $siteGroup.Name}
                        ## Check if the group has any users within it.
                        
                        if ($objGroupDetails.Value.UsersList.Length -eq 0 ) {
                            continue
                        }
                        
                        $grpMembers = $objGroupDetails.Value.UsersList -split ";"
                        
                        foreach ($grpMember in $grpMembers) {
                            $grpMember = $grpMember.Trim()
                            if ($grpMember -ne "" -and $grpMember -ne "SHAREPOINT\System") {
                            
                                if ($grpMember.contains("CN")) {                                
                                    $member = $htAllUsers.GetEnumerator() | where {$_.Key -eq $grpMember}
                                }
                                else {
                                    if ($grpMember.contains("|") ) {
                                        $userSplit = ($grpMember.Split("|")[1]) -split "\\"
                                        # Its an AD group
                                        if($userSplit[0].StartsWith("c:")) {
                                            Write-Host ("User is a Domain user - " + $userSplit[0])
                                        }
                                    }
                                    elseif ($grpMember.contains("\")) {
                                        $userSplit = $grpMember -split "\\"
                                    }
                                    else {
                                        continue
                                    }

                                    ## Process AD users that are used in SP
                                    $getADUserByName = $importedADUsers | where { (($_.samaccountname).ToLower() -like ("*" + $userSplit[1])) }
                                    $member = $htAllUsers.GetEnumerator() | where {$_.Key -eq $getADUserByName.distinguishedname}
                                }                                   
                                if ($member) {
                                    ## Update the users CAL in UsersHash
                                    Set-SPUserCAL -bIsPremium $bIsPremium -member $member -htAllUsers $htAllUsers                                    
                                }
                            }
                        }
                    }
                    
                    ## Update the Users list with CAL information required
                    ## Users - Has ADUsers and ADGroups
                    foreach ($siteUserDetail in $siteUsers) {
                        $siteUser = $siteUserDetail.LoginName.Trim()

                        ## CAL for All users within site
                        ## check if the user is Authenticated users or Everyone
                        if ($siteUser -eq "NT AUTHORITY\authenticated users" -or $siteUser -eq "c:0(.s|true") {
                            $htAllUsers_CAL = @{}
                            $allUsers = $htAllUsers.GetEnumerator()

                            foreach ( $user in $allUsers ) {                                    
                                if ($user)  {                     
                                    ## Update the users CAL in UsersHash
                                    if ($bIsPremium) {
                                        $htAllUsers_CAL.Set_Item($user.Key, "Premium")
                                    }
                                    else {
                                        if ($user.Value -ne "Premium") {
                                            $htAllUsers_CAL.Set_Item($user.Key, "Standard")
                                        }
                                    }
                                }
                            }

                            ## Update the original Hash table
                            $htAllUsers = $htAllUsers_CAL
                        }
                        else {
                            if ($siteUser -ne "" -and $siteUser -ne "SHAREPOINT\System") {
                                
                                if ($siteUser.contains("CN")) {                                
                                    $member = $htAllUsers.GetEnumerator() | where {$_.Key -eq $siteUser}
                                }
                                else {
                                    if ($siteUser.contains("|") ) {
                                        # Its a group
                                        if($siteUser.StartsWith("c:")) {
                                            Set-ADGroupMembersCAL -ADGroupName $siteUSer.Name -htADGroups $htADGroups -bIsPremium $bIsPremium -htAllUsers $htAllUsers
                                        }
                                        $userSplit = ($siteUser.Split("|")[1]) -split "\\"
                                    }
                                    elseif ($siteUser.contains("\")) {
                                        $userSplit = $siteUser -split "\\"
                                    }
                                    else {
                                        continue
                                    }

                                    ## Process AD users that are used in SP
                                    $getADUserByName = $importedADUsers | where { (($_.samaccountname).ToLower() -like ("*" + $userSplit[1])) }
                                    $member = $htAllUsers.GetEnumerator() | where {$_.Key -eq $getADUserByName.distinguishedname}
                                }                                   
                                
                                if ($member) {
                                    ## Update the users CAL in UsersHash
                                    Set-SPUserCAL -bIsPremium $bIsPremium -member $member -htAllUsers $htAllUsers 
                                }
                            }
                        }                
                    } ## End foreach siteUser
                }
            }
            
            ## Process csv
            $userCALs = $htAllUsers.GetEnumerator()

            foreach ($userCAL in $userCALs) {
                $details = New-Object PSObject
                $details | Add-Member -MemberType NoteProperty -Name "User" -Value $userCAL.Key
                $details | Add-Member -MemberType NoteProperty -Name "SPCALRequired" -Value $userCAL.Value

                $resultUserCAL += $details
            }
            
            ## Export user CAL information
            $resultUserCAL | Export-Csv -Path $OutputFile3 -NoTypeInformation -Encoding UTF8
		    LogText "User CAL report exported"
        }
		else {
            LogText "CAL cannot be calculated for following reasons:"
            LogText "   1. AD Groups file could not be located in current working directory."
            LogText "   2. AD Users file could not be located in current working directory."
        }        
    }
    catch {
        LogError -errorDescription "An exception ocured while processing CAL for SharePoint"
		LogLastException
    }
}

# Call the Get-SharePointLicenseDetails Function to 
#	- Get Web Applications, Site Collections and Site details
#	- Get AD and SharePoint Groups and AD Users
#   - Calculate Required CALs
#	- Export CSV
Get-SharePointLicenseDetails