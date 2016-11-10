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
    [alias("o1")]
    $OutputFile1 = "SharePointSites.csv",
	[alias("o2")]
    $OutputFile2 = "CALRequirements.csv",
	[alias("log")]
	[string] $LogFile = "SPLogFile.txt",
	[switch]
	$Verbose = $true
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

function LogProgress($progressDescription){
	if ($Verbose){
		LogText ""
	}

	$output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
	LogText $output -Color Green
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
	LogText -Color Gray "Log File:             $LogFile"
    LogText -Color Gray "Verbose:              $Verbose"
	LogText -Color Gray ""
}

function SearchAD ($searchFilter, [string[]]$searchAttributes, [switch]$useNamingContext){
    
    $searchResults = @()
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    
    if ($useNamingContext){
        # Connect to the Configuration Naming Context
        $rootDSE = [ADSI]"LDAP://RootDSE"
        $configSearchRoot = [ADSI]("LDAP://" + $rootDSE.Get("configurationNamingContext"))
        $objSearcher.SearchRoot = $configSearchRoot
    }
    elseif ($SearchRoot){
        $objDomain = New-Object System.DirectoryServices.DirectoryEntry($SearchRoot)
        $objSearcher.SearchRoot = $objDomain
    }
    else {
        $objDomain = New-Object System.DirectoryServices.DirectoryEntry
        $objSearcher.SearchRoot = $objDomain
    }

    $objSearcher.PageSize = 1000
    $objSearcher.Filter = $searchFilter
    $objSearcher.SearchScope = "Subtree"

    if ($searchAttributes) {
        ($searchAttributes | %{$objSearcher.PropertiesToLoad.Add($_)}) | out-null
    }
    
    $objSearcher.FindAll() | % {
        $pso = New-Object PSObject
        $value = ""
        $_.Properties.GetEnumerator() | % {
            try {
                if ($_.Name -eq "objectsid") {
                    $Counter = 0
                    $Ba = New-Object Byte[] $_.Value[0].Length
                    $_.Value[0] | %{$Ba[$Counter++] = $_}
                    $value = (New-Object System.Security.Principal.SecurityIdentifier($Ba, 0)).Value
                }
                elseif ($_.Name -eq "objectguid" -or $_.Name -eq "msExchMailboxGuid") {
                    $Counter = 0
                    $Ba = New-Object Byte[] $_.Value[0].Length
                    $_.Value[0] | %{$Ba[$Counter++] = $_}
                    $value = (New-Object System.Guid -ArgumentList @(,$Ba)).ToString()
                }
                elseif (($_.Name -eq "lastLogon") -or ($_.Name -eq "lastLogonTimestamp")) {
                    $value = [DateTime]::FromFileTime($_.Value[0]).ToString('yyyy-MM-dd hh:mm:ss')
                }
                elseif ($_.Name -eq "servicePrincipalName"){
                    $value = $_.Value -join ";"
                }
                else {
                    $value = ($_.Value | foreach {$_})
                }
                Add-Member -InputObject $pso -MemberType NoteProperty -Name $_.Name -Value $value
            }
            catch {
                LogLastException
            }
        } 
        $searchResults = $searchResults + $pso
    }

    return $searchResults | select-object $searchAttributes | Where-Object {$_ -ne $null}
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

function GetScriptFolder
{
	$strScriptPath = GetScriptPath
	return split-path -parent $strScriptPath
}

function RunADScript
{
	$strScriptFolder = GetScriptFolder
	$strADScriptPath = "$strScriptFolder\Get-ADDetails.ps1"
	if (-not (Test-Path $strADScriptPath)) {
		LogError "AD Group & User info missing."
			"AD script ($strADScriptPath) not found"
		return
	}

	LogProgress "Running AD Script"
	$strADScriptCommand = "$strADScriptPath -o5 $InputADUsersFile -o7 $InputADGroupsFile"
	
	invoke-expression -Command $strADScriptCommand
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
		
		if (IsLocalComputer($SharePointServer)) {
			QuerySharePointInfo
		}
		else{
			# Execuing SharePoint script over a remote session requires CredSSP authentication option
			# to avoid double hop authentication issues. This is generally blocked by Group Policy.
			# We use PSExec and user credentials to avoid this issue.
			
			# Ensure PSExec is available
			$strScriptFolder = GetScriptFolder
			$strPSExecPath = "$strScriptFolder\PSExec.exe"
			if (!(Test-Path $strPSExecPath)) {
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

					if (-not $UserName -like "*\*"){
						LogError "Warning: SharePoint script generally requires username in [Domain\Username] format.", 
							"Script will continue, but may fail due to invalid login."
					}
				}
			}
			
			LogProgress "Connecting to remote SharePoint server ($SharePointServer) with User Name `"$UserName`""
			if(!($UserName -and $Password)){
				LogError "Warning: SharePoint script generally requires username and password to be specified.", 
					"Script will continue, but may fail due to insufficient privileges."
				$session = New-PSSession -ComputerName $SharePointServer -ErrorAction SilentlyContinue -ErrorVariable strConnectionError
			}
			else {
				$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        		$PSCreds = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)

				$session = New-PSSession -ComputerName $SharePointServer ?Credential $PSCreds -ErrorAction SilentlyContinue -ErrorVariable strConnectionError
			}

			
			if(-not($session)) {
                LogError "Unable to connect to server $SharePointServer", 
					"It may be necessary to enable PowerShell Remoting on the remote server with the Enable-PSRemoting command", 
					$strConnectionError
                return
            }
			
			# Move required files to remote computer
			LogProgress "Moving required files to $SharePointServer"
			# Move script file
			$strLocalScriptPath = GetScriptPath
			$strLocalScriptFolder = GetScriptFolder
			$strLocalADScriptPath = "$strLocalScriptFolder\Get-ADDetails.ps1"
			$strRemoteDefaultFolder = GetRemoteDefaultFolder -RemoteSession $session
			$strRemoteScriptPath = "$strRemoteDefaultFolder\Get-SharePointLicenseDetails.ps1"
			PutFile -LocalFilePath $strLocalScriptPath -RemoteFilePath $strRemoteScriptPath -RemoteSession $session
			
			if ((Test-Path $InputADGroupsFile) -and (Test-Path $InputADUsersFile)) {
				# Move AD files (If they exist)
				PutFile -LocalFilePath $InputADGroupsFile -RemoteFilePath "ADGroups.csv" -RemoteSession $session
				PutFile -LocalFilePath $InputADUsersFile -RemoteFilePath "ADUsers.csv" -RemoteSession $session
			}
			elseif (Test-Path $strLocalADScriptPath) {
				# AD files don't exist - Move teh AD script instead
				PutFile -LocalFilePath $strLocalADScriptPath -RemoteFilePath "Get-ADDetails.ps1" -RemoteSession $session
			}
			
			# Reset remote log file
			PutFile -LocalFilePath "" -RemoteFilePath "SPLogFile.txt" -RemoteSession $session
			
			# Execute the script remotely
			LogProgress "Executing script remotely using PSExec"
			if($UserName -and $Password){
				Start-Process $strPSExecPath -ArgumentList "\\$SharePointServer -h -accepteula -w ""$strRemoteDefaultFolder"" -u $UserName -p $Password powershell.exe -ExecutionPolicy Bypass -File ""$strRemoteScriptPath"" -server $SharePointServer -headless" -Wait -NoNewWindow -WorkingDirectory $strRemoteDefaultFolder -ErrorVariable errPS
			}
			else {
				Start-Process $strPSExecPath -ArgumentList "\\$SharePointServer -h -accepteula -w ""$strRemoteDefaultFolder"" powershell.exe -ExecutionPolicy Bypass -File ""$strRemoteScriptPath"" -server $SharePointServer -headless" -Wait -NoNewWindow -WorkingDirectory $strRemoteDefaultFolder -ErrorVariable errPS
			}

			# Copy the results back to this device
			LogProgress "Collecting output files from $SharePointServer"
			GetFile -RemoteFilePath "SharePointSites.csv" -LocalFilePath $OutputFile1 -RemoteSession $session
			GetFile -RemoteFilePath "SharePointUserGroups.csv" -LocalFilePath $OutputFile4 -RemoteSession $session
			GetFile -RemoteFilePath "SharePointUserCALs.csv" -LocalFilePath $OutputFile5 -RemoteSession $session
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
			DeleteRemoteFile -RemoteSession $session -RemoteFilePath $strRemoteScriptPath
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

function GetGroupInfo {
    $groupAttributes = "name", "samaccountname", "description", "distinguishedName", "whenChanged", "whenCreated"
    $groupList = SearchAD -searchAttributes $groupAttributes -searchFilter "(objectClass=group)" 
    $groupList | %{
        $groupDN = $_.distinguishedName
        $objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($groupDN)")
        Add-Member -InputObject $_ -MemberType NoteProperty -Name "Members" -Value ($objDomain.Member -join ";")
        
        # Update HashTable of group membership
        $objDomain.Member | %{
			if ($_ -ne $null) {
				if (-not $groupMembership.ContainsKey($_)){
                $groupMembership[$_] = New-Object System.Collections.ArrayList($null)
				}
				$itemCount = $groupMembership[$_].Add($groupDN)
			}
        }
    } 
    $groupList
}

function GetUserInfo {
    $userAttributes = "sAMAccountName", "objectSid", "objectGUID", "displayName", "departmentNumber", "company", "department", "distinguishedName", "lastLogon", "lastLogonTimestamp", "logonCount", "mail", "telephoneNumber", "physicalDeliveryOfficeName", "description", "whenChanged", "whenCreated", "msExchMailboxGuid","userAccountControl"
    $userList = SearchAD -searchAttributes $userAttributes -searchFilter "(&(objectCategory=person)(objectClass=user))" 
    # Add Group Info to user list
    $userList | % {
        $groups = ""
        if ($_.distinguishedName -and $groupMembership.ContainsKey($_.distinguishedName)) {
            $groups = $groupMembership[$_.distinguishedName] -join ";"
        }
        Add-Member -InputObject $_ -MemberType NoteProperty -Name "Groups" -Value $groups
    }
    $userList
}

function DecodeSharePointUserName ([string] $SharePointUserName) {
	# SharePoint names can be decorated with claim type
	# http://social.technet.microsoft.com/wiki/contents/articles/13921.sharepoint-2013-claims-encoding-also-valuable-for-sharepoint-2010.aspx
	
	# "i:0#.w|sam360\jon.mulligan"    (AD User)         => "sam360\jon.mulligan"
	# "SHAREPOINT\system"             (SharePoint User) => "sharepoint\system"
	# "i:0#.f|partnerweb|partner_1"   (Forms User)      => "partnerweb->partner_1"
	$decodedUserName = ""
	$sharePointUserNameParts = $SharePointUserName -split "\|"
	if ($sharePointUserNameParts.count -eq 1) {
	    $decodedUserName = $SharePointUserName
	}
	elseif ($sharePointUserNameParts.count -eq 2) {
	    $decodedUserName = $sharePointUserNameParts[1]
	}
	elseif ($sharePointUserNameParts.count -gt 2) {
	    $decodedUserName = $SharePointUserName # Unknown format
	}
	return $decodedUserName.ToLower()
}

function GetSPWebList {

	$lstAllWebs = @()
	
	## Each 'Web Application' contains 1 or more 'Site Collections'
	## Each 'Site Collection' contains 1 or more 'Sites'
	## The SharePoint API calls these 'Web Applications', 'Sites' & 'SPWebs' respectively.
	Get-SPWebApplication | %{
		$spWebApplication = $_
		  
        ## Web Application - What premium features are enabled
		$waFeatureNames = ""
		$bWAPremium = $false
        $waPremiumFeatures = Get-SPFeature "PremiumWebApplication" -WebApplication $spWebApplication -ErrorAction SilentlyContinue -ErrorVariable errGetSPFeature
        foreach ($wapremiumfeature in $waPremiumFeatures) {
            $wapremiumfeatureids = $wapremiumfeature.ActivationDependencies | select FeatureId
            foreach ($fid in $wapremiumfeatureids) {
                $waFeatureNames += (Get-SPFeature -Id $fid.FeatureId).DisplayName + ", "
            }
            $bWAPremium = $true
        }
		
		Get-SPSite -Limit All -WebApplication $spWebApplication | %{
			$spSite = $_
			
            ## Site Collection - What premium features are enabled
			$scFeatureNames = ""
			$bSCPremium = $false
            $scPremiumFeatures = Get-SPFeature "PremiumSite" -Web $spSite.Url -ErrorAction SilentlyContinue
	        foreach ($scPremiumFeature in $scPremiumFeatures) {
                $scpremiumfeatureids = $scPremiumFeature.ActivationDependencies | select FeatureId
                foreach ($fid in $scpremiumfeatureids) {
                    $scFeatureNames += (Get-SPFeature -Id $fid.FeatureId).DisplayName + ", "
                }
                $bSCPremium = $true
	        }
			
			Get-SPWeb -Limit All -Site $spSite | %{
				$spWeb = $_
				$siteURL = $spWeb.URL
				$webUserNames = @()
				$webSPGroupNames = @()
				$webADGroupNames = @()
				
				if ($Verbose) {
					LogText "Querying details for site: $siteURL"
				}
				
                ## Site - What premium features are enabled
				$siteFeatureNames = ""
				$bSitePremium = $false
                $sitePremiumFeatures = Get-SPFeature "PremiumWeb" -Web $siteURL -ErrorAction SilentlyContinue
		        foreach ($sitePremiumFeature in $sitePremiumFeatures) {
                    $sitepremiumfeatureids = $sitePremiumFeature.ActivationDependencies | select FeatureId
                    foreach ($fid in $sitepremiumfeatureids) {
                        $siteFeatureNames += (Get-SPFeature -Id $fid.FeatureId).DisplayName + ", "
                    }
                    $bSitePremium = $true
		        }
				
				## Get list of Web users
                foreach ($webUser in $spWeb.Users) {
                    ## 'Users' can be AD users or AD groups
                    if ($webUser.IsDomainGroup) {
                        $webADGroupNames += $webUser.Name
						$global:lstAllSPADGroups += $webUser.Name
                    }
                    else {
						$webUserNames += (DecodeSharePointUserName($webUser.UserLogin))
                    }
                }

                ## Get list of Web groups
                foreach ($webSPUserGroup in $spWeb.Groups) {
                    ## List of groups in site
                    $webSPGroupNames += $webSPUserGroup.LoginName

                    foreach ($webGrpUser in $webSPUserGroup.Users) {
                        ## 'Users' can be AD users or AD groups
                        if ($webGrpUser.IsDomainGroup) {
                            $webADGroupNames += $webGrpUser.Name
							$global:lstAllSPADGroups += $webGrpUser.Name
                        }
                        else {
							$webUserNames += (DecodeSharePointUserName($webUser.UserLogin))
                        }
                    }
                }
				
				##
                ## Web Application Details
                ##
				$details = New-Object PSObject
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationUrl" -Value $spWebApplication.Url
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationName" -Value $spWebApplication.Name
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationDisplayName" -Value $spWebApplication.DisplayName
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationId" -Value $spWebApplication.Id
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationFarmName" -Value $spWebApplication.Farm
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationStatus" -Value $spWebApplication.Status
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationVersion" -Value $spWebApplication.Version
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationApplicationPoolName" -Value $spWebApplication.ApplicationPool

                ##
                ## Site Collection Details
                ##
                $details | Add-Member -MemberType NoteProperty -Name "SCUrl" -Value $spSite.Url
                $details | Add-Member -MemberType NoteProperty -Name "SCHostName" -Value $spSite.HostName
                $details | Add-Member -MemberType NoteProperty -Name "SCID" -Value $spSite.ID	                
				$details | Add-Member -MemberType NoteProperty -Name "SCArchived" -Value $spSite.Archived
                $details | Add-Member -MemberType NoteProperty -Name "SCCreation Date" -Value $spSite.CertificationDate
                $details | Add-Member -MemberType NoteProperty -Name "SCLastContentModifiedDate" -Value $spSite.LastContentModifiedDate
                $details | Add-Member -MemberType NoteProperty -Name "SCLastSecurityModifiedDate" -Value $spSite.LastSecurityModifiedDate
            
                ##
                ## Site Details
                ##
                $details | Add-Member -MemberType NoteProperty -Name "SiteUrl" -Value $spWeb.Url
                $details | Add-Member -MemberType NoteProperty -Name "SiteTitle" -Value $spWeb.Title
                $details | Add-Member -MemberType NoteProperty -Name "SiteName" -Value $spWeb.Name
                $details | Add-Member -MemberType NoteProperty -Name "SiteID" -Value $spWeb.ID
                $details | Add-Member -MemberType NoteProperty -Name "SiteDescription" -Value $spWeb.Description
                $details | Add-Member -MemberType NoteProperty -Name "SiteAuthor" -Value (DecodeSharePointUserName($spWeb.Author))
                $details | Add-Member -MemberType NoteProperty -Name "SiteParentWeb" -Value $spWeb.ParentWeb
                $details | Add-Member -MemberType NoteProperty -Name "SiteParentWebId" -Value $spWeb.ParentWebId
                $details | Add-Member -MemberType NoteProperty -Name "SiteIsAppWeb" -Value $spWeb.IsAppWeb
                $details | Add-Member -MemberType NoteProperty -Name "SiteIsRootWeb" -Value $spWeb.IsRootWeb
                $details | Add-Member -MemberType NoteProperty -Name "SiteHasUniqueRoleDefinitions" -Value $spWeb.HasUniqueRoleDefinitions
                $details | Add-Member -MemberType NoteProperty -Name "SiteAllowAnonymousAccess" -Value $spWeb.AllowAnonymousAccess
                $details | Add-Member -MemberType NoteProperty -Name "SiteWebTemplate" -Value $spWeb.WebTemplate
                $details | Add-Member -MemberType NoteProperty -Name "SiteUIVersion" -Value $spWeb.UIVersion
                $details | Add-Member -MemberType NoteProperty -Name "SiteCreation Date" -Value $spWeb.Created
                $details | Add-Member -MemberType NoteProperty -Name "SiteLastItemModifiedDate" -Value $spWeb.LastItemModifiedDate
                
                ##
				## Premium Features
				##
                $details | Add-Member -MemberType NoteProperty -Name "IsPremiumFeatureEnabled" -Value $bSitePremium
                $details | Add-Member -MemberType NoteProperty -Name "WebApplicationFeatureNames" -Value $waFeatureNames
                $details | Add-Member -MemberType NoteProperty -Name "SCFeatureNames" -Value $scFeatureNames
                $details | Add-Member -MemberType NoteProperty -Name "SiteFeatureNames" -Value $siteFeatureNames
                
                ##
				## User & Group access 
				##
                $details | Add-Member -MemberType NoteProperty -Name "UsersList" -Value (($webUserNames | select -Unique) -join ";")
                $details | Add-Member -MemberType NoteProperty -Name "SPGroups" -Value (($webSPGroupNames | select -Unique) -join ";")
                $details | Add-Member -MemberType NoteProperty -Name "ADGroups" -Value (($webADGroupNames | select -Unique) -join ";")
                $details | Add-Member -MemberType NoteProperty -Name "UserArray" -Value ($webUserNames | select -Unique)
                $details | Add-Member -MemberType NoteProperty -Name "ADGroupArray" -Value ($webADGroupNames | select -Unique)
                
				
                ## Populate result array
                $lstAllWebs += $details
			}
		}
	}
	$lstAllWebs
}

function GetNCName([String]$strDN) {
	$pos = $strDN.ToLower().IndexOf("dc=")
	if ($pos -gt 0){
		$strDN = $strDN.Substring($pos)
	}
	return $strDN
}

function GetCN([String]$strDN) {
	$pos = $strDN.IndexOf("=")
	if ($pos -gt 0){
		$strDN = $strDN.Substring($pos+1)
		$pos = $strDN.IndexOf(",")
		if ($pos -gt 0){
			$strDN = $strDN.Substring(0, $pos)
		}
	}
	return $strDN
}

function ExpandADGroup([string]$strADGroupName) {

	if (-not $htGroupMembers.ContainsKey($strADGroupName)) {
		## Unable to find Group using Domain\Name format
		## Try CN format
		if (-not $htGroupCNs.ContainsKey($strADGroupName)) {
			LogText "Unable to find AD group '$strADGroupName'" -color red
			return
		}
		else {
			## CN was found. Lookup Domain\Name
			$strADGroupName = $htGroupCNs[$strADGroupName]
		}
	}
	
	if ($htExpandedADGroups.ContainsKey($strADGroupName)) {
		return $htExpandedADGroups[$strADGroupName]
	}
	
	$allGroupMembers = @()
	$directGroupMembers = $htGroupMembers[$strADGroupName]
	foreach ($directGroupMember in $directGroupMembers)
	{
		if ($htUserTable.ContainsKey($directGroupMember)) {
			## The member is a user
			$allGroupMembers += $htUserTable[$directGroupMember]
		}
		elseif ($htGroupDNs.ContainsKey($directGroupMember)) {
			## The member is a group
			$allGroupMembers += (ExpandADGroup($htGroupDNs[$directGroupMember]))
		}
		elseif ($directGroupMember -eq "Everyone") {
			## The member is a group
			$allGroupMembers += "Everyone"
		}
		else {
			LogText "Unable to find AD member '$directGroupMember'" -color red
		}
	}
	
	if ($Verbose) {
		LogText "Expanded Group '$strADGroupName' - $($allGroupMembers.count) members"
	}
	
	$global:htExpandedADGroups[$strADGroupName] = $allGroupMembers
	return $allGroupMembers
}

function ExpandAllADGroups {
	## Get domain NetBIOS names
	$htDomainLookup = @{}
	$domainNetBIOSDetailsAttributes = "netbiosname", "ncname"
    $domainNetBIOSDetails = SearchAD -searchFilter "(NetBIOSName=*)" -searchAttributes $domainNetBIOSDetailsAttributes -useNamingContext
	foreach ($domainNetBIOSDetail in $domainNetBIOSDetails) {
		$htDomainLookup[$domainNetBIOSDetail.ncname] = $domainNetBIOSDetail.netbiosname
	}

	## Create a Hashtable map between user DNs and user Domain\Username
	$htUserTable = @{}
	foreach ($adUser in $adUserList) {
		$userDomainName = ""
		$userDomainNCName = GetNCName($adUser.DistinguishedName)
		if ($htDomainLookup.ContainsKey($userDomainNCName)){
			$userDomainName = $htDomainLookup[$userDomainNCName]
		}
		else {
			LogText "Unable to find NetBIOS domain name for user '$($adUser.DistinguishedName)'" -color red
		}

		$htUserTable[$adUser.DistinguishedName] = "$userDomainName\$($adUser.SAMAccountName)"
	}
	
	## Create 3 Hashtables for looking up groups
	$htGroupMembers = @{} # Map group Domain\Name to group members
	$htGroupDNs = @{} # Map group DN to group Domain\Name
	$htGroupCNs = @{} # Map group CN to group Domain\Name
	foreach ($adGroup in $adGroupList) {
		$groupDomainName = ""
		$groupDomainNCName = GetNCName($adGroup.DistinguishedName)
		if ($htDomainLookup.ContainsKey($groupDomainNCName)){
			$groupDomainName = $htDomainLookup[$groupDomainNCName]
		}
		else {
			LogText "Unable to find NetBIOS domain name for group '$($adGroup.DistinguishedName)'" -color red
		}
		
		$groupDomainSAMName = "$groupDomainName\$($adGroup.samaccountname)"
		$htGroupMembers[$groupDomainSAMName] = $adGroup.Members -split ";"
		$htGroupDNs[$adGroup.DistinguishedName] = $groupDomainSAMName
		$htGroupCNs[(GetCN($adGroup.DistinguishedName))] = $groupDomainSAMName
	}
	
	## Special Cases
	$htGroupMembers["Everyone"] = ,"Everyone"
	$htGroupMembers["NT AUTHORITY\authenticated users"] = ,"Everyone"
	#$htGroupDNs["Everyone"] = "Everyone"
	#$htGroupDNs["NT AUTHORITY\authenticated users"] = "NT AUTHORITY\authenticated users"
	
	
	## Recursively expand the group membership for all groups used in SharePoint
	foreach ($adGroupName in ($lstAllSPADGroups | select -Unique)) {
		$groupMembers = ExpandADGroup($adGroupName)
	}
}

function GetCALRequirements {
	$htCALRequirements = @{}
	foreach ($spWeb in $spWebList) {
		$bIsPremium = $spWeb.IsPremiumFeatureEnabled
		
		$lstWebUsers = @()
		$lstWebUsers += $spWeb.SiteAuthor
		$lstWebUsers += $spWeb.UserArray
		foreach ($adGroup in $spWeb.ADGroupArray){
			if ($htExpandedADGroups.ContainsKey($adGroup)){
				$lstWebUsers += $htExpandedADGroups[$adGroup]
			}
			else {
				LogText "Unable to find group '$adGroup' for Site '$($spWeb.SiteUrl)'" -color red
			}
		}
		
		foreach ($webUser in ($lstWebUsers | select -Unique)) {
			if ($bIsPremium) {
				$htCALRequirements[$webUser] = "Premium"
			}
			else {
				if (-not $htCALRequirements.ContainsKey($webUser)){
					$htCALRequirements[$webUser] = "Standard"
				}
			}
		}
	}
	
	$htCALRequirements.GetEnumerator() |  
		Select-Object -Property @{n='User';e={$_.Name}}, @{n='CAL';e={$_.Value}} |
		sort-object User	
}

function QuerySharePointInfo {
	try {
        if (!(EnvironmentConfigured)) {
            LogProgress "Adding SharePoint Snapin"
            Add-PSSnapin Microsoft.Sharepoint.Powershell
		    if (!(EnvironmentConfigured)) {
			    LogError "SharePoint Environment could not be configured. Please enter a SharePoint server name and try again"
			    return $false
		    }
		}

        LogProgress "Updating permissionf for user $Username"
        Get-SPDatabase | Add-SPShellAdmin $Username

		## Global group lists
		$global:lstAllSPADGroups = @()
		$global:htExpandedADGroups = @{}
        
        ##
        ## Get SPWeb Info
		##
		LogProgress "Getting Site Information"
		$spWebList = GetSPWebList
        $spWebList | select * -exclude "*Array" |
			Export-Csv -Path $OutputFile1 -NoTypeInformation -Encoding UTF8
		
		##
		## Get AD Group Info
		##
		LogProgress "Getting AD Group Information"
		$groupMembership = @{}
		$adGroupList = GetGroupInfo
		#$adGroupList | export-csv $OutputFile2 -notypeinformation -Encoding UTF8
		
		##
		## Get AD User Info
		##
		LogProgress "Getting AD User Information"
		$adUserList = GetUserInfo
		#$adUserList | export-csv $OutputFile3 -notypeinformation -Encoding UTF8

        ##
        ## Expand all groups
        ##
		LogProgress "Compiling group memberships"
		ExpandAllADGroups
		
		##
		## Calculate CAL requirements
		##
        LogProgress "Calculating user CAL requirements"
		$userCALRequirements = GetCALRequirements
		$userCALRequirements | export-csv $OutputFile2 -notypeinformation -Encoding UTF8

	    
		LogProgress "Export Complete"
        
    }
    catch {
		LogLastException
    }
}

# Call the Get-SharePointLicenseDetails function to 
#	- Get Web Applications, Site Collections and Site details
#	- Get AD and SharePoint Groups and Users info
#   - Calculate Required CALs
Get-SharePointLicenseDetails