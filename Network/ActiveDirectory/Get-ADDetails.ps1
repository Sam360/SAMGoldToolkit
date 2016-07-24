 ##########################################################################
 # 
 # 	Get-ADDetails
 # 	SAM Gold Toolkit
 #
 ##########################################################################
 
 <#
.SYNOPSIS
Retrieves domain, user, device, server & mobile device data from Active Directory

.DESCRIPTION
The Get-ADDetails script queries the local domain for domain, user, device, server 
& mobile device data and produces 8 CSV files
	1)    ADDomains.csv - One record per domain
	2)    ADDomainTrusts.csv - One record per external trusted domain
	3)    ADDomainNETBIOS.csv - One record per domain (Includes domain NetBIOS name)
	4)    ADDomainControllers.csv - One record per domain controller for current domain
	5)    ADUsers.csv - One record per domain user
	6)    ADDevices.csv - One record per domain computer
	7)    ADExchangeServers.csv - One record per Exchange Server
	8)    ADActiveSyncDevices.csv - One record per Exchange Active Sync Device

    Files are written to current working directory

.PARAMETER Verbose 
Flag - Display extra info to screen

.EXAMPLE
Get all domain, user, device, server & mobile device data from current domain
Get-ADDetails –Verbose

#>

 Param(
	[alias("o1")]
	[string] $OutputFile1 = "ADDomains.csv",
	[alias("o2")]
	[string] $OutputFile2 = "ADDomainTrusts.csv",
	[alias("o3")]
	[string] $OutputFile3 = "ADDomainNETBIOS.csv",
	[alias("o4")]
	[string] $OutputFile4 = "ADDomainControllers.csv",
	[alias("o5")]
	[string] $OutputFile5 = "ADUsers.csv",
	[alias("o6")]
	[string] $OutputFile6 = "ADDevices.csv",
	[alias("o7")]
	[string] $OutputFile7 = "ADExchangeServers.csv",
	[alias("o8")]
	[string] $OutputFile8 = "ADActiveSyncDevices.csv",
	[alias("log")]
	[string] $LogFile = "ADLogFile.txt",
	[alias("r")]
	[string]$SearchRoot = "",
	[switch]
	$Verbose)

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

function LogProgress($progressDescription){
	if ($Verbose){
		LogText ""
	}

	$output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
	LogText $output -Color Green
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
}

function LogEnvironmentDetails {
	LogText -Color Gray "   _____         __  __    _____       _     _   _______          _ _    _ _   "
	LogText -Color Gray "  / ____|  /\   |  \/  |  / ____|     | |   | | |__   __|        | | |  (_) |  "
	LogText -Color Gray " | (___   /  \  | \  / | | |  __  ___ | | __| |    | | ___   ___ | | | ___| |_ "
	LogText -Color Gray "  \___ \ / /\ \ | |\/| | | | |_ |/ _ \| |/ _`` |    | |/ _ \ / _ \| | |/ / | __|"
	LogText -Color Gray "  ____) / ____ \| |  | | | |__| | (_) | | (_| |    | | (_) | (_) | |   <| | |_ "
	LogText -Color Gray " |_____/_/    \_\_|  |_|  \_____|\___/|_|\__,_|    |_|\___/ \___/|_|_|\_\_|\__|"
	LogText -Color Gray " "
	LogText -Color Gray " Get-ADDetails.ps1"
	LogText -Color Gray " "

	$OSDetails = Get-WmiObject Win32_OperatingSystem
	LogText -Color Gray "Computer Name:                   $($env:COMPUTERNAME)"
	LogText -Color Gray "User Name:                       $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	LogText -Color Gray "Windows Version:                 $($OSDetails.Caption)($($OSDetails.Version))"
	LogText -Color Gray "PowerShell Host:                 $($host.Version.Major)"
	LogText -Color Gray "PowerShell Version:              $($PSVersionTable.PSVersion)"
	LogText -Color Gray "PowerShell Word size:            $($([IntPtr]::size) * 8) bit"
	LogText -Color Gray "CLR Version:                     $($PSVersionTable.CLRVersion)"
	LogText -Color Gray "Output File 1:                   $OutputFile1"
	LogText -Color Gray "Output File 2:                   $OutputFile2"
	LogText -Color Gray "Output File 3:                   $OutputFile3"
	LogText -Color Gray "Output File 4:                   $OutputFile4"
	LogText -Color Gray "Output File 5:                   $OutputFile5"
	LogText -Color Gray "Output File 6:                   $OutputFile6"
	LogText -Color Gray "Output File 7:                   $OutputFile7"
	LogText -Color Gray "Output File 8:                   $OutputFile8"
	LogText -Color Gray "Log File:                        $LogFile"
	LogText ""
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
				elseif ($_.Name -eq "objectguid") {
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

	return $searchResults | select-object $searchAttributes
}

function GetDirectoryContext {

	if ($DirectoryContext) {
		return $DirectoryContext
	}

	if (-not $DomainDNS) {
		$DirectoryContext = new-object 'System.DirectoryServices.ActiveDirectory.DirectoryContext'("domain")
		return $DirectoryContext
	}

	if (-not $UserName) {
		$DirectoryContext = new-object 'System.DirectoryServices.ActiveDirectory.DirectoryContext'("forest", $DomainDNS)
		return $DirectoryContext
	}

	$DirectoryContext = new-object 'System.DirectoryServices.ActiveDirectory.DirectoryContext'("domain", $DomainDNS, $UserName, $Password)
	return $DirectoryContext
}

function GetDomainInfo {
	# Get a list of domains in the forest
	#$DC = GetDirectoryContext

	$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
	$forest = $domain.Forest
	
	if ($Verbose){
		LogText "Current Forest:                  $forest"
		LogText "Forest Root Domain:              $($forest.RootDomain)"
		$forestDomainNames = ($forest.Domains | select -expand Name) -join ", "
		LogText "Forest Domains:                  $(CountItems($forest.Domains)) ($forestDomainNames)"

		LogText "Current Domain:                  $domain"
		$domainControllerNames = ($domain.DomainControllers | select -expand Name) -join ", "
		LogText "Domain Controllers:              $(CountItems($domain.DomainControllers)) ($domainControllerNames)"
	}
	
	if ($forest.Domains) {
		$forest.Domains | export-csv $OutputFile1 -notypeinformation -Encoding UTF8
	}
	
	$domainTrustAttributes = "name", "trustpartner", "flatname", "distinguishedname", "adspath", "trustdirection", "trustattributes", "trusttype", "trustposixoffset", "instancetype", "whencreated", "whenchanged"
	$domainTrusts = SearchAD -searchFilter "(objectClass=trustedDomain)" -searchAttributes $domainTrustAttributes
	$domainTrusts | export-csv $OutputFile2 -notypeinformation -Encoding UTF8

	$domainNetBIOSDetailsAttributes = "name", "netbiosname", "ncname", "adspath", "dnsroot", "objectguid", "whencreated", "whenchanged"
	$domainNetBIOSDetails = SearchAD -searchFilter "(NetBIOSName=*)" -searchAttributes $domainNetBIOSDetailsAttributes -useNamingContext
	$domainNetBIOSDetails | export-csv $OutputFile3 -notypeinformation -Encoding UTF8

	if ($domain.DomainControllers){
		$domain.DomainControllers | export-csv $OutputFile4 -notypeinformation -Encoding UTF8
	}

	if ($Verbose){
		$trustDomainNames = ($domainTrusts | select -expand Name) -join ", "
		LogText "Trusted Domains:                 $(CountItems($domainTrusts)) ($trustDomainNames)"
	}
}

function GetUserInfo {
	$userAttributes = "sAMAccountName", "objectSid", "objectGUID", "displayName", "departmentNumber", "company", "department", "distinguishedName", "lastLogon", "lastLogonTimestamp", "logonCount", "mail", "telephoneNumber", "physicalDeliveryOfficeName", "description", "whenChanged", "whenCreated", "msExchMailboxGuid"
	$userList = SearchAD -searchAttributes $userAttributes -searchFilter "(&(objectCategory=person)(objectClass=user))" 
	$userList  | export-csv $OutputFile5 -notypeinformation -Encoding UTF8

	if ($Verbose){
		LogText "User Count:                      $($userList.Count)"

		$cutOfftime = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd hh:mm:ss')
		$activeUsers = $userList | where {(GetMoreRecentDate -date1 $_.lastLogon -date2 $_.lastLogonTimestamp) -gt $cutOfftime}
		LogText "User Count (Active):             $(CountItems($activeUsers))"

		$exchangeMailBoxes = $userList | where {$_.msExchMailboxGuid}
		LogText "Exchange Mailbox Count:          $((($exchangeMailBoxes) | measure-object).count)"
		$activeExchangeMailBoxes = $exchangeMailBoxes | where {(GetMoreRecentDate -date1 $_.lastLogon -date2 $_.lastLogonTimestamp) -gt $cutOfftime}
		LogText "Exchange Mailbox Count (Active): $(CountItems($activeExchangeMailBoxes))"
	}
}

function GetDeviceInfo {
	$deviceAttributes = "name", "objectSid", "objectGUID", "operatingSystem", "operatingSystemVersion", "operatingSystemServicePack", "lastLogon", "lastLogonTimeStamp", "ADsPath", "location", "dNSHostName", "description", "whenChanged", "whenCreated","servicePrincipalName"
	$deviceList = SearchAD -searchAttributes $deviceAttributes -searchFilter "(objectClass=computer)" 
	$deviceList | export-csv $OutputFile6 -notypeinformation -Encoding UTF8

	if ($Verbose){
		LogText "Device Count:                    $($deviceList.Count)"

		$cutOfftime = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd hh:mm:ss')
		$activeDevices = $deviceList | where {(GetMoreRecentDate -date1 $_.lastLogon -date2 $_.lastLogonTimestamp) -gt $cutOfftime}
		LogText "Device Count (Active):           $(CountItems($activeDevices))"

		$clusters = $deviceList | where {$_.servicePrincipalName -ne $null -and
											$_.servicePrincipalName.Contains("MSServerCluster/") }
		$clusterNames = ($clusters | select -expand Name) -join ", "
		LogText "Clusters:                        $(CountItems($clusters)) ($clusterNames)"

		$hyperVHosts = $deviceList | where { $_.servicePrincipalName -ne $null -and (
							$_.servicePrincipalName.Contains("Microsoft Virtual Console Service/") -or 
							$_.servicePrincipalName.Contains("Microsoft Virtual System Migration Service/")) }
		$hyperVHostNames = ($hyperVHosts | select -expand Name) -join ", "
		LogText "HyperV Hosts:                    $(CountItems($hyperVHosts)) ($hyperVHostNames)"

		$exchangeServers = $deviceList | where { $_.servicePrincipalName -ne $null -and (
							$_.servicePrincipalName.Contains("exchangeMDB/") -or 
							$_.servicePrincipalName.Contains("exchangeRFR/")) } 
		$exchangeServerNames = ($exchangeServers | select -expand Name) -join ", "
		LogText "Exchange Servers:                $(CountItems($exchangeServers)) ($exchangeServerNames)"

		LogText ""
		LogText "Operating System Counts:"
		$deviceList | Group-Object operatingSystem | Select Name,Count | Sort Count -desc | ft -autosize | out-string | LogText
	}
}

function DecodeExchangeEdition([string] $encStr) {

    Set-Variable Seed -value 0x49 -option ReadOnly
    Set-Variable Magic -value 0x43 -option ReadOnly

Add-Type -TypeDefinition @"
    public enum ExchangeEditions
    {
        None = -1,
        Standard = 0x0,
        Enterprise = 0x1,
        Evaluation = 0x2,
        Sample = 0x3,
        BackOffice = 0x4,
        Select = 0x5,
        UpgradedStandard = 0x8,
        UpgradedEnterprise = 0x9,
        Coexistence = 0xA,
        UpgradedCoexistence = 0xB
    }
"@


    if ([string]::IsNullOrEmpty($encStr)) {
        Write-Host("Edition string is null. Exiting")
        return -1
    }

    [byte[]]$decodeBuf = [System.Text.Encoding]::Unicode.GetBytes($encStr)

    for ($i=$decodeBuf.Length; $i -gt 1 ; $i--) {
        $decodeBuf[$i - 1] = $decodeBuf[$i - 1] -bxor [byte]($decodeBuf[$i - 2] -bxor $Seed)
    }

    $decodeBuf[0] = $decodeBuf[0] -bxor ($Seed -bor $Magic)

    $strDecodedType = [System.Text.Encoding]::Unicode.GetString($decodeBuf)

    # The first part of the decoded type contains the Exchange server edition
    $strParts = $strDecodedType -split ";"

    if($strParts.Count -ne 3) {
        Write-Host "Array length mismatch. Exiting"
        return -1
    }

    # Make sure this is a valid edition - we then add the edition string back into the AD query datastore - we're going to save
    # the datastore with the edition in it
    [int]$nEdition = [convert]::ToInt32($strParts[0], 16)


    if ([enum]::GetValues([ExchangeEditions]) -contains $nEdition) {
        return ($nEdition -as [ExchangeEditions])
    }
    else {
        return ( "Unknown(" + $nEdition + ")")
    }
}


function GetExchangeInfo {
	$exchangeServerAttributes = "name", "objectGUID", "msexchproductid", "msexchcurrentserverroles", "type", "msexchserversite", "usncreated", "ADsPath", "msexchversion", "serialnumber", "msexchserverrole"
	$exchangeServers = SearchAD -searchAttributes $exchangeServerAttributes -searchFilter "(objectCategory=msExchExchangeServer)" -useNamingContext

	if ($exchangeServers) {
        # Parse Exchange Edition
        foreach ($srv in $exchangeServers) {
            $ExEditionDetails = DecodeExchangeEdition($srv.type)
            Add-Member -InputObject $srv -MemberType NoteProperty -Name "ExchangeEdition" -Value $ExEditionDetails

            $intValRoles = $srv.msexchcurrentserverroles
            $ExchangeRoles = @{2 = "Mailbox" ; 4 = "ClientAccess" ; 16 = "UnifiedMessaging" ; 32 = "HubTransport" ; 64 = "EdgeTransport" }
            $ExchServerRoles = $ExchangeRoles.Keys | where { $_ -band $intValRoles } | foreach { $ExchangeRoles.Get_Item($_) }
            $InstalledRoles = $ExchServerRoles -join ' | '
            Add-Member -InputObject $srv -MemberType NoteProperty -Name "ExchangeCurrentRoles" -Value $InstalledRoles
        }

		$exchangeServers | export-csv $OutputFile7 -notypeinformation -Encoding UTF8
	}
	
	$activeSyncDeviceAttributes = "name", "objectGUID", "ADsPath", "description", "whenChanged", "whenCreated","msExchDeviceEASVersion", "msExchDeviceFriendlyName", "msExchDeviceID", "msExchDeviceIMEI", "msExchDeviceMobileOperator", "msExchDeviceModel", "msExchDeviceOS", "msExchDeviceOSLanguage", "msExchDeviceTelephoneNumber", "msExchDeviceType", "msExchLastExchangeChangedTime", "msExchLastUpdateTime"
	$activeSyncDevices = SearchAD -searchAttributes $activeSyncDeviceAttributes -searchFilter "(objectClass=msExchActiveSyncDevice)"
	if ($activeSyncDevices) {
		$activeSyncDevices | export-csv $OutputFile8 -notypeinformation -Encoding UTF8
	}
	
	if ($Verbose){
		LogText "Active Sync Devices:"
		$activeSyncDevices | Group-Object msExchDeviceType | Select Name,Count | Sort Count -desc | ft -autosize | out-string | LogText
	}
}

function GetMoreRecentDate {
	Param(
    [string]$Date1 = "",
	[string]$Date2 = "")

	if ($Date1 -gt $Date2){
		return $Date1
	}
	else {
		return $Date2
	}
}

function CountItems {
	Param(
    $InputObject)

	if (-not $InputObject){
		return 0
	}
	elseif (Get-Member -inputobject $InputObject -name "Count" -Membertype Properties) {
		return $InputObject.Count
	}
	else {
		return 1
	}
}

function Get-ADDetails {
	try {
		InitialiseLogFile
		LogEnvironmentDetails
		
		LogProgress "Getting Domain Info"
		GetDomainInfo
		
		LogProgress "Getting User Info"
		GetUserInfo
		
		LogProgress "Getting Device Info"
		GetDeviceInfo

		LogProgress "Getting Exchange Info"
		GetExchangeInfo
		
		LogProgress "Complete"
	}
	catch {
		LogLastException
	}
}

Get-ADDetails
