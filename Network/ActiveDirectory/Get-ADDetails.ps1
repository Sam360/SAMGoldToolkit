 ##########################################################################
 # 
 # 	Get-ADDetails
 # 	SAM Gold Toolkit
 #	Original Source: Sam360
 #
 ##########################################################################
 
 Param(
    [alias("o1")]
    $OutputFile1 = "ADDomains.csv",
	[alias("o2")]
    $OutputFile2 = "ADDomainTrusts.csv",
	[alias("o3")]
    $OutputFile3 = "ADUsers.csv",
	[alias("o4")]
    $OutputFile4 = "ADDevices.csv",
	[alias("r")]
    [string]$SearchRoot = "")
 

function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Output "Computer Name:		$($env:COMPUTERNAME)" #-ForegroundColor Magenta
	Write-Output "User Name:			$($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:		$($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:		$($host.Version.Major)"
	Write-Output "PowerShell Version:	$($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size:	$($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:			$($PSVersionTable.CLRVersion)"
}

function LogProgress($progressDescription){
    $output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
    write-output $output
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

	($searchAttributes | %{$objSearcher.PropertiesToLoad.Add($_)}) | out-null

	$objSearcher.FindAll() | % {
	    $pso = New-Object PSObject
		$value = ""
	    $_.Properties.GetEnumerator() | % {
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
			else {
				$value = ($_.Value | foreach {$_})
			}
	        Add-Member -InputObject $pso -MemberType NoteProperty -Name $_.Name -Value $value
	    } 
	    $searchResults = $searchResults + $pso
	}

	return $searchResults | select-object $searchAttributes
}

function GetDomainInfo {
	#Get a list of domains in the forest
	$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
	$forest.Domains | export-csv $OutputFile1 -notypeinformation -Encoding UTF8
	
	$domainTrustAttributes = "dnsroot", "ncname", "NETBIOSName", "description", "whenChanged", "whenCreated"
	$domainTrusts = SearchAD -searchAttributes $domainTrustAttributes -searchFilter "(NETBIOSName=*)" -useNamingContext
	$domainTrusts | export-csv $OutputFile2 -notypeinformation -Encoding UTF8
}

function Get-ADDetails {
	try {
		LogEnvironmentDetails
		
		LogProgress "Getting Domain Info"
		GetDomainInfo
		
		LogProgress "Getting User Info"
		$userAttributes = "sAMAccountName", "objectSid", "objectGUID", "displayName", "departmentNumber", "company", "department", "distinguishedName", "lastLogon", "lastLogonTimestamp", "logonCount", "mail", "telephoneNumber", "physicalDeliveryOfficeName", "description", "whenChanged", "whenCreated"
		$userList = SearchAD -searchAttributes $userAttributes -searchFilter "(&(objectCategory=person)(objectClass=user))" 
		$userList  | export-csv $OutputFile3 -notypeinformation -Encoding UTF8
		
		LogProgress "Getting Device Info"
		$deviceAttributes = "name", "objectSid", "objectGUID", "operatingSystem", "operatingSystemVersion", "operatingSystemServicePack", "lastLogon", "lastLogonTimeStamp", "ADsPath", "location", "dNSHostName", "description", "whenChanged", "whenCreated"
		$deviceList = SearchAD -searchAttributes $deviceAttributes -searchFilter "(objectClass=computer)" 
		$deviceList | export-csv $OutputFile4 -notypeinformation -Encoding UTF8
		
		LogProgress "Complete"
	}
	catch {
		LogLastException
	}
}

Get-ADDetails
