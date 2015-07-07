 ##########################################################################
 # 
 # Get-HyperVVMList
 # SAM Gold Toolkit
 # Original Source: Jon Mulligan (Sam360)
 #
 ##########################################################################
 
<#
.SYNOPSIS
Retrieves physical host, virtual machine and virtual machine migration data from a Hyper-V server

.DESCRIPTION
The Get-HyperVVMList script queries a single HyperV server and produces 4 CSV files
	1)    HyperVExportBasic.csv - One record per virtual machine including fields like 
	      VM name, IP, OS, Enabled state, Physical host name etc. The data is retrieved through WMI
    2)    HyperVExportGuests.csv - One record per virtual machine. The data is retrieved through 
	      PowerShell and requires minimum Windows Server 2012 R2
	3)    HyperVExportHosts.csv - One record per Hyper-V server. The data is retrieved through 
	      PowerShell and requires minimum Windows Server 2012 R2
	4)    HyperVExportGuestMigration.csv - One record per migration event. The data is retrieved through 
	      PowerShell and requires minimum Windows Server 2012 R2

    Files are written to current working directory

.PARAMETER Server 
Host name of Hyper-V server to scan

.EXAMPLE
Get all guest, host and migration info from Hyper-V server 'Defiant'. 
Get-HyperVVMList –HyperVServer Defiant

.NOTES
File 2,3 & 4 will only contain data when querying Hyper-V servers with Windows Server 2012 R2 onwards installed.
#>

Param(
    [alias("server")]
    $HyperVServer = $env:COMPUTERNAME,
    [alias("o1")]
    $OutputFile1 = "HyperVExport" + $HyperVServer + "Basic.csv",
    [alias("o2")]
    $OutputFile2 = "HyperVExport" + $HyperVServer + "Guests.csv",
    [alias("o3")]
    $OutputFile3 = "HyperVExport" + $HyperVServer + "Hosts.csv",
    [alias("o4")]
    $OutputFile4 = "HyperVExport" + $HyperVServer + "GuestMigration.csv",
	[ValidateSet("AllData","BasicData","DetailedData")] 
	$RequiredData = "DetailedData",
	[switch]
	$Verbose)

function LogLastException()
{
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

function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Output "Computer Name:            $($env:COMPUTERNAME)"
	Write-Output "User Name:                $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:          $($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:          $($host.Version.Major)"
	Write-Output "PowerShell Version:       $($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size:     $($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:              $($PSVersionTable.CLRVersion)"
	Write-Output "Server:                   $HyperVServer"
	Write-Output "Required Data:            $RequiredData"
}

function LogProgress($progressDescription){
    $output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
    write-output $output
}

function Get-HyperVVMList1 {
	$VMRecordList = @()

	$hyperVNamespace = "root\virtualization"
	$hyperVClass = gwmi -Class 'Msvm_ComputerSystem' -List -Namespace $hyperVNamespace -computername $HyperVServer -ErrorAction SilentlyContinue
	if ($hyperVClass -eq $null) {
		$hyperVNamespace = "root\virtualization\v2"
		$hyperVClass = gwmi -Class 'Msvm_ComputerSystem' -List -Namespace $hyperVNamespace -computername $HyperVServer -ErrorAction SilentlyContinue
		if ($hyperVClass -eq $null) {
			LogProgress "Unable to locate required WMI namespace"
			return
		}
	}

	# Get all virtual machine objects on the server in question
	$VMs = gwmi -namespace $hyperVNamespace Msvm_ComputerSystem -computername $HyperVServer -filter "Caption = 'Virtual Machine'" 
 
	# Go over each of the virtual machines
	foreach ($VM in [array] $VMs) 
	{
		if ($Verbose){
			LogProgress "Retrieving Basic details for $($VM.ElementName) ($($VM.Name))"
		}
		
		$VMRecord = New-Object -TypeName System.Object

		# Add Most important Values
		$VMRecord | Add-Member -MemberType NoteProperty -Name "FQDN" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "OSName" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "HyperV_Name" -Value $VM.ElementName
		$VMRecord | Add-Member -MemberType NoteProperty -Name "EnabledState" -Value ""

		# Add base values
		$VMRecord | Add-Member -MemberType NoteProperty -Name "Host" -Value $HyperVServer
		$VMRecord | Add-Member -MemberType NoteProperty -Name "GUID" -Value $VM.Name
		$VMRecord | Add-Member -MemberType NoteProperty -Name "Description" -Value $VM.Description
		$VMRecord | Add-Member -MemberType NoteProperty -Name "EnabledStateID" -Value $VM.EnabledState
		$VMRecord | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $VM.InstallDate
		$VMRecord | Add-Member -MemberType NoteProperty -Name "OnTimeInMilliseconds" -Value $VM.OnTimeInMilliseconds  
		$VMRecord | Add-Member -MemberType NoteProperty -Name "TimeOfLastStateChange" -Value $VM.TimeOfLastStateChange

		# Add xml values
		$VMRecord | Add-Member -MemberType NoteProperty -Name "OSVersion" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "CSDVersion" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "ProductType" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "NetworkAddressIPv4" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "NetworkAddressIPv6" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "OSEditionId" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "ProcessorArchitecture" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "SuiteMask" -Value ""

		switch ($VM.EnabledState) 
		{
			0		{$VMRecord.EnabledState = "Unknown"}
			2		{$VMRecord.EnabledState = "Enabled"}
			3		{$VMRecord.EnabledState = "Disabled"}
			32768	{$VMRecord.EnabledState = "Paused"}
			3276	{$VMRecord.EnabledState = "Suspended"}
			32770	{$VMRecord.EnabledState = "Starting"}
			32771	{$VMRecord.EnabledState = "Snapshotting"}
			32773	{$VMRecord.EnabledState = "Saving"}
			32774	{$VMRecord.EnabledState = "Stopping"}
			32776	{$VMRecord.EnabledState = "Pausing"}
			32777	{$VMRecord.EnabledState = "Resuming"}
			default	{$VMRecord.EnabledState = "Unknown"}
		  }


		# Get the KVP Object
		$query = "Associators of {$VM} Where AssocClass=Msvm_SystemDevice ResultClass=Msvm_KvpExchangeComponent"
		try {
			$Kvp = gwmi -namespace $hyperVNamespace -query $query -computername $HyperVServer
		}
		catch {
			LogProgress "Error retrieving info for VM $($VM.Name)"
			$Kvp = $null
		}

		# Converting XML to Object
		foreach($StrDataItem in $Kvp.GuestIntrinsicExchangeItems)
		{

			$XmlDataItem = [xml]($StrDataItem)
			$AttributeName = $XmlDataItem.Instance.Property | ?{$_.Name -eq "Name"}
			$AttributeValue = $XmlDataItem.Instance.Property | ?{$_.Name -eq "Data"}

			switch -exact ($AttributeName.Value)
			{
				"FullyQualifiedDomainName"	{$VMRecord.FQDN = $AttributeValue.Value} 
				"OSName"      				{$VMRecord.OSName = $AttributeValue.Value}
				"OSVersion"      			{$VMRecord.OSVersion = $AttributeValue.Value}
				"CSDVersion"      			{$VMRecord.CSDVersion = $AttributeValue.Value}
				"ProductType"      			{$VMRecord.ProductType = $AttributeValue.Value}
				"NetworkAddressIPv4"      	{$VMRecord.NetworkAddressIPv4 = $AttributeValue.Value}
				"NetworkAddressIPv6"      	{$VMRecord.NetworkAddressIPv6 = $AttributeValue.Value}
				"OSEditionId"      			{$VMRecord.OSEditionId = $AttributeValue.Value}
				"ProcessorArchitecture"     {$VMRecord.ProcessorArchitecture = $AttributeValue.Value}
				"SuiteMask"      			{$VMRecord.SuiteMask = $AttributeValue.Value}		
			}
		}

		$VMRecordList += $VMRecord
	}

	$VMRecordList | export-csv $OutputFile1 -notypeinformation -Encoding UTF8
}

function Get-HyperVVMList2 {
	if ((Get-Module -ListAvailable -Name "Hyper-V") -eq $null) {
		LogProgress "Hyper PowerShell module not available"
		return
	}

	Import-Module "Hyper-V"

	Get-VM -ComputerName $HyperVServer | export-csv $OutputFile2 -notypeinformation -Encoding UTF8
	Get-VMHost -ComputerName $HyperVServer | export-csv $OutputFile3 -notypeinformation -Encoding UTF8
}

function Get-HyperVVMMigrationInfo {
	
LogProgress "Retrieving HyperV VM Events"
	
	$AllVMEvents = Get-WinEvent -LogName "Microsoft-Windows-Hyper-V-VMMS-Admin" -ComputerName $HyperVServer
	if($Verbose){
		LogProgress "Retrieved $($AllVMEvents.Count) HyperV VM Events"
	}

	$AllVMMigrationEvents = $AllVMEvents | where {$_.Id -like "2041*"} 
	if($Verbose){
		LogProgress "Retrieved $($AllVMMigrationEvents.Count) HyperV VM Migration Events"
	}

	$AllVMMigrationEvents | export-csv $OutputFile4 -notypeinformation -Encoding UTF8
}

function Get-HyperVVMList {
	try {
		LogEnvironmentDetails

		LogProgress "Getting basic HyperV Guest Info (WMI)"
		Get-HyperVVMList1
		
		if ($RequiredData -eq "DetailedData" -or $RequiredData -eq "AllData") {
			LogProgress "Getting detailed HyperV Info (PowerShell)"
			Get-HyperVVMList2
		}

		if ($RequiredData -eq "AllData") {
			LogProgress "Getting HyperV Guest Migration Info (PowerShell)"
			Get-HyperVVMMigrationInfo
		}
	}
	catch {
		LogLastException
	}
}

Get-HyperVVMList
