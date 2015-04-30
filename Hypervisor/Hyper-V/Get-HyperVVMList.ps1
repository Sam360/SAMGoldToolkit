 ##########################################################################
 # 
 # Get-HyperVVMList
 # SAM Gold Toolkit
 # Original Source: Sam360
 #
 ##########################################################################
 
Param(
    [alias("server")]
    $HyperVServer = $env:COMPUTERNAME,
    [alias("o1")]
    $OutputFile = "HyperVExport" + $HyperVServer + ".csv")

function Get-HyperVVMList {
	$VMRecordList = @()

	# Get all virtual machine objects on the server in question
	$VMs = gwmi -namespace root\virtualization Msvm_ComputerSystem -computername $HyperVServer -filter "Caption = 'Virtual Machine'" 
 
	# Go over each of the virtual machines
	foreach ($VM in [array] $VMs) 
	{

		$VMRecord = New-Object -TypeName System.Object

		# Add Most important Values
		$VMRecord | Add-Member -MemberType NoteProperty -Name "FQDN" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "OSName" -Value ""
		$VMRecord | Add-Member -MemberType NoteProperty -Name "HyperV Name" -Value $VM.ElementName
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
		$Kvp = gwmi -namespace root\virtualization -query $query -computername $HyperVServer

		# Converting XML to Object
		foreach($StrDataItem in $Kvp.GuestIntrinsicExchangeItems)
		{

			$XmlDataItem = [xml]($StrDataItem)
			$AttributeName = $XmlDataItem.Instance.Property | ?{$_.Name -eq "Name"}
			$AttributeValue = $XmlDataItem.Instance.Property | ?{$_.Name -eq "Data"}

			switch -exact ($AttributeName.Value)
					{
				"FullyQualifiedDomainName"	{$VMRecord.FQDN = $AttributeValue.Value} 
				"OSName"      			{$VMRecord.OSName = $AttributeValue.Value}
				"OSVersion"      		{$VMRecord.OSVersion = $AttributeValue.Value}
				"CSDVersion"      		{$VMRecord.CSDVersion = $AttributeValue.Value}
				"ProductType"      		{$VMRecord.ProductType = $AttributeValue.Value}
				"NetworkAddressIPv4"      	{$VMRecord.NetworkAddressIPv4 = $AttributeValue.Value}
				"NetworkAddressIPv6"      	{$VMRecord.NetworkAddressIPv6 = $AttributeValue.Value}		
					}
		}

		$VMRecordList += $VMRecord
	}

	$VMRecordList | export-csv $OutputFile -notypeinformation
}

Get-HyperVVMList
