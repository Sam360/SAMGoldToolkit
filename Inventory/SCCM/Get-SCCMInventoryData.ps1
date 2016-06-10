 ##########################################################################
 #
 # Get-SCCMInventoryData
 #
 ##########################################################################

 Param(
	[alias("server")]
	$DatabaseServer = $env:computerName,
	[alias("database")]
	$DatabaseName = "CM_P01",
	[alias("o1")]
	$OutputFile1 = "Devices.csv",
	[alias("o2")]
	$OutputFile2 = "Software.csv",
	$UserName,
	$Password,
	[ValidateSet("AllData","DeviceData","SoftwareData")]
	$RequiredData = "AllData",
	$PortNumber = "0",
	[ValidateSet("2007","2012")]
	$SCCMVersion = "2012")
	
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
	Write-Output "Username Parameter:       $UserName"
	Write-Output "Server Parameter:         $DatabaseServer"
	Write-Output "Database Parameter:       $DatabaseName"
	Write-Output "Required Data:            $RequiredData"
}

function LogProgress($progressDescription){
	$output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
	write-output $output
}

function GetConnectionString {
	$connectionString = ""
	if ($PortNumber -ne "0") {
		$connectionString = "Data Source=$DatabaseServer,$PortNumber;Network Library=DBMSSOCN;Initial Catalog=$DatabaseName; "
	}
	else {
		$connectionString = "Server=$DatabaseServer; Database=$DatabaseName; "
	}
	
	if ($UserName){
		$connectionString += "User Id=$UserName; Password=$Password;"
	}
	else {
		$connectionString += "Trusted_Connection=True;"
	}
	return $connectionString
}

function Invoke-SQL {
    param(
        [string] $sqlCommand = $(throw "Please specify a query."),
		[string] $resultsFilePath = $(throw "Please specify a file to save results in.")
      )
	
	If (Test-Path -path $resultsFilePath) {
		Remove-Item $resultsFilePath }
	$fileWriter = New-Object System.IO.StreamWriter $resultsFilePath

	$connectionString = GetConnectionString
	"Connection String: $connectionString"
	$connection = new-object system.data.SqlClient.SQLConnection(GetConnectionString)
	$command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
	$connection.Open()
	
	$reader = $command.ExecuteReader()
	
	# Write the header to file
	for ($columnCounter = 0; $columnCounter -lt $reader.FieldCount; $columnCounter++) {
		$fileWriter.Write($reader.GetName($columnCounter))
		
		if ($columnCounter -lt $reader.FieldCount - 1) {
			$fileWriter.Write(",")
		}
	}
	$fileWriter.Write("`r`n")
	
	# Write the data to file
	while ($reader.Read())
	{
		for ($columnCounter = 0; $columnCounter -lt $reader.FieldCount; $columnCounter++) {
			$cellValue = $reader.GetValue($columnCounter)
			if ($cellValue -is [string]) {
				$fileWriter.Write("`"" + $reader.GetValue($columnCounter).Replace("`"", "`"`"") + "`"")
			}
			else {
				$fileWriter.Write($reader.GetValue($columnCounter).ToString())
			}
			
			if ($columnCounter -lt $reader.FieldCount - 1) {
				$fileWriter.Write(",")
			}
		}
		$fileWriter.Write("`r`n")
	}
	
	$reader.Close()
    $connection.Close()
	$fileWriter.Close()
}

function Get-SCCMInventoryData {
	try {
		LogEnvironmentDetails
		
		if ($RequiredData -eq "DeviceData" -or $RequiredData -eq "AllData") {
			if ($SCCMVersion -eq "2007") {
				Invoke-SQL -SQLCommand $sqlCommandDevices2007 -ResultsFilePath $OutputFile1 
			}
			else {
				Invoke-SQL -SQLCommand $sqlCommandDevices2012 -ResultsFilePath $OutputFile1
			}
		}
		
		if ($RequiredData -eq "SoftwareData" -or $RequiredData -eq "AllData") {
			if ($SCCMVersion -eq "2007") {
				Invoke-SQL -SQLCommand $sqlCommandSoftware2007 -ResultsFilePath $OutputFile2 
			}
			else {
				Invoke-SQL -SQLCommand $sqlCommandSoftware2012 -ResultsFilePath $OutputFile2
			}
		}
	}
	catch{
		LogLastException
	}
}

$sqlCommandDevices2012 = @"
SELECT
	[System].[ResourceID] AS [SourceKey]
	,[System].[Client_Version0] AS [SystemClientVersion]
	,[ComputerSystem].[Name0] AS [ComputerSystemName]
	,[System].[Name0] AS [SystemName]
	,[System].[NetBios_Name0] AS [SystemNetBiosName]
	,[ComputerSystem].[Domain0] AS [Domain]
	,[System].[Resource_Domain_OR_Workgr0] AS [Resource_Domain_OR_Workgr0]
	,[System].[Distinguished_Name0] AS [Distinguished_Name0]
	,[Bios].[SerialNumber0] AS [BiosSerialNumber]  
	,[Bios].[ReleaseDate0] AS [BiosReleaseDate]
	,[NetworkAdapter].[MacAddress] AS [MacAddress]
	,[NetworkAdapter1].[IPAddress] AS [IPAddress]
	,CONVERT(varchar(19), [System].[Last_Logon_Timestamp0], 126) AS [LastLogon]
	,CONVERT(varchar(19), [WorkstationStatus].[LastHWScan], 126) AS [LastHWScan]
	,CONVERT(varchar(19), [SoftwareInventoryStatus].[LastScanDate], 126) AS [LastSWScan]
	,[OperatingSystem].[Caption0] AS [OperatingSystem]
	,[OperatingSystem].[SerialNumber0] AS [OperatingSystemSerialNumber]
	,[OperatingSystem].[CSDVersion0] AS [OperatingSystemServicePack]
	,[OperatingSystem].[InstallDate0] AS [OperatingSystemInstallDate]
	,[OperatingSystem].[Version0] AS [OperatingSystemVersion]
	,[OperatingSystem].[TotalVisibleMemorySize0] AS [PhysicalMemory]
	,[OperatingSystem].[TotalVirtualMemorySize0] AS [VirtualMemory] 
	,[Processor].[ProcessorCount] AS [ProcessorCount]
	,[Processor].[NumberOfCores] AS [CoreCount]
	,[ComputerSystem].[NumberOfProcessors0] AS [LogicalProcessorCount]
	,[Processor].[CpuType] AS [CpuType]
	,[ComputerSystem].[Model0] AS [Model]
	,[ComputerSystem].[Manufacturer0] AS [Manufacturer]
	,NULLIF([System].[Virtual_Machine_Host_Name0], '') AS [VirtualHostName]
	,[VirtualMachine].[PhysicalHostName0] AS [VirtualPhysicalHostName]
	,[VirtualMachine].[ResourceID] AS [VirtualResourceID]
	--,'vRSystem' AS [Source]
	--,[FullCollectionMembership].[CollectionID] --
FROM 
	[dbo].[v_R_System] AS [System] 
--INNER JOIN 
--	[dbo].[v_FullCollectionMembership] AS [FullCollectionMembership]
--ON
--	[System].[ResourceID] = [FullCollectionMembership].[ResourceID]
--AND
--	[FullCollectionMembership].[CollectionID] IN ('SMS00001')
--INNER JOIN 
LEFT OUTER JOIN
	(
		SELECT
			[ResourceID]
			,[Caption0]
			,[SerialNumber0]
			,[CSDVersion0]
			,[InstallDate0]
			,[Version0]
			,[TotalVisibleMemorySize0]
			,[TotalVirtualMemorySize0]
			,ROW_NUMBER() OVER (PARTITION BY [ResourceID] ORDER BY [GroupID] DESC) AS [OperatingSystemRow]
		FROM
			[dbo].[v_GS_Operating_System]
	) AS [OperatingSystem]
ON
	[System].[ResourceID] = [OperatingSystem].[ResourceID]
AND
	[OperatingSystem].[OperatingSystemRow] = 1
--INNER JOIN 
LEFT OUTER JOIN
	(
		SELECT
			[ResourceID]
			,[SerialNumber0]
			,[ReleaseDate0]
			,ROW_NUMBER() OVER (PARTITION BY [ResourceID] ORDER BY [GroupID] DESC) AS [BiosRow]
		FROM
			[dbo].[v_GS_PC_Bios]
	) AS [Bios]
ON
	[System].[ResourceID] = [Bios].[ResourceID]
AND
	[Bios].[BiosRow] = 1
--INNER JOIN 
LEFT OUTER JOIN
	(
		SELECT
			[ResourceID]
			,[Name0]
			,[Domain0]
			,[NumberOfProcessors0]
			,[Model0]
			,[Manufacturer0]
			,ROW_NUMBER() OVER (PARTITION BY [ResourceID] ORDER BY [GroupID] DESC) AS [ComputerSystemRow]
		FROM
			[dbo].[v_GS_Computer_System]
	) AS [ComputerSystem]
ON
	[System].[ResourceID] = [ComputerSystem].[ResourceID]
AND
	[ComputerSystem].[ComputerSystemRow] = 1
LEFT OUTER JOIN 
	(
		SELECT  
			[ResourceID]
			,COUNT([ResourceID]) AS [ProcessorCount]
			,SUM([NumberOfCores0]) AS [NumberOfCores]
			,MAX([Name0]) AS [CpuType]
		FROM
			[dbo].[v_GS_Processor]
		GROUP BY 
			[ResourceID]
	) AS [Processor]
ON
	[System].[ResourceID] = [Processor].[ResourceID]
LEFT OUTER JOIN
	[dbo].[V_GS_Workstation_Status] AS [WorkstationStatus]
ON
	[WorkstationStatus].[ResourceID] = [System].[ResourceID]
LEFT OUTER JOIN
	[dbo].[V_GS_LastSoftwareScan] AS [SoftwareInventoryStatus]
ON
	[SoftwareInventoryStatus].[ResourceID] = [System].[ResourceID]
LEFT OUTER JOIN
	[dbo].[v_GS_Virtual_Machine] AS [VirtualMachine]
ON
	[VirtualMachine].[ResourceID] = [System].[ResourceID]
OUTER APPLY
	(
		SELECT DISTINCT
			[Network].[MACAddress0] + ';'
		FROM 
			[dbo].[v_GS_NETWORK_ADAPTER] AS [Network]
		WHERE 
			[Network].[ResourceID] = [System].[ResourceID]
		AND
			[Network].[MACAddress0] IS NOT NULL 
		AND
			[Network].[MACAddress0] NOT IN ('00:00:00:00:00:00','33:50:6F:45:30:30','50:50:54:50:30:30')
		ORDER BY 
			[Network].[MACAddress0] + ';'
		FOR XML PATH ('')
	) AS [NetworkAdapter] (MacAddress)
OUTER APPLY
	(
		SELECT DISTINCT
			[NetworkConfig].[IPAddress0] + ';'
		FROM 
			[dbo].[v_GS_NETWORK_ADAPTER_CONFIGURATION] AS [NetworkConfig]
		WHERE 
			[NetworkConfig].[ResourceID] = [System].[ResourceID]
		AND
			[NetworkConfig].[IPAddress0] IS NOT NULL 
		AND
			[NetworkConfig].[IPAddress0] NOT IN ('0.0.0.0')
		FOR XML PATH ('')
	) AS [NetworkAdapter1] (IPAddress);
"@

$sqlCommandSoftware2012 = @"
SELECT
	[SoftwareCollection].[DeviceSourceKey] AS [DeviceSourceKey]
	,CONVERT(varchar(19), GETDATE(), 126) AS [DiscoveryDate]
	,[SoftwareCollection].[InstalledLocation] AS [InstalledLocation]
	,[SoftwareCollection].[ProductName]  AS [SwName]
	,[SoftwareCollection].[ProductVersion]  AS [SwVersion]
	,[SoftwareCollection].[Publisher]  AS [SwPublisher]
	,[SoftwareCollection].[SoftwareCode] AS [SwSoftwareCode]
	,[SoftwareCollection].[PackageCode] AS [SwSoftwareId]
	,[SoftwareCollection].[IsOperatingSystem]
	,[SoftwareCollection].[Source]
FROM
	(
-- Select entries from the v_GS_Installed_Software View
		SELECT
			[Software].[ResourceID] AS [DeviceSourceKey]
			,[Software].[InstalledLocation0] AS [InstalledLocation]
			,COALESCE(NULLIF([Software].[ProductName0], ''), N'Unknown') AS [ProductName]
			,COALESCE(NULLIF([Software].[ProductVersion0], ''), N'Unknown') AS [ProductVersion]
			,COALESCE(NULLIF([Software].[Publisher0], ''), N'Unknown') AS [Publisher]
			,[Software].[SoftwareCode0] AS [SoftwareCode]
			,[Software].[PackageCode0] AS [PackageCode]
			,0 AS [IsOperatingSystem]
			,'v_GS_IS' AS [Source]
		FROM 
			[dbo].[v_GS_Installed_Software] AS [Software]
-- Apply this filter to remove software we don't care about at the present time to reduce the volume
		WHERE
			CHARINDEX('microsoft', [Software].[Publisher0]) > 0
		AND
			CHARINDEX('KB', [Software].[ProductName0]) +
			CHARINDEX('.NET Framework', [Software].[ProductName0]) +
			CHARINDEX('Update', [Software].[ProductName0]) +
			CHARINDEX('Service Pack', [Software].[ProductName0]) +
			CHARINDEX('Proof', [Software].[ProductName0]) +
			CHARINDEX('Components', [Software].[ProductName0]) +
			CHARINDEX('Tools', [Software].[ProductName0]) +
			CHARINDEX('MUI', [Software].[ProductName0]) +
			CHARINDEX('Redistributable', [Software].[ProductName0]) = 0
		UNION ALL
-- Select entries from the v_GS_Operating_System View
		SELECT
			[OperatingSystem].[ResourceID] AS [DeviceSourceKey]
			,NULL AS [InstalledLocation]
			,COALESCE(NULLIF([OperatingSystem].[Caption0], ''), N'Unknown OS') AS [ProductName]
			,COALESCE(NULLIF([OperatingSystem].[Version0], ''), N'Unknown') AS [ProductVersion]
			,CASE
				WHEN [OperatingSystem].[Caption0] LIKE N'%windows%'
				THEN N'Microsoft'
				ELSE N'Unknown'
			END AS [Publisher]
			,NULL AS [SoftwareCode]
			,NULL AS [PackageCode]
			,1 AS [IsOperatingSystem]
			,'v_GS_OS' AS [Source]
		FROM 
			[dbo].[v_GS_Operating_System] AS [OperatingSystem]
	) AS [SoftwareCollection]
INNER JOIN 
	[dbo].[v_FullCollectionMembership] AS [FullCollectionMembership]
ON
	[SoftwareCollection].[DeviceSourceKey] = [FullCollectionMembership].[ResourceID]
AND
	[FullCollectionMembership].[CollectionID] IN ('SMS00001')
GROUP BY
	[SoftwareCollection].[DeviceSourceKey]
	,[SoftwareCollection].[InstalledLocation]
	,[SoftwareCollection].[ProductName]
	,[SoftwareCollection].[ProductVersion]
	,[SoftwareCollection].[Publisher]
	,[SoftwareCollection].[SoftwareCode]
	,[SoftwareCollection].[PackageCode]
	,[SoftwareCollection].[IsOperatingSystem]
	,[SoftwareCollection].[Source]
"@

$sqlCommandDevices2007 = @"
SELECT
	[System].[ResourceID] AS [SourceKey]
	,[System].[Client_Version0] AS [SystemClientVersion]
	,[ComputerSystem].[Name0] AS [ComputerSystemName]
	,[System].[Name0] AS [SystemName]
	,[System].[NetBios_Name0] AS [SystemNetBiosName]
	,[ComputerSystem].[Domain0] AS [Domain]
	,[System].[Resource_Domain_OR_Workgr0] AS [Resource_Domain_OR_Workgr0]
	--/--,[System].[Distinguished_Name0] AS [Distinguished_Name0]
	,[Bios].[SerialNumber0] AS [BiosSerialNumber]  
	,[Bios].[ReleaseDate0] AS [BiosReleaseDate]
	,[NetworkAdapter].[MacAddress] AS [MacAddress]
	,[NetworkAdapter1].[IPAddress] AS [IPAddress]
	--/--,CONVERT(varchar(19), [System].[Last_Logon_Timestamp0], 126) AS [LastLogon]
	,CONVERT(varchar(19), [WorkstationStatus].[LastHWScan], 126) AS [LastHWScan]
	,CONVERT(varchar(19), [SoftwareInventoryStatus].[LastScanDate], 126) AS [LastSWScan]
	,[OperatingSystem].[Caption0] AS [OperatingSystem]
	--/--,[OperatingSystem].[SerialNumber0] AS [OperatingSystemSerialNumber]
	,[OperatingSystem].[LastBootUpTime0] AS [LastBootUpTime]
	,[OperatingSystem].[CSDVersion0] AS [OperatingSystemServicePack]
	,[OperatingSystem].[InstallDate0] AS [OperatingSystemInstallDate]
	,[OperatingSystem].[Version0] AS [OperatingSystemVersion]
	,[OperatingSystem].[TotalVisibleMemorySize0] AS [PhysicalMemory]
	,[OperatingSystem].[TotalVirtualMemorySize0] AS [VirtualMemory] 
	,[Processor].[ProcessorCount] AS [ProcessorCount]
	--/--,[Processor].[NumberOfCores] AS [CoreCount]
	,[Processor].[CpuIsMulticore]
	,[Processor].[CpuNormSpeed]
	,[Processor].[CpuProcessorType]
	,[Processor].[CpuDeviceID]
	,[ComputerSystem].[NumberOfProcessors0] AS [LogicalProcessorCount]
	,[Processor].[CpuType] AS [CpuType]
	,[ComputerSystem].[Model0] AS [Model]
	,[ComputerSystem].[Manufacturer0] AS [Manufacturer]
	--/--,NULLIF([System].[Virtual_Machine_Host_Name0], '') AS [VirtualHostName]
	--/--,[VirtualMachine].[PhysicalHostName0] AS [VirtualPhysicalHostName]
	--/--,[VirtualMachine].[ResourceID] AS [VirtualResourceID]
	--,'vRSystem' AS [Source]
	--,[FullCollectionMembership].[CollectionID] --
FROM 
	[dbo].[v_R_System] AS [System] 
--INNER JOIN 
--	[dbo].[v_FullCollectionMembership] AS [FullCollectionMembership]
--ON
--	[System].[ResourceID] = [FullCollectionMembership].[ResourceID]
--AND
--	[FullCollectionMembership].[CollectionID] IN ('SMS00001')
--INNER JOIN 
LEFT OUTER JOIN
	(
		SELECT
			[ResourceID]
			,[Caption0]
			--/--,[SerialNumber0]
			,[LastBootUpTime0]
			,[CSDVersion0]
			,[InstallDate0]
			,[Version0]
			,[TotalVisibleMemorySize0]
			,[TotalVirtualMemorySize0]
			,ROW_NUMBER() OVER (PARTITION BY [ResourceID] ORDER BY [GroupID] DESC) AS [OperatingSystemRow]
		FROM
			[dbo].[v_GS_Operating_System]
	) AS [OperatingSystem]
ON
	[System].[ResourceID] = [OperatingSystem].[ResourceID]
AND
	[OperatingSystem].[OperatingSystemRow] = 1
--INNER JOIN 
LEFT OUTER JOIN
	(
		SELECT
			[ResourceID]
			,[SerialNumber0]
			,[ReleaseDate0]
			,ROW_NUMBER() OVER (PARTITION BY [ResourceID] ORDER BY [GroupID] DESC) AS [BiosRow]
		FROM
			[dbo].[v_GS_PC_Bios]
	) AS [Bios]
ON
	[System].[ResourceID] = [Bios].[ResourceID]
AND
	[Bios].[BiosRow] = 1
--INNER JOIN 
LEFT OUTER JOIN
	(
		SELECT
			[ResourceID]
			,[Name0]
			,[Domain0]
			,[NumberOfProcessors0]
			,[Model0]
			,[Manufacturer0]
			,ROW_NUMBER() OVER (PARTITION BY [ResourceID] ORDER BY [GroupID] DESC) AS [ComputerSystemRow]
		FROM
			[dbo].[v_GS_Computer_System]
	) AS [ComputerSystem]
ON
	[System].[ResourceID] = [ComputerSystem].[ResourceID]
AND
	[ComputerSystem].[ComputerSystemRow] = 1
LEFT OUTER JOIN 
	(
		SELECT  
			[ResourceID]
			,COUNT([ResourceID]) AS [ProcessorCount]
			--/--,SUM([NumberOfCores0]) AS [NumberOfCores]
			,MAX([Name0]) AS [CpuType]
			,MAX([IsMulticore0]) AS [CpuIsMulticore]
			,MAX([NormSpeed0]) AS [CpuNormSpeed]
			,MAX([ProcessorType0]) AS [CpuProcessorType]
			,MAX([DeviceID0]) AS [CpuDeviceID]
		FROM
			[dbo].[v_GS_Processor]
		GROUP BY 
			[ResourceID]
	) AS [Processor]
ON
	[System].[ResourceID] = [Processor].[ResourceID]
LEFT OUTER JOIN
	[dbo].[V_GS_Workstation_Status] AS [WorkstationStatus]
ON
	[WorkstationStatus].[ResourceID] = [System].[ResourceID]
LEFT OUTER JOIN
	[dbo].[V_GS_LastSoftwareScan] AS [SoftwareInventoryStatus]
ON
	[SoftwareInventoryStatus].[ResourceID] = [System].[ResourceID]
--/--LEFT OUTER JOIN
--/--	[dbo].[v_GS_Virtual_Machine] AS [VirtualMachine]
--/--ON
--/--	[VirtualMachine].[ResourceID] = [System].[ResourceID]
OUTER APPLY
	(
		SELECT DISTINCT
			[Network].[MACAddress0] + ';'
		FROM 
			[dbo].[v_GS_NETWORK_ADAPTER] AS [Network]
		WHERE 
			[Network].[ResourceID] = [System].[ResourceID]
		AND
			[Network].[MACAddress0] IS NOT NULL 
		AND
			[Network].[MACAddress0] NOT IN ('00:00:00:00:00:00','33:50:6F:45:30:30','50:50:54:50:30:30')
		ORDER BY 
			[Network].[MACAddress0] + ';'
		FOR XML PATH ('')
	) AS [NetworkAdapter] (MacAddress)
OUTER APPLY
	(
		SELECT DISTINCT
			[NetworkConfig].[IPAddress0] + ';'
		FROM
			[dbo].v_GS_NETWORK_ADAPTER_CONFIGUR AS [NetworkConfig]
			--/--[dbo].[v_GS_NETWORK_ADAPTER_CONFIGURATION] AS [NetworkConfig]
		WHERE 
			[NetworkConfig].[ResourceID] = [System].[ResourceID]
		AND
			[NetworkConfig].[IPAddress0] IS NOT NULL 
		AND
			[NetworkConfig].[IPAddress0] NOT IN ('0.0.0.0')
		FOR XML PATH ('')
	) AS [NetworkAdapter1] (IPAddress);
"@

$sqlCommandSoftware2007 = @"
SELECT
	[SoftwareCollection].[DeviceSourceKey] AS [DeviceSourceKey]
	,CONVERT(varchar(19), GETDATE(), 126) AS [DiscoveryDate]
	,[SoftwareCollection].[InstalledLocation] AS [InstalledLocation]
	,[SoftwareCollection].[ProductName]  AS [SwName]
	,[SoftwareCollection].[ProductVersion]  AS [SwVersion]
	,[SoftwareCollection].[Publisher]  AS [SwPublisher]
	,[SoftwareCollection].[SoftwareCode] AS [SwSoftwareCode]
	,[SoftwareCollection].[PackageCode] AS [SwSoftwareId]
	,[SoftwareCollection].[IsOperatingSystem]
	,[SoftwareCollection].[Source]
FROM
	(
-- Select entries from the v_GS_Installed_Software View
		SELECT
			[Software].[ResourceID] AS [DeviceSourceKey]
			,[Software].[InstalledLocation0] AS [InstalledLocation]
			,COALESCE(NULLIF([Software].[ProductName0], ''), N'Unknown') AS [ProductName]
			,COALESCE(NULLIF([Software].[ProductVersion0], ''), N'Unknown') AS [ProductVersion]
			,COALESCE(NULLIF([Software].[Publisher0], ''), N'Unknown') AS [Publisher]
			,[Software].[SoftwareCode0] AS [SoftwareCode]
			--/--,[Software].[PackageCode0] AS [PackageCode]
			,NULL AS [PackageCode]
			,0 AS [IsOperatingSystem]
			,'v_GS_IS' AS [Source]
		FROM 
			[dbo].[v_GS_Installed_Software] AS [Software]
-- Apply this filter to remove software we don't care about at the present time to reduce the volume
		WHERE
			CHARINDEX('microsoft', [Software].[Publisher0]) > 0
		AND
			CHARINDEX('KB', [Software].[ProductName0]) +
			CHARINDEX('.NET Framework', [Software].[ProductName0]) +
			CHARINDEX('Update', [Software].[ProductName0]) +
			CHARINDEX('Service Pack', [Software].[ProductName0]) +
			CHARINDEX('Proof', [Software].[ProductName0]) +
			CHARINDEX('Components', [Software].[ProductName0]) +
			CHARINDEX('Tools', [Software].[ProductName0]) +
			CHARINDEX('MUI', [Software].[ProductName0]) +
			CHARINDEX('Redistributable', [Software].[ProductName0]) = 0
		UNION ALL
-- Select entries from the v_GS_Operating_System View
		SELECT
			[OperatingSystem].[ResourceID] AS [DeviceSourceKey]
			,NULL AS [InstalledLocation]
			,COALESCE(NULLIF([OperatingSystem].[Caption0], ''), N'Unknown OS') AS [ProductName]
			,COALESCE(NULLIF([OperatingSystem].[Version0], ''), N'Unknown') AS [ProductVersion]
			,CASE
				WHEN [OperatingSystem].[Caption0] LIKE N'%windows%'
				THEN N'Microsoft'
				ELSE N'Unknown'
			END AS [Publisher]
			,NULL AS [SoftwareCode]
			,NULL AS [PackageCode]
			,1 AS [IsOperatingSystem]
			,'v_GS_OS' AS [Source]
		FROM 
			[dbo].[v_GS_Operating_System] AS [OperatingSystem]
	) AS [SoftwareCollection]
INNER JOIN 
	[dbo].[v_FullCollectionMembership] AS [FullCollectionMembership]
ON
	[SoftwareCollection].[DeviceSourceKey] = [FullCollectionMembership].[ResourceID]
AND
	[FullCollectionMembership].[CollectionID] IN ('SMS00001')
GROUP BY
	[SoftwareCollection].[DeviceSourceKey]
	,[SoftwareCollection].[InstalledLocation]
	,[SoftwareCollection].[ProductName]
	,[SoftwareCollection].[ProductVersion]
	,[SoftwareCollection].[Publisher]
	,[SoftwareCollection].[SoftwareCode]
	,[SoftwareCollection].[PackageCode]
	,[SoftwareCollection].[IsOperatingSystem]
	,[SoftwareCollection].[Source]
"@

Get-SCCMInventoryData