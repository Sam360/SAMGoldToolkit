 ##########################################################################
 #
 # Get-MAPInventoryData
 # SAM Gold Toolkit
 # Original Source: Sam360, Microsoft SAM Workspace Discovery Tool
 #
 ##########################################################################

 Param(
	[alias("database")]
	[alias("MapDatabaseName")]
	$DatabaseName = "MAP",
	[alias("o1")]
	$OutputFile1 = "ADDiscoveredDevices.csv",
	[alias("o2")]
	$OutputFile2 = "Devices.csv",
	[alias("o3")]
	$OutputFile3 = "SQLServerInventory1.csv",
	[alias("o4")]
	$OutputFile4 = "ExchangeInventory.csv",
	[alias("o5")]
	$OutputFile5 = "VMwareGuests.csv",
	[alias("o6")]
	$OutputFile6 = "VMwareHosts.csv",
	[alias("o7")]
	$OutputFile7 = "NetworkAdapters.csv",
	[alias("o8")]
	$OutputFile8 = "OEMData.csv",
	[alias("o9")]
	$OutputFile9 = "Processors.csv",
	[alias("o10")]
	$OutputFile10 = "Products1.csv",
	[alias("o11")]
	$OutputFile11 = "Products2.csv",
	[alias("o12")]
	$OutputFile12 = "Services.csv",
	[alias("o13")]
	$OutputFile13 = "SoftwareLicensingProducts.csv",
	[alias("o14")]
	$OutputFile14 = "MSClusters.csv",
	[alias("o15")]
	$OutputFile15 = "SQLServerInventory2.csv",
	[alias("o16")]
	$OutputFile16 = "MAPDBNames.csv",
	[ValidateSet("DBList","AllData","DeviceData","SoftwareData")]
	$RequiredData = "AllData")

function LogLastException()
{
    $currentException = $Error[0].Exception;
    $exceptionCounter = 1

    while ($currentException)
    {
        write-output "Exception $exceptionCounter:            $currentException"
        write-output "Exception $exceptionCounter Data:       $($currentException.Data)"
        write-output "Exception $exceptionCounter HelpLink:   $($currentException.HelpLink)"
        write-output "Exception $exceptionCounter HResult:    $($currentException.HResult)"
        write-output "Exception $exceptionCounter Message:    $($currentException.Message)"
        write-output "Exception $exceptionCounter Source:     $($currentException.Source)"
        write-output "Exception $exceptionCounter StackTrace: $($currentException.StackTrace)"
        write-output "Exception $exceptionCounter TargetSite: $($currentException.TargetSite)"

        $currentException = $currentException.InnerException
        $exceptionCounter++
    }
}

function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Output "Computer Name:			$($env:COMPUTERNAME)"
	Write-Output "User Name:				$($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:			$($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:			$($host.Version.Major)"
	Write-Output "PowerShell Version:		$($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size:		$($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:				$($PSVersionTable.CLRVersion)"
	Write-Output "Database Parameter:		$DatabaseName"
	Write-Output "Required Data:			$RequiredData"
}

function LogProgress($progressDescription){
    $output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
    write-output $output
}

function GetConnectionString {
	# Start the local db instance
	$sqlLocalDBOutput = [string] (& sqllocaldb start MAPToolkit 2>&1)

	# Get info for the db instance
	$sqlLocalDBOutput = [string] (& sqllocaldb info MAPToolkit 2>&1)
	$sqlLocalDBInstanceInfo = $sqlLocalDBOutput.Split(“`n ”)

	# Search for the named pipe definition
	foreach ($infoString in $sqlLocalDBInstanceInfo) {
		if ($infoString.StartsWith("np:\\")){
			# We've found the named pipe definition
			return "Server=$infoString;Trusted_Connection=True"
		}
	}

	# unable to find
	return ""
}

function Invoke-SQL {
    param(
        [string] $sqlCommand = $(throw "Please specify a query."),
		[string] $resultsFilePath = $(throw "Please specify a file to save results in."),
		[string] $connectionString = $(throw "Please specify a connection string.")
      )
	
	Remove-Item $resultsFilePath -ErrorAction SilentlyContinue
	$fileWriter = New-Object System.IO.StreamWriter $resultsFilePath

    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
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

function VerifyDBExists {
	try {
		$sqlLocalDBOutput = [string] (& sqllocaldb info 2>&1)
	}
	catch {
		LogProgress "Unable to locate local SQL Instance"
		return $False
	}
	
	$dbNames = $sqlLocalDBOutput.Split(“`n ”)

	return $dbNames -contains "MAPToolkit"
}

function Get-MAPInventoryData {
	try {
		LogEnvironmentDetails

		if (-not (VerifyDBExists)) {
			LogProgress "Unable to find MAP database"
			return
		}

		$connectionString = GetConnectionString
		Write-Output "ConnectString: $connectionString"
		
		if ($RequiredData -eq "DBList") {
			Invoke-SQL -SQLCommand $sqlCommandGetMAPDBList -ResultsFilePath $OutputFile16 -connectionString $connectionString
			return
		}
			
		if ($RequiredData -eq "DeviceData" -or $RequiredData -eq "AllData") {
			LogProgress "Getting AdDiscoveredDevices Data"
			Invoke-SQL -SQLCommand $sqlCommand1 -ResultsFilePath $OutputFile1 -connectionString $connectionString

			LogProgress "Getting Devices Data"
			Invoke-SQL -SQLCommand $sqlCommand2 -ResultsFilePath $OutputFile2 -connectionString $connectionString

			LogProgress "Getting VMware Guests Data"
			Invoke-SQL -SQLCommand $sqlCommand5 -ResultsFilePath $OutputFile5 -connectionString $connectionString

			LogProgress "Getting VMware Hosts Data"
			Invoke-SQL -SQLCommand $sqlCommand6 -ResultsFilePath $OutputFile6 -connectionString $connectionString

			LogProgress "Getting Network Adapter Data"
			Invoke-SQL -SQLCommand $sqlCommand7 -ResultsFilePath $OutputFile7 -connectionString $connectionString

			LogProgress "Getting Processor Data"
			Invoke-SQL -SQLCommand $sqlCommand9 -ResultsFilePath $OutputFile9 -connectionString $connectionString

			LogProgress "Getting Cluster Data"
			Invoke-SQL -SQLCommand $sqlCommand14 -ResultsFilePath $OutputFile14 -connectionString $connectionString
		}
		
		if ($RequiredData -eq "SoftwareData" -or $RequiredData -eq "AllData") {
			LogProgress "Getting SQL Server Inventory Data"
			Invoke-SQL -SQLCommand $sqlCommand3 -ResultsFilePath $OutputFile3 -connectionString $connectionString
			
			LogProgress "Getting Exchange Inventory Data"
			Invoke-SQL -SQLCommand $sqlCommand4 -ResultsFilePath $OutputFile4 -connectionString $connectionString

			LogProgress "Getting OEM Data"
			Invoke-SQL -SQLCommand $sqlCommand8 -ResultsFilePath $OutputFile8 -connectionString $connectionString

			LogProgress "Getting Products Data"
			Invoke-SQL -SQLCommand $sqlCommand10 -ResultsFilePath $OutputFile10 -connectionString $connectionString

			LogProgress "Getting Uninstall Products Data"
			Invoke-SQL -SQLCommand $sqlCommand11 -ResultsFilePath $OutputFile11 -connectionString $connectionString

			LogProgress "Getting Services Data"
			Invoke-SQL -SQLCommand $sqlCommand12 -ResultsFilePath $OutputFile12 -connectionString $connectionString

			LogProgress "Getting Software Licensing Products Data"
			Invoke-SQL -SQLCommand $sqlCommand13 -ResultsFilePath $OutputFile13 -connectionString $connectionString

			LogProgress "Getting SQL Server Inventory Data"
			Invoke-SQL -SQLCommand $sqlCommand15 -ResultsFilePath $OutputFile15 -connectionString $connectionString
		}
	}
	catch{
		$currentException = $Error[0].Exception;
		if ($currentException.Message -like "*Make sure that the name is entered correctly*") {
			LogProgress "Can not find database $DatabaseName"
		}
		else {
			LogLastException
		}
	}
}

$sqlCommandGetMAPDBList = "Select Name, database_id, master.dbo.fn_varbintohexstr(owner_sid), create_date from Sys.Databases"

$sqlCommand1 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,CONVERT(varchar(19), [lastLogonTimestamp], 126) AS [lastLogonTimestamp]
	,CONVERT(varchar(19), [pwdLastSet], 126) AS [pwdLastSet]
	,CONVERT(varchar(19), [LastLogon], 126) AS [LastLogon]
FROM
	[Core_Inventory].[AdDiscoveredDevices];
"@

$sqlCommand2 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber]
	,[AdOsVersion] AS [AdOsVersion]
	,[ActivationRequired] AS [ActivationRequired]
	,[AdFullyQualifiedDomainName] AS [AdFullyQualifiedDomainName]
	,CONVERT(varchar(19), [BiosReleaseDate], 126) AS [BiosReleaseDate]
	,[BiosSerialNumber] AS [BiosSerialNumber]
	,[ComputerSystemName] AS [ComputerSystemName]
	,[CreateCollectorId] AS [CreateCollectorId]
	,CONVERT(varchar(19), [CreateDatetime], 126) AS [CreateDatetime]
	,[DnsHostName] AS [DnsHostName]
	,[EnclosureManufacturer] AS [EnclosureManufacturer] 
	,[EnclosureSerialNumber] AS [EnclosureSerialNumber]
	,[HostNameForVm] AS [HostNameForVm]
	,[Model] AS [Model]
	,[NetServerEnumOsVersion] AS [NetServerEnumOsVersion]
	,[OperatingSystem] AS [OperatingSystem]
	,[OperatingSystemSku] AS [OperatingSystemSku]
	,[OsCaption] AS [OsCaption]
	,CONVERT(varchar(19), [OsInstallDate], 126) AS [OsInstallDate]
	,[OsProductSuite] AS [OsProductSuite]
	,[WmiOsVersion] AS [WmiOsVersion]
	,[WmiScanResult] AS [WmiScanResult]
FROM
	[Core_Inventory].[Devices];
"@

$sqlCommand3 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,[Uid] AS [Uid]
	,CONVERT(varchar(19), [CreateDatetime], 126) AS [CreateDatetime]
	,[Clustered] AS [Clustered]
	,[SkuName] AS [SkuName]
	,[Version] AS [Version]
	,[VsName] AS [VsName]
	,[InstanceName] AS [InstanceName]
	,[Servicename] AS [Servicename]
FROM
	[SqlServer_Inventory].[Inventory];
"@
	
$sqlCommand4 = @"
use "$DatabaseName";
SELECT
	[ObjectGuid] AS [ObjectGuid]
	,[Name] AS [Name]
	,CONVERT(varchar(19), [CreateDatetime], 126) AS [CreateDatetime]
	,[Enterprise] AS [Enterprise]
	,[MergedVersionNumber] AS [MergedVersionNumber]
FROM
	[UT_Exchange_Inventory].[AdServers];
"@
	
$sqlCommand5 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,[MobRef] AS [MobRef]
	,[GuestHostName] AS [GuestHostName]
	,[RunTimeHost] AS [RunTimeHost]
	,[SourceApiType] AS [SourceApiType]
	,[VmwareName] AS [VmwareName]
	,CONVERT(varchar(19), [CreateDatetime], 126) AS [CreateDatetime]
FROM
	[VMware_Inventory].[Guest];
"@

$sqlCommand6 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,[MobRef] AS [MobRef]
	,CONVERT(varchar(19), [CreateDatetime], 126) AS [CreateDatetime]
	,[DnsConfigDomainName] AS [DnsConfigDomainName]
	,[DnsConfigHostName] AS [DnsConfigHostName]
	,[HardwareCpuModel] AS [HardwareCpuModel]
	,[HardwareModel] AS [HardwareModel]
	,[HardwareNumCpuCores] AS [HardwareNumCpuCores]
	,[HardwareNumCpuPkgs] AS [HardwareNumCpuPkgs]
	,[HardwareNumCpuThreads] AS [HardwareNumCpuThreads]
	,[HardwareVendor] AS [HardwareVendor]
	,[ProductFullName] AS [ProductFullName]
	,[ProductProductLineId] AS [ProductProductLineId]
	,[ProductVersion] AS [ProductVersion]
	,[SourceApiType] AS [SourceApiType]
	,[VmwareName] AS [VmwareName]
FROM
	[VMware_Inventory].[Host];
"@
	
$sqlCommand7 = @"
use "$DatabaseName";
SELECT
	 [DeviceNumber] AS [DeviceNumber]
	,[Uid] AS [Uid]
	,[MacAddress] AS [MacAddress]
FROM
	[Win_Inventory].[NetworkAdapters];
"@

$sqlCommand8 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,[SlicTable] AS [SlicTable]
	,[HasMsdmTable] AS [HasMsdmTable]
FROM
	[Win_Inventory].[OemData];
"@

$sqlCommand9 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber]AS [DeviceNumber]
	,[Uid] AS [Uid]
	,[ProcessorId] AS [ProcessorId]
	,[SocketDesignation] AS [SocketDesignation]
	,[NumberOfCores] AS [NumberOfCores]
	,[NumberOfLogicalProcessors] AS [NumberOfLogicalProcessors]
	,[Name] AS [Name]
FROM
	[Win_Inventory].[Processors];
"@

$sqlCommand10 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber]AS [DeviceNumber]
	,[Uid] AS [Uid]
	,[Caption] AS [Caption]
	,CONVERT(varchar(19), [CreateDatetime], 126) AS [CreateDatetime]
	,[IdentifyingNumber] AS [IdentifyingNumber]
	,CONVERT(varchar(19), [InstallDate], 126) AS [InstallDate]
	,[InstallLocation] AS [InstallLocation]
	,[Vendor] AS [Vendor]
	,[Version] AS [Version]
FROM
	[Win_Inventory].[Products]
WHERE
	[Vendor] LIKE '%Microsoft%'
OR
	[Vendor] LIKE '%VMWare%';
"@

$sqlCommand11 = @"
use "$DatabaseName";
SELECT
	 [DeviceNumber] AS [DeviceNumber]
	,[ProductCode] AS [ProductCode]
	,[CreateCollectorId] AS [CreateCollectorId]
	,CONVERT(varchar(19), [CreateDatetime], 126) AS [CreateDatetime]
	,[DisplayName] AS [DisplayName]
	,[DisplayVersion] AS [DisplayVersion]
	,[InstallDate] AS [InstallDate]
	,[InstallLocation] AS [InstallLocation]
	,[Publisher] AS [Publisher]
FROM
	[Win_Inventory].[ProductsUninstall]
WHERE
	[Publisher] LIKE '%Microsoft%'
AND
	CHARINDEX('KB', [DisplayName]) +
	CHARINDEX('.NET Framework', [DisplayName]) +
	CHARINDEX('Update', [DisplayName]) +
	CHARINDEX('Service Pack', [DisplayName]) +
	CHARINDEX('Proof', [DisplayName]) +
	CHARINDEX('Components', [DisplayName]) +
	CHARINDEX('Tools', [DisplayName]) +
	CHARINDEX('MUI', [DisplayName]) +
	CHARINDEX('Redistributable', [DisplayName]) = 0
"@

$sqlCommand12 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber]AS [DeviceNumber]
	,[Uid] AS [Uid]
	,[Name] AS [Name]
	,[State] AS [State]
FROM
	[Win_Inventory].[Services];
"@

$sqlCommand13 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,[ID] AS [ID]
	,[Name] AS [Name]
	,[Description] AS [Description]
FROM
	[Win_Inventory].[SoftwareLicensingProducts];
"@
	
$sqlCommand14 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,[Name] AS [Name]
FROM
	[WinServer_Inventory].[MSClusterCluster];
"@

$sqlCommand15 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber]
	,[Clustered]
	,[Skuname]
	,[VersionCoalesce]
	,[Vsname]
	,[InstanceName]
	,[Servicename]
FROM
	[SqlServer_Assessment].[SqlInstances];
"@
	

Get-MAPInventoryData