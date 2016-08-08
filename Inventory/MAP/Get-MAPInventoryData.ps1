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
	$OutputFile1 = "$PWD\Devices.csv",  
	[alias("o2")]
	$OutputFile2 = "$PWD\Products1.csv",
	[alias("o3")]
	$OutputFile3 = "$PWD\Products2.csv",
	[alias("o4")]
	$OutputFile4 = "$PWD\ADDiscoveredDevices.csv",
	[alias("o5")]
	$OutputFile5 = "$PWD\SQLServerInventory1.csv",
	[alias("o6")]
	$OutputFile6 = "$PWD\ExchangeInventory.csv",
	[alias("o7")]
	$OutputFile7 = "$PWD\VMwareGuests.csv",
	[alias("o8")]
	$OutputFile8 = "$PWD\VMwareHosts.csv",
	[alias("o9")]
	$OutputFile9 = "$PWD\NetworkAdapters.csv",
	[alias("o10")]
	$OutputFile10 = "$PWD\OEMData.csv",
	[alias("o11")]
	$OutputFile11 = "$PWD\Processors.csv",
	[alias("o12")]
	$OutputFile12 = "$PWD\Services.csv",
	[alias("o13")]
	$OutputFile13 = "$PWD\SoftwareLicensingProducts.csv",
	[alias("o14")]
	$OutputFile14 = "$PWD\MSClusters.csv",
	[alias("o15")]
	$OutputFile15 = "$PWD\SQLServerInventory2.csv",
	[alias("o16")]
	$OutputFile16 = "MAPDBNames.csv",
	[alias("log")]
	[string] $LogFile = "MAPQueryLog.txt",
	[ValidateSet("DBList","AllData","DeviceData","SoftwareData","BasicData")]
	$RequiredData = "BasicData")

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
	LogText -Color Gray " Get-MAPInventoryData.ps1"
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
	LogText -Color Gray "Log File:                        $LogFile"
	LogText -Color Gray "Database:                        $DatabaseName"
	LogText -Color Gray "Required Data:                   $RequiredData"
	LogText ""
}

function GetConnectionString {
	# Start the local db instance
	$sqlLocalDBOutput = [string] (& sqllocaldb start MAPToolkit 2>&1)

	# Get info for the db instance
	$sqlLocalDBOutput = [string] (& sqllocaldb info MAPToolkit 2>&1)
	$sqlLocalDBInstanceInfo = $sqlLocalDBOutput.Split("`n ")

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
	$fileWriter = New-Object System.IO.StreamWriter($resultsFilePath, $false, [System.Text.Encoding]::UTF8)

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
	
	$dbNames = $sqlLocalDBOutput.Split("`n ")

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

		if ($RequiredData -eq "BasicData") {
			LogProgress "Getting Device Data"
			Invoke-SQL -SQLCommand $sqlCommandHardware -ResultsFilePath $OutputFile1 -connectionString $connectionString

			LogProgress "Getting Software Data"
			Invoke-SQL -SQLCommand $sqlCommandSoftware -ResultsFilePath $OutputFile2 -connectionString $connectionString

			return
		}
			
		if ($RequiredData -eq "DeviceData" -or $RequiredData -eq "AllData") {
			LogProgress "Getting Devices Data"
			Invoke-SQL -SQLCommand $sqlCommand1 -ResultsFilePath $OutputFile1 -connectionString $connectionString
			
			LogProgress "Getting AdDiscoveredDevices Data"
			Invoke-SQL -SQLCommand $sqlCommand4 -ResultsFilePath $OutputFile4 -connectionString $connectionString

			LogProgress "Getting VMware Guests Data"
			Invoke-SQL -SQLCommand $sqlCommand7 -ResultsFilePath $OutputFile7 -connectionString $connectionString

			LogProgress "Getting VMware Hosts Data"
			Invoke-SQL -SQLCommand $sqlCommand8 -ResultsFilePath $OutputFile8 -connectionString $connectionString

			LogProgress "Getting Network Adapter Data"
			Invoke-SQL -SQLCommand $sqlCommand9 -ResultsFilePath $OutputFile9 -connectionString $connectionString

			LogProgress "Getting Processor Data"
			Invoke-SQL -SQLCommand $sqlCommand11 -ResultsFilePath $OutputFile11 -connectionString $connectionString

			LogProgress "Getting Cluster Data"
			Invoke-SQL -SQLCommand $sqlCommand14 -ResultsFilePath $OutputFile14 -connectionString $connectionString
		}
		
		if ($RequiredData -eq "SoftwareData" -or $RequiredData -eq "AllData") {
			LogProgress "Getting SQL Server Inventory Data"
			Invoke-SQL -SQLCommand $sqlCommand5 -ResultsFilePath $OutputFile5 -connectionString $connectionString
			
			LogProgress "Getting Exchange Inventory Data"
			Invoke-SQL -SQLCommand $sqlCommand6 -ResultsFilePath $OutputFile6 -connectionString $connectionString

			LogProgress "Getting OEM Data"
			Invoke-SQL -SQLCommand $sqlCommand10 -ResultsFilePath $OutputFile10 -connectionString $connectionString

			LogProgress "Getting Products Data"
			Invoke-SQL -SQLCommand $sqlCommand2 -ResultsFilePath $OutputFile2 -connectionString $connectionString

			LogProgress "Getting Uninstall Products Data"
			Invoke-SQL -SQLCommand $sqlCommand3 -ResultsFilePath $OutputFile3 -connectionString $connectionString

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

$sqlCommand2 = @"
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

$sqlCommand3 = @"
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

$sqlCommand4 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,CONVERT(varchar(19), [lastLogonTimestamp], 126) AS [lastLogonTimestamp]
	,CONVERT(varchar(19), [pwdLastSet], 126) AS [pwdLastSet]
	,CONVERT(varchar(19), [LastLogon], 126) AS [LastLogon]
FROM
	[Core_Inventory].[AdDiscoveredDevices];
"@

$sqlCommand5 = @"
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
	
$sqlCommand6 = @"
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
	
$sqlCommand7 = @"
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

$sqlCommand8 = @"
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
	
$sqlCommand9 = @"
use "$DatabaseName";
SELECT
	 [DeviceNumber] AS [DeviceNumber]
	,[Uid] AS [Uid]
	,[MacAddress] AS [MacAddress]
FROM
	[Win_Inventory].[NetworkAdapters];
"@

$sqlCommand10 = @"
use "$DatabaseName";
SELECT
	[DeviceNumber] AS [DeviceNumber]
	,[SlicTable] AS [SlicTable]
	,[HasMsdmTable] AS [HasMsdmTable]
FROM
	[Win_Inventory].[OemData];
"@

$sqlCommand11 = @"
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

$sqlCommandHardware = @"
use "$DatabaseName";
SELECT
	DeviceNumber As 'SourceKey',
	1 As 'SystemClientVersion',
	ComputerSystemName,
	ComputerSystemName As 'SystemName',
	DNSHostName As 'SystemNetBiosName',
	AdFullyQualifiedDomainName As 'Domain',
	AdDomainName As 'Resource_Domain_OR_Workgr0',
	DistinguishedName As 'Distinguished_Name0',
	BiosSerialNumber As 'BiosSerialNumber',
	CONVERT(varchar(19), BiosReleaseDate, 126) As 'BiosReleaseDate',
	(SELECT DISTINCT
			[MACAddress] + ';'
		FROM 
			Win_Inventory.NetworkAdapters
		WHERE 
			Win_Inventory.NetworkAdapters.DeviceNumber = Core_Inventory.Devices.DeviceNumber
		AND
			[MACAddress] IS NOT NULL 
		AND
			[MACAddress] NOT IN ('00:00:00:00:00:00','33:50:6F:45:30:30','50:50:54:50:30:30')
		ORDER BY 
			[MACAddress] + ';'
		FOR XML PATH ('')
	) AS 'MacAddress',
	'' As 'IPAddress',
	CONVERT(varchar(19), LastLogonTimestamp, 126) As 'LastLogon',
	CONVERT(varchar(19), LocalDatetime, 126) As 'LastHWScan',
	CONVERT(varchar(19), LocalDatetime, 126) As 'LastSWScan',
	OperatingSystem As 'OperatingSystem',
	'' As 'OperatingSystemSerialNumber',
	OperatingSystemServicePack As 'OperatingSystemServicePack',
	CONVERT(varchar(19), OsInstallDate, 126) As 'OperatingSystemInstallDate',
	OsVersion As 'OperatingSystemVersion',
	TotalPhysicalMemory/1000000 As 'PhysicalMemory', --6
	TotalVirtualMemorySize/1000 As 'VirtualMemory', --3
	NumberOfProcessors As 'ProcessorCount',
	NumberOfCores As 'CoreCount',
	NumberOfLogicalProcessors As 'LogicalProcessorCount',
	(SELECT Top 1 Name 
		FROM Win_Inventory.Processors
		WHERE Win_Inventory.Processors.DeviceNumber = Core_Inventory.Devices.DeviceNumber
	) AS 'CpuType',
	Model As 'Model',
	EnclosureManufacturer As 'Manufacturer',
	VmFriendlyName As 'VirtualHostName',
	HostNameForVm As 'VirtualPhysicalHostName',
	'' As 'VirtualResourceID'
 FROM
	Core_Inventory.Devices
"@

$sqlCommandSoftware = @"
use "$DatabaseName";
SELECT 
	DeviceNumber As 'DeviceSourceKey',
	CreateDatetime As 'DiscoveryDate',
	InstallLocation As 'InstalledLocation',
	DisplayName As 'SwName',
	DisplayVersion As 'SwVersion',
	Publisher As 'SwPublisher',
	IdentifyingNumber As 'SwSoftwareCode',
	Uid As 'SwSoftwareId',
	'' As IsOperatingSystem,
	Source
FROM
	(SELECT
		DeviceNumber,
		CONVERT(varchar(19), CreateDatetime, 126) AS 'CreateDatetime',
		InstallLocation,
		Caption As 'DisplayName',
		Version As 'DisplayVersion',
		Vendor As 'Publisher',
		IdentifyingNumber,
		'{' + convert(nvarchar(50), Uid) + '}' As 'Uid',
		'MAP1' As 'Source'
	FROM
		Win_Inventory.Products
	WHERE
		Vendor LIKE '%Microsoft%' OR Vendor LIKE '%VMWare%'
	UNION ALL
	SELECT
		 DeviceNumber,
		 CreateDatetime,
		 InstallLocation,
		 DisplayName,
		 DisplayVersion,
		 Publisher,
		 '' As 'IdentifyingNumber',
		 '' As 'Uid',
		 'MAP2' As 'Source'
	FROM
		Win_Inventory.ProductsUninstall
	WHERE
		Publisher LIKE '%Microsoft%'
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
	) As ProductsCombined
GROUP BY
	DeviceNumber, Publisher, DisplayName, DisplayVersion, CreateDatetime, IdentifyingNumber, Uid, InstallLocation, Source
"@

Get-MAPInventoryData