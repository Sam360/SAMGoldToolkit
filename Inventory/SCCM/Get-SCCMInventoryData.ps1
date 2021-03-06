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
    [alias("o3")]
    $OutputFile3 = "Services.csv",
    [alias("o4")]
    $OutputFile4 = "IEFiles.csv",
    [alias("log")]
    [string] $LogFile = "SCCMLogFile.txt",
    $UserName,
    $Password,
    [ValidateSet("AllData","DeviceData","SoftwareData")]
    $RequiredData = "AllData",
    $PortNumber = "0",
    $SQLCommandTimeout = 300,
    [ValidateSet("2007","2012")]
    $SCCMVersion = "2012",
    [switch]
    $AllVendors)
    
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

function LogEnvironmentDetails {
    LogText -Color Gray "   _____         __  __    _____       _     _   _______          _ _    _ _   "
    LogText -Color Gray "  / ____|  /\   |  \/  |  / ____|     | |   | | |__   __|        | | |  (_) |  "
    LogText -Color Gray " | (___   /  \  | \  / | | |  __  ___ | | __| |    | | ___   ___ | | | ___| |_ "
    LogText -Color Gray "  \___ \ / /\ \ | |\/| | | | |_ |/ _ \| |/ _`` |    | |/ _ \ / _ \| | |/ / | __|"
    LogText -Color Gray "  ____) / ____ \| |  | | | |__| | (_) | | (_| |    | | (_) | (_) | |   <| | |_ "
    LogText -Color Gray " |_____/_/    \_\_|  |_|  \_____|\___/|_|\__,_|    |_|\___/ \___/|_|_|\_\_|\__|"
    LogText -Color Gray " "
    LogText -Color Gray " Get-SCCMInventoryData.ps1"
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
    LogText -Color Gray "Username Parameter:   $UserName"
    LogText -Color Gray "Server Parameter:     $DatabaseServer"
    LogText -Color Gray "Database Parameter:   $DatabaseName"
    LogText -Color Gray "Required Data:        $RequiredData"
    LogText -Color Gray "Output File 1:        $OutputFile1"
    LogText -Color Gray "Output File 2:        $OutputFile2"
    LogText -Color Gray "Output File 3:        $OutputFile3"
    LogText -Color Gray ""
}

function SetupDateFormats {
    # Standardise date/time output to ISO 8601'ish format
    $bDateFormatConfigured = $false
    $currentThread = [System.Threading.Thread]::CurrentThread
    
    try {
        $CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = 'yyyy-MM-dd'
        $CurrentThread.CurrentCulture.DateTimeFormat.LongDatePattern = 'yyyy-MM-dd HH:mm:ss'
        $bDateFormatConfigured = $true
    }
    catch {
    }

    if (!($bDateFormatConfigured)) {
        try {
            $cultureCopy = $CurrentThread.CurrentCulture.Clone()
            $cultureCopy.DateTimeFormat.ShortDatePattern = 'yyyy-MM-dd'
            $cultureCopy.DateTimeFormat.LongDatePattern = 'yyyy-MM-dd HH:mm:ss'
            $currentThread.CurrentCulture = $cultureCopy
        }
        catch {
        }
    }
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
    $connection = new-object system.data.SqlClient.SQLConnection(GetConnectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
    $command.CommandTimeout = $SQLCommandTimeout
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
        InitialiseLogFile
        LogEnvironmentDetails
        SetupDateFormats

        if ($RequiredData -eq "DeviceData" -or $RequiredData -eq "AllData") {
            
        LogText "Querying Device Data"
            if ($SCCMVersion -eq "2007") {
                Invoke-SQL -SQLCommand $sqlCommandDevices2007 -ResultsFilePath $OutputFile1 
            }
            else {
                Invoke-SQL -SQLCommand $sqlCommandDevices2012 -ResultsFilePath $OutputFile1
            }
        }
        
        if ($RequiredData -eq "SoftwareData" -or $RequiredData -eq "AllData") {
            if ($AllVendors) {
                $sqlCommandSoftware = $sqlCommandSoftware.replace('--//All Vendors//--','')
            }

            LogText "Querying Software Data"
            Invoke-SQL -SQLCommand $sqlCommandSoftware -ResultsFilePath $OutputFile2

            LogText "Querying Services Data"
            Invoke-SQL -SQLCommand $sqlCommandServices -ResultsFilePath $OutputFile3

            LogText "Querying IE Data"
            Invoke-SQL -SQLCommand $sqlIEFileData -ResultsFilePath $OutputFile4
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
    ,[System].[AD_Site_Name0] AS [AD_Site_Name0]
    ,[System].[Is_Virtual_Machine0] AS [Is_Virtual_Machine0]
    ,[System].[User_Domain0] AS [User_Domain0]
    ,[System].[User_Name0] AS [User_Name0]
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
    ,[ComputerSystem].[UserName0] AS [UserName]
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
            ,[UserName0]
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

# This script used to use v_GS_INSTALLED_SOFTWARE
# v_GS_INSTALLED_SOFTWARE requires Asset Intelligence and software scanning to be enabled
# This table included install location

# Now using v_ADD_REMOVE_PROGRAMS which is populated during HW scan
# It is the combination of ARP 32 bit and ARP 64 bit
# i.e. Union of v_GS_ADD_REMOVE_PROGRAMS and v_GS_ADD_REMOVE_PROGRAMS_64 

# Other notes: 
# v_HS_ADD_REMOVE_PROGRAMS & v_HS_ADD_REMOVE_PROGRAMS_64 contain historical data
# v_GS_SoftwareProduct is populated by software inventory and is based on information obtained from the file headers


$sqlCommandSoftware = @"
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
            ,NULL AS [InstalledLocation]
            ,COALESCE(NULLIF([Software].[DisplayName0], ''), N'Unknown') AS [ProductName]
            ,COALESCE(NULLIF([Software].[Version0], ''), N'Unknown') AS [ProductVersion]
            ,COALESCE(NULLIF([Software].[Publisher0], ''), N'Unknown') AS [Publisher]
            ,[Software].[ProdID0] AS [SoftwareCode]
            ,[Software].[ProdID0] AS [PackageCode]
            ,[Software].[InstallDate0] AS [InstallDate]
            ,0 AS [IsOperatingSystem]
            ,'v_ARP' AS [Source]
        FROM 
            [dbo].[v_ADD_REMOVE_PROGRAMS] AS [Software]
-- Apply this filter to remove software we don't care about at the present time to reduce the volume
        WHERE
            (CHARINDEX('microsoft', [Software].[Publisher0]) > 0
            AND
                CHARINDEX('KB', [Software].[DisplayName0]) +
                CHARINDEX('.NET Framework', [Software].[DisplayName0]) +
                CHARINDEX('Update', [Software].[DisplayName0]) +
                CHARINDEX('Service Pack', [Software].[DisplayName0]) +
                CHARINDEX('Proof', [Software].[DisplayName0]) +
                CHARINDEX('Components', [Software].[DisplayName0]) +
                CHARINDEX('Tools', [Software].[DisplayName0]) +
                CHARINDEX('MUI', [Software].[DisplayName0]) +
                CHARINDEX('Redistributable', [Software].[DisplayName0]) = 0
            )
            --//All Vendors//-- OR CHARINDEX('microsoft', [Software].[Publisher0]) <= 0
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
            ,[OperatingSystem].[InstallDate0] AS [InstallDate]
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
    ,[ComputerSystem].[UserName0] AS [UserName]
    --/--,NULLIF([System].[Virtual_Machine_Host_Name0], '') AS [VirtualHostName]
    --/--,[VirtualMachine].[PhysicalHostName0] AS [VirtualPhysicalHostName]
    --/--,[VirtualMachine].[ResourceID] AS [VirtualResourceID]
    --,'vRSystem' AS [Source]
    --,[FullCollectionMembership].[CollectionID] --
FROM 
    [dbo].[v_R_System] AS [System] 
--INNER JOIN 
--  [dbo].[v_FullCollectionMembership] AS [FullCollectionMembership]
--ON
--  [System].[ResourceID] = [FullCollectionMembership].[ResourceID]
--AND
--  [FullCollectionMembership].[CollectionID] IN ('SMS00001')
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
            ,[UserName0]
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

$sqlCommandServices = @"
SELECT DISTINCT
    [Services].[ResourceID] AS [SourceKey],
    [Services].[AcceptPause0] AS [AcceptPause],
    [Services].[AcceptStop0] AS [AcceptStop],
    [Services].[Caption0] AS [Caption],
    [Services].[Description0] AS [Description],
    [Services].[DesktopInteract0] AS [DesktopInteract],
    [Services].[DisplayName0] AS [DisplayName],
    [Services].[ErrorControl0] AS [ErrorControl],
    [Services].[ExitCode0] AS [ExitCode],
    [Services].[Name0] AS [Name],
    [Services].[PathName0] AS [PathName],
    [Services].[ProcessId0] AS [ProcessId],
    [Services].[ServiceSpecificExitCode0] AS [ServiceSpecificExitCode],
    [Services].[ServiceType0] AS [ServiceType],
    [Services].[Started0] AS [Started],
    [Services].[StartMode0] AS [StartMode],
    [Services].[StartName0] AS [StartName],
    [Services].[State0] AS [State],
    [Services].[Status0] AS [Status],
    [Services].[TagId0] AS [TagId],
    [Services].[WaitHint0] AS [WaitHint]
FROM
    [dbo].[v_GS_SERVICE] AS [Services]
"@

$sqlIEFileData = @"
SELECT
    [SoftwareFiles].[ResourceID],
    [SoftwareFiles].[FileName],
    [SoftwareFiles].[FileDescription],
    [SoftwareFiles].[FileVersion],
    [SoftwareFiles].[FileSize],
    [SoftwareFiles].[FilePath]
From
    [dbo].[v_GS_SoftwareFile] As [SoftwareFiles]
Where
    [SoftwareFiles].FileName = 'iexplore.exe'
    and [SoftwareFiles].FilePath like '%Internet Explorer%'
"@

Get-SCCMInventoryData