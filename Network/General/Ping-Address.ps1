 ##########################################################################
 #
 # Ping-Address
 # SAM Gold Toolkit
 # Original Source: Jon Mulligan (Sam360)
 #
 ##########################################################################

 Param(
	$Target = $env:COMPUTERNAME,
	[alias("o1")]
    $OutputFile1 = "PingResults-" + $Target + ".csv")

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
	Write-Output "Target:                   $Target"
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

function Ping-Address {
	
	LogEnvironmentDetails
    SetupDateFormats

	try {

		$PingResults = Get-WmiObject Win32_PingStatus -Filter "Address='$Target' AND ResolveAddressNames=TRUE" #RecordRoute=1

		$PingResultsTidied = $PingResults | select -Property PSComputerName, 
			IPV4Address, IPV6Address, Address, ProtocolAddress, ProtocolAddressResolved, 
			ReplySize, ResolveAddressNames, ResponseTime, PrimaryAddressResolutionStatus, StatusCode
		
		$PingResultsTidied | export-csv $OutputFile1 -notypeinformation -Encoding UTF8
	}
	catch{
		LogLastException
	}
}

Ping-Address