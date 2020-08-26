 ##########################################################################
 #
 # Get_XenAppDetails
 # SAM Gold Toolkit
 # Original Source: Sam360
 #
 ##########################################################################

 Param(
	[alias("o1")]
	$OutputFile1 = "ApplicationReport.csv")
	
function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Output "Computer Name:        $($env:COMPUTERNAME)" #-ForegroundColor Magenta
	Write-Output "User Name:            $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:      $($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:      $($host.Version.Major)"
	Write-Output "PowerShell Version:   $($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size: $($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:          $($PSVersionTable.CLRVersion)"
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

function Get_XenAppDetails {
	LogEnvironmentDetails
    SetupDateFormats

	try {
		add-pssnapin citrix*
		Get-XAApplicationReport * | Export-Csv $OutputFile1 -notypeinformation -Encoding UTF8
	}
	catch {
		LogLastException
	}
}

Get_XenAppDetails
