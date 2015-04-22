 ##########################################################################
 # 
 # Get_NetworkNeighborhoodList
 # SAM Gold Toolkit
 # Original Source: Sam360
 #
 ##########################################################################
 
  Param(
	[alias("o1")]
	$OutputFile1 = "NetworkNeighborhoodList.csv"
    )

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
	Write-Output "Computer Name:		$($env:COMPUTERNAME)" #-ForegroundColor Magenta
	Write-Output "User Name:			$($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:		$($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:		$($host.Version.Major)"
	Write-Output "PowerShell Version:	$($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size:	$($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:			$($PSVersionTable.CLRVersion)"
	Write-Output "Username Parameter:	$UserName"
	Write-Output "Server Parameter:		$Server"
}

 function Get_NetworkNeighborhoodList {
	#ShellSpecialFolderConstants 
	#http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx
	
	try {
		LogEnvironmentDetails

		$shellFolder = ( new-object -com shell.application ).NameSpace(0x12)
		$shellFolder.Items() | select Name, Path | export-csv $OutputFile1 -notypeinformation -Encoding UTF8 }
	catch {
		LogLastException
	}
 }

 Get_NetworkNeighborhoodList