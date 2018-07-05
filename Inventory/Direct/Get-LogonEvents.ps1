 ##########################################################################
 #
 # Get-LogonEvents
 # SAM Gold Toolkit
 # Original Source: Sam360, Microsoft SAM Workspace Discovery Tool
 #
 ##########################################################################
 
 <#
.SYNOPSIS
#>

 Param(
    [alias("o1")]
    [string] $OutputFile1 = "Events.csv",
    [alias("log")]
    [string] $LogFile = "LogFile.txt",
    [DateTime] $StartDate = [DateTime]::MinValue,
    [int] $DaysToCollect = 90)

function LogText {
    param(
        [Parameter(Position=0, ValueFromRemainingArguments=$true, ValueFromPipeline=$true)]
        [Object] $Object,
        [System.ConsoleColor]$color = [System.Console]::ForegroundColor,
        [switch]$NoNewLine = $false  
    )

    # Display text on screen
    Write-Host -Object $Object -ForegroundColor $color -NoNewline:$NoNewLine

    if ($LogFile) {
        $Object | Out-File $LogFile -Encoding utf8 -Append 
    }
}

function InitialiseLogFile {
    if ($LogFile -and (Test-Path $LogFile)) {
        Remove-Item $LogFile
    }
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

function GetScriptPath
{
    if($PSCommandPath){
        return $PSCommandPath; }
        
    if($MyInvocation.ScriptName){
        return $MyInvocation.ScriptName }
        
    if($script:MyInvocation.MyCommand.Path){
        return $script:MyInvocation.MyCommand.Path }

    return $script:MyInvocation.MyCommand.Definition
}


function LogEnvironmentDetails {
    LogText -Color Gray " "
    LogText -Color Gray "   _____         __  __    _____       _     _   _______          _ _    _ _   "
    LogText -Color Gray "  / ____|  /\   |  \/  |  / ____|     | |   | | |__   __|        | | |  (_) |  "
    LogText -Color Gray " | (___   /  \  | \  / | | |  __  ___ | | __| |    | | ___   ___ | | | ___| |_ "
    LogText -Color Gray "  \___ \ / /\ \ | |\/| | | | |_ |/ _ \| |/ _`` |    | |/ _ \ / _ \| | |/ / | __|"
    LogText -Color Gray "  ____) / ____ \| |  | | | |__| | (_) | | (_| |    | | (_) | (_) | |   <| | |_ "
    LogText -Color Gray " |_____/_/    \_\_|  |_|  \_____|\___/|_|\__,_|    |_|\___/ \___/|_|_|\_\_|\__|"
    LogText -Color Gray " "
    LogText -Color Gray " Get-LogonEvents.ps1"
    LogText -Color Gray " "

    $OSDetails = Get-WmiObject Win32_OperatingSystem
    $ScriptPath = GetScriptPath
    $Elevated = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
    LogText -Color Gray "Computer Name:        $($env:COMPUTERNAME)"
    LogText -Color Gray "User Name:            $($env:USERNAME)@$($env:USERDNSDOMAIN)"
    LogText -Color Gray "Windows Version:      $($OSDetails.Caption)($($OSDetails.Version))"
    LogText -Color Gray "PowerShell Host:      $($host.Version.Major)"
    LogText -Color Gray "PowerShell Version:   $($PSVersionTable.PSVersion)"
    LogText -Color Gray "PowerShell Word size: $($([IntPtr]::size) * 8) bit"
    LogText -Color Gray "CLR Version:          $($PSVersionTable.CLRVersion)"
    LogText -Color Gray "Elevated:             $Elevated"
    LogText -Color Gray "Current Date Time:    $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
    LogText -Color Gray "Script Path:          $ScriptPath"
    LogText -Color Gray "Output File 1:        $OutputFile1"
    LogText -Color Gray "StartDate:            $StartDate"
    LogText -Color Gray "DaysToCollect:        $DaysToCollect"
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

function GetLogonEvents() {
    try {
        InitialiseLogFile
        LogEnvironmentDetails
        SetupDateFormats

        if ($StartDate -eq [DateTime]::MinValue) {
            $StartDate = (Get-Date).addDays(-$DaysToCollect)
        }

        $events = @()
        $logEvents = Get-EventLog -LogName Security -After $StartDate | 
            where {($_.EventID -eq 4624) -or ($_.EventID -eq 528)}
        
        foreach ($logEvent in $logEvents) {
            
            $eventParameters = $logEvent.ReplacementStrings
            if ($eventParameters -eq $null){
                LogText "Event Missing Parameters. ID: $($logEvent.EventID) Time: $($logEvent.TimeGenerated)"
                continue
            }

            $logonType = "Unknown"
            if ($logEvent.EventID -eq 528){ # Win2K, WinXP, Win2003
                if ($eventParameters[3]){
                    $logonType = $eventParameters[3]
                }
            }
            elseif ($logEvent.EventID -eq 4624){ # WinVista+
                if ($eventParameters[8]){
                    $logonType = $eventParameters[8]
                }
            }
            
            if (($logonType -eq "3") -or # Network (e.g. connection to shared folder on this computer from elsewhere on network)
                ($logonType -eq "4") -or # Batch (e.g. scheduled task)
                ($logonType -eq "5") -or # Service (i.e. Service startup)
                ($logonType -eq "8")) { # NetworkCleartext (Logon with credentials sent in clear text e.g. IIS "Basic Authentication")
                continue
            }

            $event = new-object PSObject
            $event | Add-Member -MemberType NoteProperty -Name "ID" -Value $logEvent.EventID
            $event | Add-Member -MemberType NoteProperty -Name "TimeGenerated" -Value $logEvent.TimeGenerated
            $event | Add-Member -MemberType NoteProperty -Name "LogonType" -Value $logonType
            $event | Add-Member -MemberType NoteProperty -Name "UserName" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "UserDomain" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "SID" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "SourceNetworkAddress" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "SourcePort" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "WorkstationName" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "LogonProcess" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "AuthenticationPackage" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "PackageName" -Value ""
            $event | Add-Member -MemberType NoteProperty -Name "KeyLength" -Value ""

            if ($logEvent.EventID -eq 528){ # Win2K, WinXP, Win2003
                $event.UserName = $eventParameters[0]
                $event.UserDomain = $eventParameters[1]
                $event.WorkstationName = $eventParameters[6]
                $event.LogonProcess = $eventParameters[4]
                $event.AuthenticationPackage = $eventParameters[5]
            }
            elseif ($logEvent.EventID -eq 4624){ # WinVista+
                $event.UserName = $eventParameters[5]
                $event.UserDomain = $eventParameters[6]
                $event.SID = $eventParameters[4]
                $event.SourceNetworkAddress = $eventParameters[18]
                $event.SourcePort = $eventParameters[19]
                $event.WorkstationName = $eventParameters[11]
                $event.LogonProcess = $eventParameters[9]
                $event.AuthenticationPackage = $eventParameters[10]
                $event.PackageName = $eventParameters[14]
                $event.KeyLength = $eventParameters[15]
            }

            $events += $event
        }

        if ($events) {
            $events | export-csv $OutputFile1 -notypeinformation -Encoding UTF8
        }
    }
    catch {
        LogLastException
    }
}

GetLogonEvents