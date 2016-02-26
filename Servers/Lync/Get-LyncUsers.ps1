 ##########################################################################
 # 
 # Get-LyncUsers
 # SAM Gold Toolkit
 # Original Source: Akshay Chiddarwar (Sam360)
 #
 ##########################################################################

<#
.SYNOPSIS
 Retrieves list of lync users from remote computer
 
    Files are written to current working directory


.PARAMETER      UserName
Lync Server Username

.PARAMETER		Password
Lync Server Password

.PARAMETER      LyncServerName
Accepts a Fully Qualified Lync Server Domain Name
e.g. Server01.<Domain Name>

.PARAMETER		OutputFile1
Output CSV file to store the results

#>

param(
	$UserName,
	$Password,
	[alias("server")]
	$LyncServerName = ($env:USERDNSDOMAIN + "\" + $env:COMPUTERNAME),
    [ValidateSet("RemoteSession","SnapIn")] 
	$ConnectionMethod = "RemoteSession",
	[alias("o1")]
	[string] $OutputFile1 = "LyncUsers.csv",
	[alias("log")]
	[string] $LogFile = "LyncLogFile" + $LyncServerName + ".txt"
)

function LogEnvironmentDetails {
	$OSDetails = Get-WmiObject Win32_OperatingSystem
	Write-Output "Computer Name:            $($env:COMPUTERNAME)"
	Write-Output "User Name:                $($env:USERNAME)@$($env:USERDNSDOMAIN)"
	Write-Output "Windows Version:          $($OSDetails.Caption)($($OSDetails.Version))"
	Write-Output "PowerShell Host:          $($host.Version.Major)"
	Write-Output "PowerShell Version:       $($PSVersionTable.PSVersion)"
	Write-Output "PowerShell Word size:     $($([IntPtr]::size) * 8) bit"
	Write-Output "CLR Version:              $($PSVersionTable.CLRVersion)"	
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

function LogProgress([string]$Activity, [string]$Status, [Int32]$PercentComplete, [switch]$Completed ){
	
	Write-Progress -activity $Activity -Status $Status -percentComplete $PercentComplete -Completed:$Completed
	
	if ($Verbose)
	{
		write-output ""
		$output = Get-Date -Format HH:mm:ss.ff
		$output += " - "
		$output += $Status
		write-output $output
	}
}

# Standard CAL([bool]$bStdCAL), Enterprise CAL(bEntCAL), Plus CAL(bPlusCAL)
function CalculateCAL($bStdCAL, $bEntCAL, $bPlusCAL) {
    if ($bPlusCAL) {
        return "Plus CAL"
    }
    elseif ($bEntCAL) {
        return "Enterprise CAL"
    }
    elseif ($bStdCAL) {
        return "Standard CAL"
    }
    else {
        return "NO CAL"
    }    
}

function Get-LyncUsers {
	
	LogEnvironmentDetails

    $boolAuthenticate = $false

	Try {
        if ($ConnectionMethod -eq "SnapIn") {
            Import-Module Lync -ErrorAction SilentlyContinue -ErrorVariable $importModule
            if($importModule -eq $null) {
		        LogProgress -activity "Lync Server - Module" -Status "Loaded Lync module" -percentComplete 10 
                $boolAuthenticate = $true
            }
        }
        else {
            # URI for the LYNC Server
            $uri = ("https://" + $LyncServerName.ToLower() + "/OcsPowershell")
			
		    if ($UserName -and $Password) {
			    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
			    $cred = New-Object -TypeName System.Management.Automation.PSCredential ($UserName, $securePassword)
		    }
		    else {
			    $cred = Get-Credential “Domain\Lync_Administrator”
		    }

			LogProgress -activity "Lync Server - Authentication" -Status "Connect to the remote device" -percentComplete 5		
            $session = New-PSSession -ConnectionURI $uri -Credential $cred
                        
            if(-not($session)) {
		        LogProgress -activity "Lync Server - Authentication" -Status "Authentication - Failed" -percentComplete 100
                Write-Output "Process Failed"
			    Exit
            }
            else {
		        LogProgress -activity "Lync Server - Session" -Status "Loading authenticated session" -percentComplete 10    
                # Import session
                Import-PsSession $session -AllowClobber
                
                $boolAuthenticate = $true
            }
	    }
		
        if($boolAuthenticate) {        
            
		    LogProgress -activity "Lync Server - Get Users" -Status "Begin process" -percentComplete 12 

            $bStdCAL = $false
            $bEntCAL = $false
            $bPlusCAL = $false
            $bVoicePolicy = $false
    
		    LogProgress -activity "Lync Server - Get Users" -Status "Load all Lync users" -percentComplete 20
            $users = Get-CsUser
    
			$userCount = $users.count
			$subInterval = [int](70 / $userCount) / 6
			$percent = 20
		
            foreach ($user in $users) {
				$percent += [int]$subInterval
				LogProgress -activity "Lync Server - User Data Collection" -Status "Checking for Standard CAL requirements" -percentComplete $percent
				
                # Check for Standard CAL requirement
                if ($user.Enabled) {            
                    $bStdCAL = $true
                }
				
				# Check for Enterprise CAL requirement
                if ($user.ConferencingPolicy -ne $null) {                    
					$percent += [int]$subInterval
					LogProgress -activity "Lync Server - User Data Collection" -Status "Checking for Enterprise CAL requirements" -percentComplete $percent
				
					# Get details of the Conferencing Policy
                    if($ConnectionMethod -eq "SnapIn") {
                        $userConfPolicy = Get-CsConferencingPolicy $user.ConferencingPolicy
                    }
                    else {
                        $userConfPolicy = Get-CsConferencingPolicy $user.ConferencingPolicy.FriendlyName
                    }

                    if ( ($userConfPolicy.AllowIPAudio -eq $true) -or `
                         ($userConfPolicy.AllowIPVideo -eq $true) -or `
                         ($userConfPolicy.AllowUserToScheduleMeetingsWithAppSharing -eq $true) -or `
                         ($userConfPolicy.EnableDataCollaboration -eq $true) ) {
                        # User needs Enterprise CAL
                        $bEntCAL = $true
                    }
                }

                if ($bEntCAL) {
                    $user | Add-Member -MemberType NoteProperty -Name AllowIPAudio -Value $userConfPolicy.AllowIPAudio
                    $user | Add-Member -MemberType NoteProperty -Name AllowIPVideo -Value $userConfPolicy.AllowIPVideo
                    $user | Add-Member -MemberType NoteProperty -Name AllowUserToScheduleMeetingsWithAppSharing -Value $userConfPolicy.AllowUserToScheduleMeetingsWithAppSharing
                    $user | Add-Member -MemberType NoteProperty -Name EnableDataCollaboration -Value $userConfPolicy.EnableDataCollaboration            
                }
                else {
                    $user | Add-Member -MemberType NoteProperty -Name AllowIPAudio -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name AllowIPVideo -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name AllowUserToScheduleMeetingsWithAppSharing -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name EnableDataCollaboration -Value $false            
                }
				
                $percent += [int]$subInterval
				LogProgress -activity "Lync Server - User Data Collection" -Status "Checking for Plus CAL requirements" -percentComplete $percent
											
                # Check for Plus CAL requirement
                if ($user.EnterpriseVoiceEnabled -eq $true) {
                    # If true then the user requires Plus CAL, Check with the Voice Policy for appropriate flags
                    $bPlusCAL = $true
                }
        
                if ($user.VoicePolicy -ne $null) {            
					$percent += [int]$subInterval
					LogProgress -activity "Lync Server - User Data Collection" -Status "Checking for VoicePolicy" -percentComplete $percent
								
                    # Get details of the Voice Policy
                    if($ConnectionMethod -eq "SnapIn") {
                        $userVoicePolicy = Get-CsVoicePolicy $user.VoicePolicy
                    }
                    else {
                        $userVoicePolicy = Get-CsVoicePolicy $user.VoicePolicy.FriendlyName
                    }
            
                    $user | Add-Member -MemberType NoteProperty -Name AllowSimulRing -Value $userVoicePolicy.AllowSimulRing
                    $user | Add-Member -MemberType NoteProperty -Name AllowCallForwarding -Value $userVoicePolicy.AllowCallForwarding
                    $user | Add-Member -MemberType NoteProperty -Name AllowPSTNReRouting -Value $userVoicePolicy.AllowPSTNReRouting
                    $user | Add-Member -MemberType NoteProperty -Name EnableDelegation -Value $userVoicePolicy.EnableDelegation
                    $user | Add-Member -MemberType NoteProperty -Name EnableTeamCall -Value $userVoicePolicy.EnableTeamCall
                    $user | Add-Member -MemberType NoteProperty -Name EnableCallTransfer -Value $userVoicePolicy.EnableCallTransfer
                    $user | Add-Member -MemberType NoteProperty -Name EnableCallPark -Value $userVoicePolicy.EnableCallPark
                    $user | Add-Member -MemberType NoteProperty -Name EnableMaliciousCallTracing -Value $userVoicePolicy.EnableMaliciousCallTracing
                    $user | Add-Member -MemberType NoteProperty -Name EnableBWPolicyOverride -Value $userVoicePolicy.EnableBWPolicyOverride
                    $user | Add-Member -MemberType NoteProperty -Name PreventPSTNTollBypass -Value $userVoicePolicy.PreventPSTNTollBypass
                }
                else {
                    $user | Add-Member -MemberType NoteProperty -Name AllowSimulRing -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name AllowCallForwarding -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name AllowPSTNReRouting -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name EnableDelegation -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name EnableTeamCall -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name EnableCallTransfer -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name EnableCallPark -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name EnableMaliciousCallTracing -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name EnableBWPolicyOverride -Value $false
                    $user | Add-Member -MemberType NoteProperty -Name PreventPSTNTollBypass -Value $false
                }
				
                $percent += [int]$subInterval
				LogProgress -activity "Lync Server - User Data Collection" -Status "Estimating CAL requirement" -percentComplete $percent
								
                $CAL = CalculateCAL($bStdCAL, $bEntCAL, $bPlusCAL)

                # Add CAL to CSV    
                $user | Add-Member -MemberType NoteProperty -Name CALRequired -Value $CAL
            }   # End users for loop
    
			LogProgress -activity "Lync Server - Get Users" -Status "Export the result" -percentComplete 95
            $users | Select-Object SamAccountName, UserPrincipalName, FirstName, LastName, `
                WindowsEmailAddress, Sid, LineServerURI, OriginatorSid, AudioVideoDisabled, `
                IPPBXSoftPhoneRoutingEnabled, RemoteCallControlTelephonyEnabled, PrivateLine, `
                HostedVoiceMail, DisplayName, HomeServer, TargetServerIfMoving, EnabledForFederation, `
                EnabledForInternetAccess, PublicNetworkEnabled, EnterpriseVoiceEnabled, EnabledForRichPresence, `
                LineURI, SipAddress, Enabled, TenantId, TargetRegistrarPool, VoicePolicy, MobilityPolicy, `
                ConferencingPolicy, PresencePolicy, RegistrarPool, DialPlan, LocationPolicy, ClientPolicy, `
                ClientVersionPolicy, ArchivingPolicy, PinPolicy, ExternalAccessPolicy, HostedVoicemailPolicy, `
                HostingProvider, Name, DistinguishedName, Identity, Guid, ObjectCategory, `
                WhenChanged, WhenCreated, OriginatingServer, IsValid, ObjectState, AllowIPAudio, AllowIPVideo, `
                AllowUserToScheduleMeetingsWithAppSharing, EnableDataCollaboration, AllowSimulRing, `
                AllowCallForwarding, AllowPSTNReRouting, EnableDelegation, EnableTeamCall, EnableCallTransfer, `
                EnableCallPark, EnableMaliciousCallTracing, EnableBWPolicyOverride, PreventPSTNTollBypass, CALRequired `
                | Export-Csv -NoTypeInformation -Path $OutputFile1 -Encoding UTF8


            # Clear session
            if ($session) { 
                Remove-PsSession $session
            }

		    LogProgress -activity "Lync Server - User Data Collection" -Status "Export - Completed" -percentComplete 100
            Write-Output "Process Completed"
        }
        else {
		    LogProgress -activity "Lync Server - Authentication" -Status "Authentication - Failed" -percentComplete 100
            Write-Output "Process Failed"
        }
	}
    catch {
        # Clear session
        if ($session) { 
            Remove-PsSession $session
        }
		LogLastException
		LogProgress -activity "Script - Exception" -Status "An Error occured" -percentComplete 100 
        Write-Output "Process Failed"
    }
}

# Call the Get-LyncUsers Function to 
#	- Load Account Details
#   - Calculate an approximation on CAL required 
#	- Export CSV
Get-LyncUsers