 ##########################################################################
 # 
 # Get-VMwareVMList
 # Sam Gold Toolkit
 # Original Source: Sam360
 #
 ##########################################################################
 

Param(
	[switch]$TestCredentials,
	[alias("o1")]
    $OutputFile1 = "VMwareData.csv",
	[alias("username")]
	$VMwareUsername = $(Throw "Missing Parameter: Username must be specified"),
	[alias("password")]
	$VMwarePassword = $(Throw "Missing Parameter: Password must be specified"),
	[alias("server")]
	$VMserver = $(Throw "Missing Parameter: Server must be specified"))
	

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

function LogProgress($progressDescription){
    $output = Get-Date -Format HH:mm:ss.ff
	$output += " - "
	$output += $progressDescription
    write-output $output
}

function Get-VMwareVMList
{
    LogProgress "importing modules"
	Add-PSSnapin VMware.DeployAutomation
	Add-PSSnapin VMware.ImageBuilder
	Add-PSSnapin VMware.VimAutomation.Core
	Add-PSSnapin VMware.VimAutomation.License
	Add-PSSnapin VMware.VimAutomation.Vds

    Write-Output "Script Parameters"
    Write-Output "UserName:   $VMwareUsername"
    Write-Output "OutputFile: $OutputFile2"
	Write-Output "Server:     $VMServer"

    # Connect to the VMware server
    try
    {
        LogProgress "Connecting to VMware server"
        $viServer = Connect-VIServer -Server $VMServer -User $VMwareUsername -Password $VMwarePassword
		
		if($viServer){
			LogProgress "Connected to VMware server"
		}
		else {
			Throw "VMware logon failed"
		}
    }
    catch
    {
        LogLastException
        return
    }
	
	if ($TestCredentials)
	{
		# We're only testing credentials
		return
	}

    # Get the VM Data
    try
    {
        LogProgress "Getting VM Info"

        $VmInfo = ForEach ($Datacenter in (Get-Datacenter | Sort-Object -Property Name)) {
            ForEach ($VM in ($Datacenter | Get-VM | Sort-Object -Property Name)) {
                
                $IPAddresses = "";
                $MACAddresses = "";
                $NetworkNames = "";
                foreach ($NIC in $VM.Guest.Nics) {
                    $IPAddresses += $NIC.IPAddress -join ','
                    $MACAddresses += $NIC.MacAddress
                    $NetworkNames += $NIC.NetworkName
                
                    $IPAddresses += ','
                    $MACAddresses += ','
                    $NetworkNames += ','
                }
                
                "" | Select-Object -Property @{N="VM";E={$VM.Name}},
                @{N="VMCPUCount";E={$vm.ExtensionData.Config.Hardware.NumCPU/$vm.ExtensionData.Config.Hardware.NumCoresPerSocket}},
                @{N="VMCPUCoreCount";E={$vm.NumCPU}},
                @{N="PowerState";E={$VM.PowerState}},
                @{N="FaultToleranceState";E={$VM.ExtensionData.Summary.Runtime.FaultToleranceState}},
                @{N="OnlineStandby";E={$VM.ExtensionData.Summary.Runtime.OnlineStandby}},
                @{N="Version";E={$VM.Version}},
                @{N="Description";E={$VM.Description}},
                @{N="Notes";E={$VM.Notes}},
                @{N="MemoryMB";E={$VM.MemoryMB}},
                @{N="ResourcePool";E={$VM.ResourcePool}},
                @{N="ResourcePoolID";E={$VM.ResourcePoolId}},
                @{N="PersistentID";E={$VM.PersistentId}},
                @{N="ID";E={$VM.Id}},
                @{N="UID";E={$VM.Uid}},
                @{N="UUID";E={$VM.ExtensionData.Config.Uuid}},
                @{N="IPs";E={$IPAddresses}},
                @{N="MACs";E={$MACAddresses}},
                @{N="NetworkNames";E={$NetworkNames}},
                @{N="OS";E={$VM.Guest.OSFullName}},
                @{N="FQDN";E={$VM.Guest.HostName}},
                @{N="ScreenDimensions";E={$VM.Guest.ScreenDimensions}},
                @{N="OS2";E={$VM.ExtensionData.Guest.OSFullName}},
                @{N="FQDN2";E={$VM.ExtensionData.Guest.HostName}},
                @{N="GuestHostname";E={$VM.ExtensionData.Summary.Guest.HostName}},
                @{N="GuestId";E={$VM.ExtensionData.Summary.Guest.GuestId}},
                @{N="GuestFullName";E={$VM.ExtensionData.Summary.Guest.GuestFullName}},
                @{N="GuestIP";E={$VM.ExtensionData.Summary.Guest.IpAddress}},
                @{N="Datacenter";E={$Datacenter.name}},
                @{N="Cluster";E={$vm.VMHost.Parent.Name}},
                @{N="HostName1";E={$vm.VMHost.Name}},
                @{N="HostName2";E={$vm.VMHost.ExtensionData.Summary.Config.Name}},
                @{N="HostName3";E={$VM.VMHost.NetworkInfo.HostName}},
                @{N="HostDomainName";E={$VM.VMHost.NetworkInfo.DomainName}},
                @{N="HostCPUCount";E={$vm.VMHost.ExtensionData.Summary.Hardware.NumCpuPkgs}},
                @{N="HostCPUCoreCount";E={$vm.VMHost.ExtensionData.Summary.Hardware.NumCpuCores/$vm.VMHost.ExtensionData.Summary.Hardware.NumCpuPkgs}},
                @{N="HyperthreadingActive";E={$vm.VMHost.HyperthreadingActive}},
                @{N="HostManufacturer";E={$VM.VMHost.Manufacturer}},
                @{N="HostID";E={$VM.VMHost.Id}},
                @{N="HostUID";E={$VM.VMHost.Uid}},
                @{N="HostPowerState";E={$VM.VMHost.PowerState}},
                @{N="HostHyperThreading";E={$VM.VMHost.HyperthreadingActive}},
                @{N="HostvMotionEnabled";E={$vm.VMHost.ExtensionData.Summary.Config.VmotionEnabled}},
                @{N="HostVendor";E={$VM.VMHost.ExtensionData.Hardware.SystemInfo.Vendor}},
                @{N="HostModel";E={$VM.VMHost.ExtensionData.Hardware.SystemInfo.Model}},
                @{N="HostRAM";E={$VM.VMHost.ExtensionData.Summary.Hardware.MemorySize}},
                @{N="HostCPUModel";E={$VM.VMHost.ExtensionData.Summary.Hardware.CpuModel}},
                @{N="HostThreadCount";E={$VM.VMHost.ExtensionData.Summary.Hardware.NumCpuThreads}},
                @{N="HostCPUSpeed";E={$VM.VMHost.ExtensionData.Summary.Hardware.CpuMhz}},
                @{N="HostProductName";E={$VM.VMHost.ExtensionData.Summary.Config.Product.Name}},
                @{N="HostProductVersion";E={$VM.VMHost.ExtensionData.Summary.Config.Product.Version}},
                @{N="HostProductBuild";E={$VM.VMHost.ExtensionData.Summary.Config.Product.Build}},
                @{N="HostProductOS";E={$VM.VMHost.ExtensionData.Summary.Config.Product.OsType}},
                @{N="HostProductLicenseKey";E={$VM.VMHost.LicenseKey}},
                @{N="HostProductLicenseName";E={$VM.VMHost.ExtensionData.Summary.Config.Product.LicenseProductName}},
                @{N="HostProductLicense Version";E={$VM.VMHost.ExtensionData.Summary.Config.Product.LicenseProductVersion}},
                @{N="HostvMotionIPAddress";E={$VM.VMHost.ExtensionData.Config.vMotion.IPConfig.IpAddress}},
                @{N="HostvMotionSubnetMask";E={$VM.VMHost.ExtensionData.Config.vMotion.IPConfig.SubnetMask}}   
            }
          }
        $VmInfo | Export-Csv -NoTypeInformation -Path $OutputFile1
    }
    catch
    {
        LogLastException
    }
}

Get-VMwareVMList


