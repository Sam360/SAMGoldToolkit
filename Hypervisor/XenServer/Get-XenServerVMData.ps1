 ##########################################################################
 # 
 # Get-XenServerVMData
 # SAM Gold Toolkit
 # Original Source: Jon Mulligan (Sam360)
 #
 ##########################################################################
 
 Param(
	[alias("username")]
	$XenUsername = $(Throw "Missing Parameter: Username must be specified"),
	[alias("password")]
	$XenPassword = $(Throw "Missing Parameter: Password must be specified"),
	[alias("server")]
    $XenServer = $(Throw "Missing Parameter: Server must be specified"),
	[alias("o1")]
	$OutputFile1 = "XenHostList" + $XenServer + ".txt",
	[alias("o2")]
	$OutputFile2 = "XenVMList" + $XenServer + ".txt"
    )

<#
.SYNOPSIS
Retrieves physical host and virtual machine data from a XenServer Hypervisor 

.DESCRIPTION
The Get-XenServerVMData script queries a single XenServer hypervisor and produces 2 text files
including virtual machine and physical host details. 
    1)    XenHostList.txt - One record per virtual machine including fields like 
	      VM name, IP, OS, Enabled state, Physical host name etc. The data is retrieved through WMI
    2)    XenVMList.txt - One record per hypervisor. 


If the hypervisor is part of a XenServer farm, details for all VMs and physical hosts in 
the farm are returned.

The script uses the XenServer CLI to retrive the VM data. XenServer CLI must be installed
in its default location ("C:\Program Files (x86)\Citrix\XenCenter\") for this script to work.

.PARAMETER Server 
Host name of XenServer server to scan

.PARAMETER Username
XenServer Username (Required)

.PARAMETER Password
XenServer Password (Required)

.EXAMPLE
Get all guest & host info from from the farm that includes the XenServer hypervisor 'Omaha'. 
Get-XenServerVMData –VMserver Omaha

.NOTES

#>

$xeExe = "C:\Program Files (x86)\Citrix\XenCenter\xe.exe"

& $xeExe -s $XenServer -u $XenUsername -pw $XenPassword host-list params=all 2>&1 | out-file $OutputFile1
& $xeExe -s $XenServer -u $XenUsername -pw $XenPassword vm-list params=all 2>&1 | out-file $OutputFile2
