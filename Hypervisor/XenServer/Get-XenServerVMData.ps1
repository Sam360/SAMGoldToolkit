 ##########################################################################
 # 
 # Get-XenServerVMData
 # SAM Gold Toolkit
 # Original Source: Sam360
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
	$OutputFile1 = "XenHostList" + $XenServer + ".csv",
	[alias("o2")]
	$OutputFile2 = "XenVMList" + $XenServer + ".csv"
    )

$xeExe = "C:\Program Files (x86)\Citrix\XenCenter\xe.exe"

& $xeExe -s $XenServer -u $XenUsername -pw $XenPassword host-list params=all 2>&1 | out-file $OutputFile1
& $xeExe -s $XenServer -u $XenUsername -pw $XenPassword vm-list params=all 2>&1 | out-file $OutputFile2
