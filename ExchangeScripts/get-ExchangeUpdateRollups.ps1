# Get-ExchangeUpdateRollups.ps1 - Gets the Exchange Server 2007, Exchange 2010 and Exchange 2013 Update Rollups

<# 
.SYNOPSIS
Get-ExchangeUpdateRollups.ps1 - Gets the Exchange Server 2007, Exchange 2010 and Exchange 2013 Update Rollups
.NOTES
Written By: Bhargav Shukla
Website:	http://www.bhargavs.com
Revised By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
* 8:55 AM 11/29/2016 add name sort (make them output into a rational order)
* 8:54 AM 11/29/2016 MyAccount: fixed typo in $ofile gen
* 8:52 AM 11/29/2016 MyAccount: added Ex Org, progress verbose outputs 
* 8:23 AM 11/29/2016 MyAccount: added/reformated comments into pshelp, 
	ren chg rpt file results.csv => $ofile (scriptname-report-timestamped.csv)
	added $ofile invoke item, reformated to OTB, 
* Posted version:"UPDATED: Feb 26, 2014, updated script to accommodate Exchange 2013
The script was not written to accommodate Exchange 2013 as it doesn’t use RU. 
It uses different servicing model. I have updated script to work with 2013 but 
it won’t report CU, it will provide build numbers that can be matched to CUs 
published on TechNet. Links are included in the script."
.DESCRIPTION
Get-ExchangeUpdateRollups.ps1
# Gets the Exchange Server 2007, Exchange 2010 and Exchange 2013 Update Rollups
# installed writes output to CSV file in same folder where script is called from
#
# Exchange 2013 CU Build Numbers - http://social.technet.microsoft.com/wiki/contents/articles/15776.exchange-server-2013-and-cumulative-updates-cus-build-numbers.aspx
# Exchange Server Update Rollups and Build Numbers - http://social.technet.microsoft.com/wiki/contents/articles/240.exchange-server-and-update-rollups-build-numbers.aspx
#
# This script won't report RUs for Exchange Server 2013 since it uses Cummulative Updates (CU).
# More details on Exchange Team Blog: Servicing Exchange 2013
# http://blogs.technet.com/b/exchange/archive/2013/02/08/servicing-exchange-2013.aspx
 DISCLAIMER
# ==========
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
# RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Outputs a csv file in the scripts subdir of the 
.EXAMPLE
.\get-ExchangeUpdateRollups.ps1
Default run against all Ex07|10|13 servers in the Org
.LINK
http://www.bhargavs.com/index.php/2009/12/14/how-do-i-check-update-rollup-version-on-exchange-20xx-server/

#>

# 8:32 AM 11/29/2016 add simple code for report based on scriptname
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
$ofile = (join-path -path $ScriptDir -childpath "logs") ; 
if(!(test-path -path $ofile)){write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Creating missing logs dir $($ofile)" ; mkdir -path $ofile -force ; } ; 
$ofile+="\$($ScriptNameNoExt)-$(get-date -uformat "%Y%m%d-%H%M")-Report.csv" ;
write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):`ofile:$($ofile)" ;

# Store header in variable
$headerLine = 
@"
Exchange 2013 CU Build Numbers - http://social.technet.microsoft.com/wiki/contents/articles/15776.exchange-server-2013-and-cumulative-updates-cus-build-numbers.aspx
Exchange Server Update Rollups and Build Numbers - http://social.technet.microsoft.com/wiki/contents/articles/240.exchange-server-and-update-rollups-build-numbers.aspx
 
Server Name,Rollup Update Description,Installed Date,ExSetup File Version
"@
 
# Write header to file
#$headerLine | Out-File .\results.csv -Encoding ASCII -Append
# 9:06 AM 11/29/2016 update to $ofile
$headerLine | Out-File -filepath $ofile -Encoding ASCII -Append ; 

#*------v Function getRU v------ 
function getRU([string]$Server) {
# Set server to connect to
	$Server = $Server.ToUpper()
	
	# Check if server is running Exchange 2007 or Exchange 2010

	$ExchVer = (Get-ExchangeServer $Server | ForEach {$_.AdminDisplayVersion})
	$sMsg=$Server ; 
	# Set appropriate base path to read Registry
	# Exit function if server is not running Exchange 2007 or Exchange 2010
	# 8:49 AM 11/29/2016 added version echo - this runs awhile with no progress/status updates
    if ($ExchVer -match "Version 15") {
		$sMsg+="(Exch2013)" ; 
		$REG_KEY = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
		$Reg_ExSetup = "SOFTWARE\\Microsoft\\ExchangeServer\\v15\\Setup"
	} elseif ($ExchVer -match "Version 14") {
		$sMsg+="(Exch2010)" ; 
		$REG_KEY = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
		$Reg_ExSetup = "SOFTWARE\\Microsoft\\ExchangeServer\\v14\\Setup"
	} elseif	($ExchVer -match "Version 8") {
		$sMsg+="(Exch2007)" ; 
		$REG_KEY = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\461C2B4266EDEF444B864AD6D9E5B613\\Patches"
		$Reg_ExSetup = "SOFTWARE\\Microsoft\\Exchange\\Setup"
	} else {
		return
	}
	
	# 8:48 AM 11/29/2016 add progress echo
	write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Processing Server $($sMsg)..." ;

	# Read Rollup Update information from servers
	# Set Registry constants
	$VALUE1 = "DisplayName"
	$VALUE2 = "Installed"
	$VALUE3 = "MsiInstallPath"

	# Open remote registry
	$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)

	# Set regKey for MsiInstallPath
	$regKey= $reg.OpenSubKey($REG_ExSetup)

	# Get Install Path from Registry and replace : with $
	$installPath = ($regkey.getvalue($VALUE3) | foreach {$_ -replace (":","`$")})
	
	# Set ExSetup.exe path
	$binFile = "Bin\ExSetup.exe"
	
	# Get ExSetup.exe file version
	$exSetupVer = ((Get-Command "\\$Server\$installPath$binFile").FileVersionInfo | ForEach {$_.FileVersion})

	# Create an array of patch subkeys
	$regKey= $reg.OpenSubKey($REG_KEY).GetSubKeyNames() | ForEach {"$Reg_Key\\$_"}

	# Walk through patch subkeys and store Rollup Update Description and Installed Date in array variables
	$dispName = [array] ($regkey | %{$reg.OpenSubKey($_).getvalue($VALUE1)})
	$instDate = [array] ($regkey | %{$reg.OpenSubKey($_).getvalue($VALUE2)})

	# Loop Through array variables and output to a file
	$countmembers = 0
	
	if ($regkey -ne $null) {
		while ($countmembers -lt $dispName.Count){
			#$server+","+$dispName[$countmembers]+","+$instDate[$countmembers].substring(0,4)+"/"+$instDate[$countmembers].substring(4,2)+"/"+$instDate[$countmembers].substring(6,2)+","+$exsetupver | 
			#	Out-File .\results.csv -Encoding ASCII -Append
			# 8:40 AM 11/29/2016 shift to report from scriptname
			$server+","+$dispName[$countmembers]+","+$instDate[$countmembers].substring(0,4)+"/"+$instDate[$countmembers].substring(4,2)+"/"+$instDate[$countmembers].substring(6,2)+","+$exsetupver | 
				Out-File -filepath $ofile -Encoding ASCII -Append
			$countmembers++
		} ; 
	} else {
		#$server+",No Rollup Updates are installed,,"+$exsetupver | Out-File .\results.csv -Encoding ASCII -Append
		# 8:40 AM 11/29/2016 shift to report from scriptname 
		$server+",No Rollup Updates are installed,,"+$exsetupver | Out-File -filepath $ofile -Encoding ASCII -Append

	}
} #*------^ END Function getRU ^------

# Get Exchange 2007/2010 servers and write Rollup Updates to results file
# 8:51 AM 11/29/2016 add org Name
$exOrg=(get-organizationconfig).Name ; 
write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Collecting Exchange Servers in Org:$($exOrg)..." ;
#$Servers = (Get-ExchangeServer | Where-Object {($_.AdminDisplayVersion -match "Version 8" -OR $_.AdminDisplayVersion -match "Version 14" -OR $_.AdminDisplayVersion -match "Version 15") -AND $_.ServerRole -ne "ProvisionedServer" -and $_.ServerRole -ne "Edge"} | ForEach {$_.Name})
# 8:55 AM 11/29/2016 add name sort (make them output into a rational order)
$Servers = (Get-ExchangeServer | Where-Object {($_.AdminDisplayVersion -match "Version 8" -OR $_.AdminDisplayVersion -match "Version 14" -OR $_.AdminDisplayVersion -match "Version 15") -AND $_.ServerRole -ne "ProvisionedServer" -and $_.ServerRole -ne "Edge"} | 
	sort Name | ForEach {$_.Name})
$Servers | ForEach {getRU $_} ; 
#Write-Output "Results are stored in $(Get-Location)\results.csv"
# 8:40 AM 11/29/2016 shift to report from scriptname & invoke-item it
Write-Output "Results are stored in $($ofile)" ; 
if(get-childitem -path $ofile){Invoke-Item -path $ofile ;} ;  


