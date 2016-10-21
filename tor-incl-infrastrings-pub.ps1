# c:\usr\work\exch\scripts\tor-incl-infrastrings.ps1

#*----------V Comment-based Help (leave blank line below) V---------- 

<# 
    .SYNOPSIS
tor-incl-infrastrings.ps1 - included server infra strings file

    .NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Additional Credits: [REFERENCE]
Website:	[URL]
Twitter:	[URL]

Change Log
# 9:01 AM 1/7/2015 initial version

To implement, copy the include files to the $LocalInclDir & $LocalInclSIDDir, and then add this to $profile bottom:
#-----------
$DomainWork = "DOMAIN";
$DomHome = "DOMAIN";
$DomLab="DOMAIN";

$LocalInclDir="c:\usr\home\db" ;
$LocalInclSIDDir = "c:\usr\work\ps\scripts\";
# distrib shares:
$InclShareCent = "\\Server3.domain.com\USR$\MyAccount\dev\ps\scripts";
$InclShareL13 = "\\Server0.domain.com\SITEc_FS\scripts";
$inclShareLab = "\\Server0\e$\scripts";
$InclSIDDirL = "\\Server0.domain.com\SITEc_FS\scripts";
$inclShareLab = "\\Server0\e$\scripts\";

# profile search path
if (!(Test-Path $LocalInclDir)) {
  if (Test-Path $InclShareCent) {
    $LocalInclDir = $InclShareCent ; Write-Verbose -verbose "Using CentralShare Includes:$LocalInclDir";
  } elseif ($env:USERDOMAIN -eq "DOMAIN") {
    $LocalInclDir = $inclShareLab ; Write-Verbose -verbose "Using LabShare Includes:$LocalInclDir";
  } elseif (Test-Path $InclShareL13) {
    $LocalInclDir = $InclShareL13 ; Write-Verbose -verbose "Using L13 Includes:$LocalInclDir";
  } else {
    Write-Warning "No available include source found. Exiting..."; Exit;
  };
};

Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":LOADING INCLUDES:")
$sLoad=(join-path -path $LocalInclDir -childpath "tor-incl-infrastrings.ps1") ; ;if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;};
#---------

    .DESCRIPTION
tor-incl-infrastrings.ps1 - included server infra strings file

    .INPUTS
None. Does not accepted piped input.

    .OUTPUTS
None. Returns no objects or output.


*----------^ end Comment-based help  ^---------- #>

# generic block of standardized paths and arrays of objects for the enterprise 
#   (also avoids need to use EMS to dynamically query Exchange roles and components)
# Change Log
# 8:48 AM 1/7/2015 fixed missing $ on $SitesNameList
# 12:37 PM 12/4/2014 built $ServersMailALL list
# 12:03 PM 10/1/2014 added $L13FEIisInt & $L13FEIisExt
# 11:03 AM 9/30/2014 added L13sql & l13owas, fixed typo in L13Edg
# 2:16 PM 9/12/2014 corrected typo in $ServersMbxALL
# 12:35 PM 9/12/2014 - added SITEc servers
# 10:44 AM 12/23/2013 - TOR revision, relays rem'd
# 9:43 AM 10/25/2013 - initial revision, shifting the DELIMITED UNI CONSTANTS block out of each script to here
# *----------------v DELIMITED TOR CONSTANTS v----------------
$SitesNameList = "SiteName;SiteName;SiteName" ; $SitesNameList=$SitesNameList.split(";") ; 
$SitesList = "*SiteName*;*SiteName*;*SiteName*" ; $SitesList=$SitesList.split(";") ; 
$ServersMbxNA="Server1" ; $ServersMbxNA=$ServersMbxNA.split(";") ; 
$ServersHCNA="Server1" ; $ServersHCNA=$ServersHCNA.split(";") ; 
$ServersMbxALL="Server1" ; $ServersMbxALL=$ServersMbxALL.split(";") ;
$ServersMbxAU="Server1" ; $ServersMbxAU=$ServersMbxAU.split(";") ; 
$ServersHCAU="Server1" ; $ServersHCAU=$ServersHCAU.split(";") ; 
$ServersMbxEU="Server1" ; $ServersMbxEU=$ServersMbxEU.split(";") ; 
$ServersHCEU="Server1" ; $ServersHCEU=$ServersHCEU.split(";") ; 
# 12:37 PM 12/4/2014 built allmail list
$ServersMailALL="Server1" ; $ServersMailALL =$ServersMailALL.split(";") ;
$IisLogsHCNA="\\Server1\" ; $IisLogsHCNA=$IisLogsHCNA.split(";") ; 
$IisLogsHCAU="\\Server1\" ; $IisLogsHCAU=$IisLogsHCAU.split(";") ; 
$IisLogsHCEU="\\Server1\" ; $IisLogsHCEU=$IisLogsHCEU.split(";") ; 
$MsgTrkLogsHCNA="\\Server4\TransportRoles\Logs\MessageTracking\" ; $MsgTrkLogsHCNA=$MsgTrkLogsHCNA.split(";") ; 
$MsgTrkLogsHCEU="\\Server4\TransportRoles\Logs\MessageTracking\" ; $MsgTrkLogsHCEU=$MsgTrkLogsHCEU.split(";") ; 
$MsgTrkLogsHCAU="\\Server4\TransportRoles\Logs\MessageTracking\" ; $MsgTrkLogsHCAU=$MsgTrkLogsHCAU.split(";") ;
$MsgTrkLogsHCNATemplate="\\Server\C$\Program Files\Microsoft\Exchange Server\V14\TransportRoles\Logs\MessageTracking\" ; 
$MsgTrkLogsHCAUTemplate="\\Server4\TransportRoles\Logs\MessageTracking\" ;
$MsgTrkLogsHCEUTemplate="\\Server4\TransportRoles\Logs\MessageTracking\" ;
$L13LabFE="Server2";
$L13LabFE=$L13LabFE.split(";") ;
$L13LabEDG="Server1";
$L13LabEDG=$L13LabEG.split(";") ;
$L13LabALL="Server1";
$L13LabALL =$L13LabALL.split(";") ;
# prod SITEc 2013
$L13FE="Server3FE.split(";") ; 
$L13EDG="Server3EDG.split(";") ;
$L13ALL="Server3ALL.split(";") ;
$L13FEfqdn="Server3FEfqdn.split(";") ;
$L13EDGfqdn="Server3EDGfqdn.split(";") ;
$L13ALLfqdn="Server3ALLfqdn.split(";") ;
$L13SQL="Server3SQL.split(";") ;
$L13OWAS="Server3OWAS.split(";") ;
$L13FEIisInt="\\Server3FEIisInt.split(";");
$L13FEIisExt="\\Server3FEIisExt.split(";");
# *----------------^ DELIMITED TOR CONSTANTS ^----------------


