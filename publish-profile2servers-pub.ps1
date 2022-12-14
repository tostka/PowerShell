#c:\usr\work\exch\scripts\publish-profile2servers.ps1
# push the local $profile copy in uwes to the full range of Ex & L13 servers
# debug command syntax: Clear-Host ; .\publish-profile2servers.ps1 -whatif

#*----------V Comment-based Help (leave blank line below) V---------- 

<# 
.SYNOPSIS
publish-profile2servers.ps1 - push the local $profile copy in uwes to the full range of Ex & L13 servers
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Additional Credits: [REFERENCE]
Website:	[URL]
Twitter:	[URL]

Change Log
* 10:24 AM 6/12/2015 - functionalized, rewrote
# 8:00 AM 6/12/2015 switched to $SourceProfileMachine & $AcctWAdmn varis
# 8:25 AM 1/8/2015 added SID test
#12:49 PM 12/31/2014 initial version 

.DESCRIPTION
.PARAMETER Whatif
Parameter to run a Test no-change pass, and log results [-Whatif switch]
.PARAMETER ShowProgress
Parameter to display progress meter [-ShowProgress switch]
.PARAMETER ShowDebug
Parameter to display Debugging messages [-ShowDebug switch]
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.\publish-profile2servers.ps1 
.EXAMPLE
.\publish-profile2servers.ps1 -whatif
Whatif test pass
.LINK
*----------^ END Comment-based Help  ^---------- #>
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

[CmdletBinding()]
  
PARAM (
    [parameter(Position=1,Mandatory=$False,
        HelpMessage="Credentials under which to run profile copy process")]
    [System.Management.Automation.PSCredential]$Credential
,
    [Parameter(HelpMessage='ShowProgress [$switch]')]
    [switch] $showProgress
,
    [Parameter(HelpMessage='Debugging Flag [$switch]')]
    [switch] $showDebug
,
    [Parameter(HelpMessage='Whatif Flag  [$switch]')]
    [switch] $whatIf

) ;  # PARAM-E

$SourceProfileMachine="MyComputer"
$AdminLogon="MyAccount" ; 
$AdminDomain="DOMAIN" ; 
$TargetProfileAcct="$($AdminDomain)\$($AdminLogon)";

#*================v BOILERPLATE SCRIPT-REFERENCES v================
#8:31 AM 4/24/2015 add a Verbose vari, to permit programatic control off all the -Verbose:$true entries (as -Verbose:$verbose)
$verbose = $TRUE
# 12:23 PM 2/20/2015 add gui vb prompt support
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null ; 
# 11:00 AM 3/19/2015 should use Windows.Forms where possible, more stable

# 2:10 PM 2/4/2015 shifted to here to accommodate include locations
$scriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
# 2:12 PM 3/25/2015 moved below to script head to make avail for html footer boilerplate
# fr 2266
$scriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName)) ;
# fr 2248
$scriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
$timeStampNow = get-date -uformat "%Y%m%d-%H%M" ;


# 11:14 AM 4/2/2015 start banner
Write-Verbose "---------------------------------------------------------------------------------" -Verbose:$verbose
Write-Verbose "START ==== $scriptBaseName ====" -Verbose:$verbose

# 12:48 PM 3/11/2015 detect -noprofile runs (in case you need to add profile content/functions to get to function)
$noProf=[bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); 
# if($noProf){# do this};

$myBox="MyComputer","MyComputer"
$domainWork = "DOMAIN";
$domHome = "DOMAIN";
$domLab="DOMAIN";

# 2:01 PM 2/5/2015
$sQot = [char]34
$sQotS = [char]39

#*================^ END BOILERPLATE SCRIPT-REFERENCES ^================


if (([System.Security.Principal.WindowsIdentity]::GetCurrent().Name.tolower()) -ne $TargetProfileAcct) {write-warning "This script must be run SID! Exiting..."; play-beep ;Exit };

#*================v FUNCTIONS  v================

#*----------------v Function Get-TimeStamp v----------------
function Get-TimeStamp {

    <# 
    .SYNOPSIS
    Get-TimeStamp - Return "HH:mm:ss" Timestamp
    .NOTES
    Written By: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka

    Change Log
    * 11:23 AM 3/19/2015 cleanup added pshelp
    * 2:06 PM 12/3/2014 fixed, lms says  non-existent func
    * 8:43 AM 10/31/2014 - simple timestamp echo

    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns "HH:mm:ss" Timestamp
    .EXAMPLE
    Write "$(Get-TimeStamp):MESSAGE"
    .EXAMPLE
    write-host "$(Get-TimeStamp):MESSAGE" �verbose ;
    .EXAMPLE
    write-host -foregroundcolor yellow  "$(Get-TimeStamp):MESSAGE";
    .LINK
    *----------^ END Comment-based Help  ^---------- #>

	Get-Date -Format "HH:mm:ss";	
} #*----------------^ END Function Get-TimeStamp ^----------------

#*----------------v Function copy-Profile v------
function copy-Profile {
  #*----------V Comment-based Help (leave blank line below) V----- 

  <#
  .SYNOPSIS
  copy-Profile() - Copies $SourceProfileMachine WindowsPowershell profile dir to the specified machine(s)

  .NOTES
  Written By: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka

  Change Log
  8:07 AM 6/12/2015 - functionalize copy code from the EMS block

  .PARAMETER  ComputerName
  Name or IP address of the target computer
  .PARAMETER SourceProfileMachine
  Source Name or IP address of the source Profile computer
  .PARAMETER TargetProfile
  Target Account for Profile copy process [domain\logon]
  .PARAMETER showProgress
  Show Progress bar reflecting progress toward completion
  .PARAMETER showDebug
  Show Debugging messages
  .PARAMETER whatIf
  Execute solely a test pass
  .PARAMETER Credential
  Credential object for use in accessing the computers.
  .INPUTS
  Accepts piped input.
  .OUTPUTS
  Returns an object with uptime data to the pipeline.
  .EXAMPLE
  copy-Profile USEA-MAILEXP | select Computername,Uptime
  *----------^ END Comment-based Help  ^---------- #>

  #[CmdletBinding()]
  
  PARAM (
      [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0,
      Mandatory=$True,HelpMessage="Specify Target Computer for Profile Copy[ServerName]")]
      [Alias('__ServerName','Server','Computer','Name','IPAddress','CN')]   
      [string[]]$ComputerName = $env:COMPUTERNAME
  ,
      [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Source Profile Machine [ServerName]")]
      [ValidateNotNullOrEmpty()]
      [string]$SourceProfileMachine
  ,    
      [parameter(Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Target Account for Profile copy process [domain\logon]")]
      [ValidateNotNullOrEmpty()]
      [string]$TargetProfile
  ,      
      [parameter(HelpMessage="Credential object for use in accessing the computers")]
      [System.Management.Automation.PSCredential]$Credential
  ,
      [Parameter(HelpMessage='ShowProgress [$switch]')]
      [switch] $showProgress
,
      [Parameter(HelpMessage='Debugging Flag [$switch]')]
      [switch] $showDebug
,
      [Parameter(HelpMessage='Whatif Flag  [$switch]')]
      [switch] $whatIf
  ) ;  # PARAM-E
  
  BEGIN {
    if ($Credential -ne $Null) {
      $WmiParameters.Credential = $Credential ; 
    }  ;  # if-E    
    
    if ($showDebug) {
        write-host "`$showDebug is $true. `nSetting `$DebugPreference = 'Continue'" ; 
        $DebugPreference = "Continue" ; 
        $bDebug=$true
    };
    If ($whatIf){write-host "`$whatIf is $true" ; $bWhatif=$true}; 
    #$bDebug = $false;
    #$bDebug = $true;


    $progInterval= 500 ; # write-progress wait interval in ms
    # parse out params
    #$SourceProfileMachine="MyComputer"
    $AdminLogon=$TargetProfile.split("\")[1] ;
    #"MyAccount" ; 
    $AdminDomain=$TargetProfile.split("\")[0] ;
    #"DOMAIN" ; 
    $TargetProfileAcct=$TargetProfile ;
    #"$($AdminDomain)\$($AdminLogon)";
  
    $Info = @() 
    $iProcd=0; 
  }  # BEG-E
  PROCESS {
    # always foreach, to accommodate arrays passed in
    foreach ($Computer in $Computername) {
      $iProcd++
      $continue = $true
      try {
        $ErrorActionPreference = "Stop" ;
        # =========== PROCESSING BLOCK ===========
        # COLLECT YOUR DATA OBJECTS HERE
        
        write-verbose -verbose:$true "$(get-timestamp):Processing: $($Computer)..." ;
        if(!(test-path "\\$Computer\c$\Users\$AdminLogon\Documents\WindowsPowerShell")) {
            new-item -path \\$Computer\c$\Users\$AdminLogon\Documents\WindowsPowerShell -itemtype directory -Force #-whatif
        }; # if-E
        
        if($whatIf){
          copy "\\$SourceProfileMachine\c$\usr\work\exch\scripts\profile.ps1" "\\$Computer\c$\Users\$AdminLogon\Documents\WindowsPowerShell\" -whatif ;
          # get the ISE profile too
          copy "\\$SourceProfileMachine\c$\usr\work\exch\scripts\Microsoft.PowerShellISE_profile.ps1" "\\$Computer\c$\Users\$AdminLogon\Documents\WindowsPowerShell\" -whatif ;

        } else {
          copy "\\$SourceProfileMachine\c$\usr\work\exch\scripts\profile.ps1" "\\$Computer\c$\Users\$AdminLogon\Documents\WindowsPowerShell\" #-whatif ;
          # get the ISE profile too
          copy "\\$SourceProfileMachine\c$\usr\work\exch\scripts\Microsoft.PowerShellISE_profile.ps1" "\\$Computer\c$\Users\$AdminLogon\Documents\WindowsPowerShell\" #-whatif ;
        }  # if-E
        
        # dumps confirmation into the pipeline, returns at call point
        Write-Output $true

      } catch {
        # BOILERPLATE ERROR-TRAP
        Write "$(Get-TimeStamp): -- SCRIPT PROCESSING CANCELLED"
        Write "$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName)."
        Write "$(Get-TimeStamp): -- Error information"
        Write "$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)"
        Write "$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)"
        Write "$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)"
        Write "$(Get-TimeStamp): Line: $($_.InvocationInfo.Line)"
        Write "$(Get-TimeStamp): Error Details: $($_)"
        # dumps fail into the pipeline, returns at call point
        Write-Output $false
        Continue
        # Exit; here if you want processing to die and not continue on next for-pass
      } # try/cat-E
    } # loop-E
  }# BPE-E
} #*----------------^ END Function copy-Profile ^--------

#*================^ END FUNCTIONS  ^================

#*----------------v SUB MAIN v----------------

# *** REGION MARKER LOAD
#region LOAD
# *** LOADING 

#if (($env:COMPUTERNAME -ne '$SourceProfileMachine')) {write-warning "This machine is not appropriate for uploading profiles to SITEc Servers. Exiting..."; write-host "`a"; Exit;} ;

#write-verbose -verbose "Y"
$ModsLoad=Get-Module ; # 12:24 PM 12/31/2014 this CAN'T pickup on the Ems under 2010 - because EMS is REMOTED!
$EMSSess= (Get-PSSession | where { $_.ConfigurationName -eq 'Microsoft.Exchange' }) ;
<#Detect EMS10 (remote session):
if (get-PSSession | where {$_.ConfigurationName -eq 'Microsoft.Exchange'}) {[commands for the remote session] }

#>

#$SnapsLoad=Get-PSSnapin ;($SnapsLoad | where {$_.Name -like "Microsoft.Exchange.Management.PowerShell*"})
#$SnapsLoad=Get-PSSnapin ;

if ( (Get-PSSession | where {$_.ConfigurationName -eq 'Microsoft.Exchange'}) ){
  write-verbose -verbose:$true "$(get-timestamp):Copying $($SourceProfileMachine) machine WindowsPowerShell dir to All Mbx & HT servers)"

} else {
    $sName="Microsoft.Exchange.Management.PowerShell*"; if (!(Get-PSSnapin | where {$_.Name -eq $sName})) {Add-PSSnapin $sName -ea Stop};
}  
  
$Exs=(get-exchangeserver | ?{(($_.IsMailboxServer) -OR ($_.IsHubTransportServer))} )
if($Exs){
    if($whatIf){ 
        write-verbose -verbose:$true "$(get-timestamp):Whatif test...)"
        copy-Profile -ComputerName $Exs -SourceProfileMachine $SourceProfileMachine -TargetProfile $TargetProfileAcct -showDebug -whatIf
    } else {
       copy-Profile -ComputerName $Exs -SourceProfileMachine $SourceProfileMachine -TargetProfile $TargetProfileAcct -showDebug
    }
} else {
write-verbose -verbose:$true "$(get-timestamp):No Mbx or HT servers found)"
} # if-E

# ========
$sName="Lync"; if (!(Get-Module | where {$_.Name -eq $sName})) {Import-Module $sName -ea Stop};
$FEs = (get-cspool server.domain.com).computers ; 
write-verbose -verbose:$true "$(get-timestamp):Copying $() machine WindowsPowerShell dir to All SITEc FE servers)" ; 
if($FEs){
    if($whatIf){ 
        write-verbose -verbose:$true "$(get-timestamp):Whatif test...)"
        copy-Profile -ComputerName $FEs -SourceProfileMachine $SourceProfileMachine -TargetProfile $TargetProfileAcct -showDebug -whatIf
    } else {
        copy-Profile -ComputerName $FEs -SourceProfileMachine $SourceProfileMachine -TargetProfile $TargetProfileAcct -showDebug
    } # if-E
} else {
    write-verbose -verbose:$true "$(get-timestamp):No Mbx or HT servers found)"
} # if-E
#} # if-E LMS present

# 11:14 AM 4/2/2015 end banner
Write-Verbose "---------------------------------------------------------------------------------" -Verbose:$verbose
Write-Verbose "END ==== $scriptBaseName ====" -Verbose:$verbose
# =======

#get-date;

