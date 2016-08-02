# check-PingSlowLoop.ps1

  <# 
  .SYNOPSIS
  check-PingSlowLoop.ps1 - Ping monitor on a slow cycle
  .NOTES
  Written By: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka
  Change Log
  6:59 AM 6/7/2016: initial version, paste from 1-liner
  .DESCRIPTION
  check-PingSlowLoop.ps1 - Ping monitor on a slow cycle
  .PARAMETER  Computer
  Computer Fqdn,NBName,IP to be pinged
  .PARAMETER  Delay
  Seconds between pings
  .PARAMETER showDebug
  Display debug messages
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  check-PingSlowLoop.ps1 -Computer lynms650 -Delay 5
  Ping lynms650 on a 5 second delay between pings
  .LINK
#>


Param(
    [Parameter(Mandatory=$true,HelpMessage="Computer to be pinged [Fqdn,NBName,IP]")]
    [string]$Computer
    ,[Parameter(HelpMessage='Delay time in seconds[nn]')]
    [int]$Delay=2
    ,[Parameter(HelpMessage='Debugging Flag [$switch]')]
    [switch] $showDebug
) ;

$iSpt=@{computername="8.8.8.8";
    count=1;
    erroraction=0;
} ;

if($Computer){$iSpt.computername=$Computer} 
else { write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Defaulting to $($iSpt.computername)" ;}
 write-host -fore green -nonewline "Pinging $($iSpt.computername):[";
for(;;){
    write-host -nonewline ",$((Test-Connection @iSpt).responsetime)" ;
    sleep $($Delay);
 } ;
 
