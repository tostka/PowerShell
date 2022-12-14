# check-PingSlowLoop.ps1

<# 
    .SYNOPSIS
    check-PingSlowLoop.ps1 - Ping monitor on a slow cycle with dotcrawl-style color-coded ping metrics
    .NOTES
    Written By: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka
    Change Log
    # 8:21 AM 7/25/2017 general cleanup, purged unused bits
    # 9:50 AM 1/11/2017 added win title store & restore on exit; 
        added F12 cancel to the loopm switch Delay default 2=>5; 
        color coding console output; 
        added a distinctive non-default title string (avoid killing it when closing other powershell blocks by title)
    # 6:59 AM 6/7/2016: pasted from routinely used 1-liner
    .DESCRIPTION
    check-PingSlowLoop.ps1 - Ping monitor on a 5-sec refresh cycle with dotcrawl-style color-coded ping metrics
    [array]$rpThresh spec's color-coding ms ping levels (yellow v red): Defaults to yellow @80ms, red @120ms, -lt 80 => green
    Provides a 'crawl' of values over time, to evaluate slow connection status (VPN, flakey high-latency etc):
    Sample Output: 
        Pinging samplecomputer:(F12=Cancel)[,9,0,3,1,1,1,0,0,0,1,2,1,0,0,0,0,2,0,1,0,0,0,0,1,0
        ,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,1,0,0,0,0,0,0,0,1,1,2,0,0,0,1,0,0,0,0,0,0,
    I routinely run it from home when working over a VPN.            
    .PARAMETER  Computer
    Computer to be pinged (Fqdn,NBName,IP)[-computer hostname]
    .PARAMETER  Delay
    Delay time in seconds[-delay nn]
    .PARAMETER showDebug
    Debugging Flag [-showDebug]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Outputs to console a comma-separated stream of the ping results over time. 
    .EXAMPLE
    check-PingSlowLoop.ps1 -Computer ServerX
    Ping ServerX second delay between pings
    .LINK
#>

Param(
    [Parameter(Position=0,Mandatory=$true,HelpMessage='Computer to be pinged (Fqdn,NBName,IP)[-computer hostname]',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [Alias('__ServerName', 'Server', 'ComputerName', 'Name')]
    [string[]]$Computer,
    [Parameter(HelpMessage='Delay time in seconds[-delay nn]')]
    [int]$Delay=5,
    [Parameter(HelpMessage='Debugging Flag [-showDebug]')]
    [switch] $showDebug
) ;

# 7:24 AM 1/11/2017 add response thresholds
$rpThresh=@{"Yellow"=120 ; "Green"=80 ; } ; 
#write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Thresholds in use: $($rpThresh| out-string)" ;

#*======v FUNCTIONS  v======
Function Test-AdminAcct {
    if(!$adminusername){
        $adminusername = 'MYDOMAIN\user' # If you have a separate user account for admin type tasks, provide the DOMAIN\USERNAME here
    } ;
    # 7:02 AM 1/11/2017 leverage the vari
    if("$($env:USERDOMAIN)\$($env:USERNAME)" -eq $adminusername){ write-verbose -verbose:$true "acct:$($adminusername)" ; write-output $true}  
    else { write-host "NOT MyAccount" ;write-output $false} ; 
} #*------^ END Function Test-AdminAcct ^------ ; 
#*======^ END FUNCTIONS ^======

#*======v SUB MAIN v======
# defaulting, but with Mandetory, it will autoprompt - more conven than defaulting to ggl. 
$iSpt=@{computername="192.168.0.1";
    count=1;
    erroraction=0;
} ;

# set custom title, cache orig for postrestore and append admin name (if set) to the title
$sWinTitleO=$host.ui.RawUI.WindowTitle ; 
$adminusername = 'MYDOMAIN\user' # If you have a separate user account for admin type tasks, provide the DOMAIN\USERNAME here
$sWinTitle="PS-mon : $(Split-Path -Leaf ((&{$myInvocation}).ScriptName)) :"
if(Test-AdminAcct){ $sWinTitle+=" $($adminusername)"} ; 
$host.ui.RawUI.WindowTitle = $sWinTitle ; 

if($Computer){$iSpt.computername=$Computer} 
else { write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Defaulting to $($iSpt.computername)" ;}
write-host -fore green -nonewline "Pinging $($iSpt.computername):(F12=Cancel)[";
# running F12 cancel loop
$continue=$true ; 
while($continue){
    if ([console]::KeyAvailable){
        write-host "`nCancelled with F12" ; 
        $x = [System.Console]::ReadKey() ;
        switch ( $x.key){
            F12 { $continue = $false ; } ;
        } ; 
    } ;
    $rp=(Test-Connection @iSpt).responsetime ; 
    if($rp -lt $rpThresh.Green){
        $whColor="green" ; 
    }elseif($rp -lt $rpThresh.Yellow){
        $whColor="yellow" ; 
    }elseif($rp -gt $rpThresh.Yellow){
        $whColor="red" ; 
    } ; 
    write-host -nonewline -foregroundcolor $whColor ",$($rp)" ;
    sleep $($Delay);
} ; 
# 9:48 AM 1/11/2017 on exit restore title
$host.ui.RawUI.WindowTitle = $sWinTitleO ; 
 
 #*======^ END SUB MAIN ^======
 
