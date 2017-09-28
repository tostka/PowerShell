# ping-RDP.ps1
<#
.SYNOPSIS
ping-RDP.ps1 - runs a dawdle loop polling for RDP port 3389 up on target server. Once avail launches an rdp (mstsc.exe /v:xxx) to the specified box.
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
10:35 AM 9/28/2017 been running these for years in my profile, decided to wrap them up into a simple postable combo - very handy when running server maint reboots across multiple servers at once, and you want to get back into a given box as soon as it's accessible. (don't have to watch, it'll beep & pop open the tsc logon prompt)
.INPUTS
Accepts piped input.
.OUTPUTS
Returns an object with uptime data to the pipeline.
.EXAMPLE
ping-RDP USEA-MAILEXP | select Computername,Uptime
#>

#*------v Function Test-Port() v------
function Test-Port {
    # attempt to open a port (telnet xxx.yyy.zzz nnn)
    # call: Test-Port $server $port
    PARAM([parameter(Mandatory=$true)] [alias("s")] [string]$Server, [parameter(Mandatory=$true)][alias("p")]
    [int]$port
    ) ; 
    $ErrorActionPreference = “SilentlyContinue” ; 
    $socket = new-object Net.Sockets.TcpClient ; 
    $socket.Connect($Server, $port) ; 
    if ($socket.Connected) {
        write-verbose -verbose:$true  "Successful connection to $($Server):$($port)"
        $socket.Close()
        return $True;
    } else {
        write-verbose -verbose:$true  "Failed to connect to $($Server):$($port)"
         return $False;
    } # if-block end
    $socket = $null
}#*------^ END Function Test-Port() ^------

#*======v SUB MAIN v======
$tsrvr=read-host -prompt "Server to be pinged until RDP available" ;
Do {write-host "." -NoNewLine;Start-Sleep -m (1000 * 5)} Until ((test-port $tsrvr 3389)) ;
 write-host "`a" ; # beep
 write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Launching rdp to $($tsrvr)..." ;
mstsc.exe /v:$($tsrvr) ;
#*======^ END SUBM MAIN ^======
