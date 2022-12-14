# close-WinsAll.ps1

<# 
.SYNOPSIS
close-WinsAll.ps1 - Gracefully close (prompted saves) on all desktop windows, Control Panel Wins, and Explorer Wins
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
# 10:23 AM 3/23/2018 and we're ending up with open Outlook windows as well. Rerun it!
# 8:52 AM 3/20/2018 rewrote the shell win code: hybrided CP,Explr & IE (also still overrides the IE 'close all tabs?' prompt).
# 10:07 AM 3/13/2018 #25 IE close: first do all IE, windows and blow past the 'Close All' prompts (only IE turns up in ShellApp):
# 10:33 AM 9/29/2017 add a trailing beep to notify done
7:06 AM 6/1/2016 - fix missing ; at #30, char: 169
8:48 AM 5/27/2016 - initial version
.DESCRIPTION
close-WinsAll.ps1 - Gracefully close (prompted saves) on all desktop windows, Control Panel Wins, and Explorer Wins. Does leave apps that prompted to save open (obviously)
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Echos to console
.EXAMPLE
. close-WinsAll.ps1
.LINK
#>

# w proc echo, trailing control.exe close, and trailing explr close
write-host -foregroundcolor green "***v ISSUING CLOSE ON _ALL_ DESKTOP WINDOWS! v***" ;
(get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -ne "powershell" } )| % {"proc:$($_.Name)" ; $_.CloseMainWindow() ; } ;
<# express equiv: KILL everything but ps, dead
(get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -ne "powershell" } )| stop-process ; 
#>

# 10:07 AM 3/13/2018 #25 IE close: first do all IE, windows and blow past the 'Close All' prompts (only IE turns up in ShellApp):
$ShellApp = New-Object -ComObject Shell.Application ; 
$SWinNames="Control Panel","Windows Explorer","Windows Internet Explorer","Internet Explorer" ; 
write-host -foregroundcolor green "***v ISSUING CLOSE ON _ALL_ SHELL WINDOWS! v***" ;
foreach($SWinN in $SWinNames){
    if($TWins=(New-Object -comObject Shell.Application).Windows() |?{$_.Name -eq $SWinN}){
        switch -regex ($SWinN){
           "Control\sPanel" { $wTag="CP" }
           "Windows\sExplorer" { $wTag="Explr" }
           "((Windows\s)*)Internet\sExplorer" { $wTag="IE" }
        } ; 
        "`n==close $(($TWins|measure).Count) $($wTag) wins..." ; 
        $TWins| % {"$($wTag):$($_.LocationName)" ; $_.quit()} ;
    } ; 
} ;  # loop-E

# 10:40 AM 3/20/2018 FINE, still have IE sitting there on close all tabs! KILL IT!
if($ieZombs=(get-process | ? { $_.processname -eq "iexplore" }) ){
    write-host -foregroundcolor green "***v KILLING REMAINING IEXPLORE WINDOWS! v***" ;
    $ieZombs| % {"proc:$($_.Name)" ; $_.Kill() ; } ;
} ; 
# 10:23 AM 3/23/2018 and we're ending up with open Outlook windows as well. Rerun it!
if($othrZombs=(get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -ne "powershell" } ) ){ 
    write-host -foregroundcolor green "***v RE-CLOSING REMAINING  DESKTOP WINDOWS! v***" ;
    $othrZombs | % {"proc:$($_.Name)" ; $_.CloseMainWindow() ; } ;
} ; 


