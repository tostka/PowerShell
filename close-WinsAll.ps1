# close-WinsAll.ps1

<# 
.SYNOPSIS
close-WinsAll.ps1 - Gracefully close (prompted saves) on all desktop windows, Control Panel Wins, and Explorer Wins
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
7:06 AM 6/1/2016 - fix missing ; at #30, char: 169
8:48 AM 5/27/2016 - initial version
.DESCRIPTION
close-WinsAll.ps1 - Gracefully close (prompted saves) on all desktop windows, Control Panel Wins, and Explorer Wins. Does leave apps that prompted to save open (obviously)
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
. close-WinsAll.ps1
.LINK
#>

# w proc echo, trailing control.exe close, and trailing explr close
write-host -foregroundcolor green "***v ISSUING CLOSE ON _ALL_ DESKTOP WINDOWS! v***" ;
(get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -ne "powershell" } )| % {"proc:$($_.Name)" ; $_.CloseMainWindow() ; } ;
<# express equiv: KILL everything but ps dead
(get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -ne "powershell" } )| stop-process ; 
#>

"close CP wins..." ; 
(New-Object -comObject Shell.Application).Windows() | where-object {$_.LocationName -eq "Control Panel"} | foreach-object {"CP:$($_.LocationName)" ; $_.quit()} ;

"close Explr wins..." ; (New-Object -comObject Shell.Application).Windows() |?{$_.Name -eq 'Windows Explorer'} | % {"explr:$($_.LocationName)" ; $_.quit()} ;
write-host -foregroundcolor green "***^ ISSUED CLOSE ON _ALL_ DESKTOP WINDOWS! ^***" ;

