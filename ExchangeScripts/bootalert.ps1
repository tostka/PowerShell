#bootalert.ps1

<#
    .SYNOPSIS
bootalert.ps1 - Send email notification post-reboot

    .NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Additional Credits: [REFERENCE]
Website:	[URL]
Twitter:	[URL]

Change Log
* 12:06 PM 12/2/2013 finish up PS port and test
* 5:14 PM 7/18/2013 initial version

    .DESCRIPTION
Send email notification post-reboot, non-blat version of bootalert.cmd. 

    .PARAMETER  <Parameter-Name>
<parameter comment>

    .INPUTS
None. Does not accepted piped input.

    .OUTPUTS
None. Returns no objects or output.

    .EXAMPLE
.\bootalert.ps1
[use an .EXAMPLE keyword per syntax sample]

#>

<# :::================v WRAPPER v================
::: bootalert.cmd
:::	===remote WIN2K8 system===
::: on 2008R2 new /rp params For Task Scheduler 2.0 tasks:
::: /ru "NT AUTHORITY\LOCALSERVICE"
::: /ru "NT AUTHORITY\NETWORKSERVICE"

::: 1) schtasks /create /s USEA-fdhubcas2 /tn "bootalert" /tr e:\scripts\bootalert.ps1 /sc onstart /ru "NT AUTHORITY\NETWORKSERVICE" 
::: 2) launch it to verify function:
:::		schtasks /run /tn bootalert /s USEA-fdhubcas2
::: 3) Verify that you receive an email.

@ echo off

::: 7:50 AM 5/8/2013 need to do some path detection since we're running this on diff dirs
::: PERFORM STANDARD TESTS
::: DERIVE SCRIPTPATH DRIVE LETTER
SET SCRIPTPATH=%~d0
::: DERIVE & ADD SCRIPTPATH PATH
SET SCRIPTPATH=%SCRIPTPATH%%~p0
ECHO SCRIPTPATH=%SCRIPTPATH%
::: echos as 'c:\Scripts\'
set PSExec=@%WINDIR%\system32\windowspowershell\v1.0\powershell.exe
::: DEFAULT install locations, correct to actual installs loc
set Ex10BinPath=%ProgramFiles%\Microsoft\Exchange Server\V14\bin\
set Ex7BinPath=%ProgramFiles%\Microsoft\Exchange Server\bin\
::: TargetScript
set targPS1=%SCRIPTPATH%bootalert.ps1
ECHO targPS1=%targPS1%
:RERUNIT

::: ***EMS10 Powershell call***
::IF EXIST %targPS1% (@%PSExec% -noexit -command "& {. '%Ex10BinPath%RemoteExchange.ps1'; Connect-ExchangeServer -auto ; %targPS1% }") ELSE ((ECHO MISSING %targPS1%) & (PAUSE))

::: ***EMS7 Powershell call*** 
::IF EXIST %targPS1% (@%PSExec% -PSConsoleFile "%Ex7BinPath%\exshell.psc1" -noexit -command ". '%Ex7BinPath%Exchange.ps1' ; %targPS1%") ELSE ((ECHO MISSING %targPS1%) & (PAUSE))

::: ***STOCK POWERSHELL call (or EMS7, if call the snapin in the target ps1)***
::IF EXIST %targPS1% (@%PSExec% -command "& {%targPS1% }") ELSE ((ECHO MISSING %targPS1%) & (PAUSE))
IF EXIST %targPS1% (@%PSExec% -command "& {%targPS1% }") ELSE ((ECHO MISSING %targPS1%) & (PAUSE))

::: ***STOCK POWERSHELL ON WIN7 UNDER ExecutionPolicy RESTRICTIONS***
::: This completely bypasses ExecutionPolicy restrictions 
:::IF EXIST %targPS1% (@%PSExec% -ExecutionPolicy Bypass -NoLogo -NoProfile -File %targPS1%) ELSE ((ECHO MISSING %targPS1%) & (PAUSE))
::: works to run script from java

:::pause if needed to sustain results, or for bootalert etc
::pause

::: ADDED TO MOVE PAST CRASHES
::GOTO :RERUNIT 
:::================^ END WRAPPER ^================#>

$SMTPserver = "LYNMS650"
$SMTPport = 8111
# seconds to wait for startup to complete, before sending notice
$SendDelay=60
$bDebug = $FALSE
#$TRUE

$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\"
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))
$ComputerName = ($env:COMPUTERNAME)
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName)  
$TimeStampNow = get-date -uformat "%Y%m%d-%H%M" 
if ($bDebug) {
  write-host -foregroundcolor yellow "*** DEBUG MODE ***"
  $outransfile=$ScriptDir + $ScriptNameNoExt + "-" + "trans.txt" 
} else {
  $outransfile=$ScriptDir + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.txt" 
}
#stop-transcript -ErrorAction SilentlyContinue
# stop transcript,trap any error & eat complaint
# note, this will suppress all errors coming out of the transcript commands - even one's you WANT to see:
Trap {Continue} Stop-Transcript | Out-Null
start-transcript -path $outransfile

$strQuote = [char]34
#$SMTPfrom = "<" + ($ScriptBaseName.replace(".","-")) + "@unisys.com>" 
$SMTPfrom = "<" + ($ScriptBaseName.replace(".","-")) + "@toro.com>" 
$SMTPto = "Todd Kadrie <todd.kadrie@toro.com>" 

#$SMTPsubject = ("Script:" + $ScriptBaseName + "-" + $ComputerName + " has been rebooted " + (get-date).toshorttimestring())
$SMTPsubject = ($ComputerName + " has been rebooted " + (get-date).toshorttimestring())
$SMTPbody = ("Script:" + $ScriptBaseName + ":" + $ComputerName + " has been rebooted " + (get-date).toshorttimestring())
$SMTPfrom=$strQuote + $SMTPfrom + $strQuote
$SMTPto=$strQuote + $SMTPto + $strQuote
$SMTPsubject=$strQuote + $SMTPsubject + $strQuote
$SMTPbody=$strQuote + $SMTPbody + $strQuote 
# 9:35 AM 7/3/2014 adding port spec: -Port <Int32>
$ExecCommand = "send-mailmessage -from $SMTPfrom -to $SMTPto -subject $SMTPsubject -body $SMTPbody -smtpServer $SMTPserver -Port $SMTPPort" 
write-host "ExecCommand: " $ExecCommand
write-host ("waiting " + $SendDelay + " seconds to send")
#Start-Sleep -m (1000 * $SendDelay)
# 12:11 PM 12/2/2013 add a dot-crawl

write-host "[" 
for ($a = 1; $a -le $SendDelay; $a++) {
  write-host "." -NoNewLine;
  Start-Sleep -seconds 1
}
# assert a wrap on final dot
write-host -foregroundcolor red "]" 



Invoke-Expression $ExecCommand -verbose
Clear-Variable ExecCommand -ErrorAction SilentlyContinue

stop-transcript
