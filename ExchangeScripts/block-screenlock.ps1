# block-screenlock.ps1

<# 
.SYNOPSIS
block-screenlock.ps1 - blocks screensaver trigger by using wiggling mouse or using vbs sendkeys to type silent F15
.NOTES
To detect and/or kill: get-process | ?{$_.mainwindowtitle -match "block-screenlock-monitor"} | stop-process ;
Setup as a Sched Task, on-logon, single userid (won't work from 'on any user'). 
Task Program/Script: 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Add argument (optional): for Stock wrapper-less Action: 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
-noprofile -windowstyle hidden -noexit -command "& {c:\scripts\block-screenlock.ps1 }"
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Add argument (optional): for visible Debugging wrapper-less Action: 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
-noprofile -noexit -command "& {c:\scripts\block-screenlock.ps1 -showDebug}"
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Distribution: Config Task, Export to XML.
Distribution command (Ex servers)
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$files = (gci -path "\\$env:COMPUTERNAME\c$\scripts\block-screenlock.ps1*" | ?{$_.Name -match "(?i:(.*\.(PS1|CMD|XML)))$" }) ; get-exchangeserver | ?{$_.IsE14OrLater} | foreach { write-host $_ ; copy $files -Destination \\$_\c$\scripts\ -whatif ; } ; get-date ;
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Lync servers: 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$files = (gci -path "\\$env:COMPUTERNAME\c$\scripts\block-screenlock.ps1*" | ?{$_.Name -match "(?i:(.*\.(PS1|CMD|XML)))$" }) ; $L13ProdALL="LYNMS6200;LYNMS6201;LYNMS6202;LYNMS9209;LYNMS9201-1;LYNMS9201-2".split(";") ;$L13ProdALL | foreach { write-host $_ ; copy $files -Destination \\$_\c$\scripts\ -whatif ; } ; get-date ;
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

* # import on older systems via schtasks.exe; (Since my XML definition does not include the credentials for the task, I need to specify them at import time):
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
schtasks /create /tn "\block-screenlock.ps1" /xml "c:\scripts\block-screenlock.ps1 (kadrtiss Screenlock Block).xml" /ru toro\kadritss /rp * /s lynms651 ; 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
* Bulk remote Win2008+ copy & import exported .xml into remote boxes: (note, -whatif only blocks copy, if a matching .xml is at dest, the import will occur even in -whatif!)
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$srcTask="c:\scripts\block-screenlock.ps1 (kadrtiss Screenlock Block).xml" ;$file =  get-childitem -path $srcTask ;$L13ProdALL="LYNMS6200;LYNMS6201;LYNMS6202;LYNMS9209;LYNMS9201-1;LYNMS9201-2".split(";") ;$L13ProdALL | foreach {  write-host -fore yell "importing SchedTask $($file) to $_" ;  copy-item -path $file –destination \\$_\c$\scripts\ -whatif ;   if(test-path "\\$_\c$\scripts\$($file.Name)") {    Register-ScheduledTask -CimSession $_ -Xml (get-content "$file" | out-string) -TaskName $((split-path $srcTask -leaf).replace(".xml","")) -User toro\kadritss –Force ;    } else { write-warning "$((get-date).ToString("HH:mm:ss")):Missing src xml at far end!" } ;} ;
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
# adapted gtom anon post at http://dmitrysotnikov.wordpress.com/2009/06/29/prevent-desktop-lock-or-screensaver-with-powershell/
# launch from cmd as: powershell.exe -windowstyle hidden -file C:\usr\local\bin\block-screenlock-ps1.cmd
# powershell.exe -windowstyle hidden -file C:\scripts\block-screenlock-ps1.cmd
# powershell.exe -file C:\scripts\block-screenlock-ps1.cmd
Change Log
9:40 AM 7/27/2016 updated older tasks import cmdline
9:01 AM 7/27/2016 added help, distribution & task docs
7:36 AM 7/27/2016 doesn't seem to be working on lynms650 (as of a few days ago), add a dot crawl, shift it to schedtask launch.
10:02 AM 6/18/2015 added -ShowDebug and echo's to tell it's working 
11:44 AM 12/27/2013 trying out a hybrid version
11:30 AM 12/27/2013 flipped to F15 keystroke version ;was getting errors grabbing cursor pos
10:31 AM 10/17/2013 initial version
.DESCRIPTION
blocks screensaver trigger by either using Mouse-wiggle [-Target MOUSE] or using vbs sendkeys to type silent F15 [-Target KEY]
.PARAMETER  Target
block via KEYpress or MOUSEcursor [KEY | MOUSE]
.PARAMETER  ShowDebug
Display debugging commands and console dot crawl

.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Outputs a report to console and logfile, and emails the report. 
.EXAMPLE
.\block-screenlock.ps1 
Load with default [Mouse]
.EXAMPLE
.\block-screenlock.ps1 -Target Mouse
Load with Mouse-wiggle blocking
.EXAMPLE
.\block-screenlock.ps1 -Target KEY
Load with Key-press ([F15]) blocking (F15 is generally unbound in most applications).
.LINK
#>

#------------- Declaration and Initialization of Global Variables -------------	


PARAM(
  [alias("t")]
  [string] $Target="MOUSE"
  ,[switch] $showDebug
) ;

## immed after param, setup to stop executing the script on the first error
trap { break; }

# debugging flag
#$bDebug = $TRUE
#$bDebug = $FALSE
if($ShowDebug){ $bDebug = $true} ;

If ($bDebug -eq $TRUE) {write-host "*** DEBUGGING MODE ***"}

<#
# assign specs per params: 
$SiteName=$SiteName.ToUpper()
$ExTargVers=$ExTargVers.ToUpper()
# validate the inputs
if (!($SiteName -match "^(US|EU|AU)$")) {
  write-warning ("INVALID SiteName SPECIFIED: " + $SiteName + ", EXITING...")
  exit 1
} # if-block end
if (!($ExTargVers -match "^(2007|2010)$")) {
  write-warning ("INVALID ExTargVers SPECIFIED: " + $ExTargVers + ", EXITING...")
  exit 1
} # if-block end
#>

# single/double quote constants
$sQuot = [char]34 ; 
$sQuotS = [char]39 ;

$ComputerName = ($env:COMPUTERNAME)

# derive paths\filenames relative to script location (for output logs)
$ScriptDir = (Split-Path -parent $MyInvocation.MyCommand.Definition) + "\"
$ScriptBaseName = [system.io.path]::GetFilename($MyInvocation.InvocationName)
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName)  

If ($bDebug -eq $TRUE) {
  write-host ("`$ScriptDir: " + $ScriptDir)
  write-host ("`$ScriptNameNoExt: " + $ScriptNameNoExt)
  write-host ("`$ScriptBaseName: " + $ScriptBaseName)
} # if-block end

$TimeStampNow = get-date -uformat "%Y%m%d-%H%M" 

write-host " "
<#
# confirm log dir exists
$testPath= $ScriptDir + "logs\" 
if (!(test-path $testPath)) {
  write-host "Creating " $testPath
  New-Item $testPath -type directory    
} # if-block end

# setup mailing information
# build From address as a variant of the script name
$SMTPFrom = (($ScriptBaseName.replace(".","-")) + "@domain.com") ; 
$SMTPServer = "mail.domain.com" ; 
$SmtpToAdmin="todd.kadrie@domain.com" ; 
$SMTPTo=$SmtpToAdmin ; 

# setup msg body as a hash
$SmtpBody = @() ; 
#>

#*================v FUNCTION LISTINGS v================

#---------------------------v Function Cleanup v---------------------------
Function Cleanup {
  # clear all objects and exit
  
  exit
} #*----------------^ END Function Cleanup ^----------------

#---------------------------v Function BlockScrKeys v---------------------------
Function BlockScrKeys ($Target) {
  # 7:45 AM 7/27/2016 transplanted dbg & dawdle dotcrawl from BlockScrMouse
  If ($bDebug) {"pressing {F15} key via VBS: 5sec debugging (50s normally)..."  }
  $host.ui.RawUI.WindowTitle = "block-screenlock-monitor"
  $VbsShell = New-Object -com "Wscript.Shell"

  for(;;) {
    # 7:46 AM 7/27/2016 copied from Mouse
    If ($bDebug) { Start-Sleep -Seconds 5 ; write-host "." -NoNewLine ; } 
    else { Start-Sleep -Seconds 50} 
    $VbsShell.sendkeys("{F15}")
  } # for-loop end

  # then to target and kill the above:
  #get-process | ?{$_.mainwindowtitle -match "block-screenlock-monitor"} | kill
} #*----------------^ END Function BlockScrKeys ^----------------

#---------------------------v Function BlockScrMouse v---------------------------
Function BlockScrMouse ($Target){
  
  <# 
  .SYNOPSIS
  BlockScrMouse - Wiggles the mouse to block screensaver
  .NOTES
   Updated By: Todd Kadrie, smoothed out some real problems with mouse handling, added debugging output etc.
   Written By: gtom anon post
  Website:	http://dmitrysotnikov.wordpress.com/2009/06/29/prevent-desktop-lock-or-screensaver-with-powershell/
  Change Log
  * 8:12 AM 7/27/2016 tweaked mouse position code - update pos each loop to stop it from jumping to starting point. Added help to func
  * 10:02 AM 6/18/2015 added -ShowDebug and echo's to tell it's working 
  .DESCRIPTION
  BlockScrMouse - Wiggles the mouse to block screensaver
  .PARAMETER  Target
  ParaHelpTxt
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  .LINK
  #>
  If ($bDebug -eq $TRUE) {
    write-host " "
  } # if-block end

  $host.ui.RawUI.WindowTitle = "block-screenlock-monitor"
  [System.Reflection.Assembly]::LoadWithPartialName("system.windows.forms")
  # grab cursor object
  $Cursor = [system.windows.forms.cursor]::Clip
  <# 7:53 AM 7/27/2016 issue only pulling pos once: it reverts to that position every 50s, need to pull pos every loop!
  $Position = [system.windows.forms.cursor]::Position
  #>
  If ($bDebug) {
    "wiggling: 5sec debugging (50s normally)..." ; 
    "`$Cursor:$Cursor" ; 
    "`$Position:$Position" ; 
  } ; 
  
  for(;;) {
    # 7:42 AM 7/27/2016 added dot crawl
    If ($bDebug) { Start-Sleep -Seconds 5 ; write-host "." -NoNewLine ; } 
    else { Start-Sleep -Seconds 50}  ; 
    # 7:56 AM 7/27/2016: shifting to poll cur cursor pos each pass, avoid curosr jumps!
    $Position = [system.windows.forms.cursor]::Position ; 
    # wiggle it 1pxl up and then back down
    [system.windows.forms.cursor]::Position = New-Object system.drawing.point($Position.x, ($Position.y + 1)) ; 
    If ($bDebug) { write-host "$(([system.windows.forms.cursor]::Position).y)," -NoNewLine ; } ; 
    # add slight delay
    Start-Sleep -m 5 ; 
    [system.windows.forms.cursor]::Position = New-Object system.drawing.point($Position.x, $Position.y) ; 
    If ($bDebug) { write-host "$(([system.windows.forms.cursor]::Position).y)," -NoNewLine ; } ; 
  } # for-loop end
  
  # then to target and kill the above:
  #get-process | ?{$_.mainwindowtitle -match "block-screenlock-monitor"} | kill
} #*----------------^ END Function BlockScrMouse ^----------------


#*================^ END FUNCTION LISTINGS ^================

#*================v SCRIPT BODY v================
#--------------------------- Invocation of SUB_MAIN ---------------------------

# 11:56 AM 12/27/2013 $Target="KEY | MOUSE
if ($Target.ToUpper() -eq "KEY") {
  # block by hitting F15 every 50secs
  BlockScrKeys
} elseif($Target.ToUpper() -eq "MOUSE") {
  # block by wiggling cursor 1pxl every 50secs
  BlockScrMouse
}
# cleanup and exit
Cleanup
  
#*================^ END SCRIPT BODY ^================
 
