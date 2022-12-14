# tsk-incl-Desktop.ps1.ps1

# Desktop Function Includes

#*======v NON-ADMIN: C:\Users\Account\Documents\WindowsPowerShell\profile.ps1 v======
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\Account\Documents\WindowsPowerShell\profile.ps1
# NON-ADMIN acct $profile.CurrentUserCurrentHost loc
#C:\Users\kadriets\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1

# ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\Accounts\Documents\WindowsPowerShell\profile.ps1
# notepad2 $profile.CurrentUserAllHosts ;
#*======

#*------V Comment-based Help (leave blank line below) V------ 

## 
#     .SYNOPSIS
# NON-ADMIN: C:\Users\Account\Documents\WindowsPowerShell\profile.ps1 - 
# My primary non-admin Profile file
# 
#     .NOTES
# Written By: Todd Kadrie
# Website:	http://tinstoys.blogspot.com
# Twitter:	http://twitter.com/tostka
# 
# Additional Credits: [REFERENCE]
# Website:	[URL]
# Twitter:	[URL]
# 
# Change Log.
# 12:38 PM 5/24/2016 Sub Main: adding Git support & Posh-Git
# 9:01 AM 3/16/2016 get-password(): debugged back to function
# 8:33 AM 3/16/2016 get-password(): added optional -times/-t param, specifying number of passpharses to produce
# 8:19 AM 3/16/2016 get-password(): added optional -reportLength/-rl 'length' reporting to the output string
# 9:30 AM 2/22/2016: ported in get-password.ps1 from uwps copy
# 7:28 AM 12/17/2015 added alias ctw=>close-taskwindows.ps1
# 7:20 AM 11/23/2015 restart-ff(): test $PMpub before invoking
# 8:37 AM 10/30/2015 trimmed length of --/==
# 7:10 AM 5/18/2015 added restart-ff
# 7:30 AM 4/30/2015 added restart-ahk & replaced write-output (echos) with "write-verbose -verbose:$true"
# 9:03 AM 12/30/2014 sub'd ulb for $binpath, and uhd for $InclPath, added $TextEd vari for notepad2/notepad
# 9:46 AM 12/23/2014 - split out Toys to include file
# 8:13 AM 12/23/2014 constructed SERVER-SIDE CORE subset of functions appropr for profile on server desktops

<# creation cmds
# create the CurrentUserAllHosts file:
new-item -path $PROFILE.CurrentUserAllHosts -ItemType file -Force ; notepad2 $PROFILE.CurrentUserAllHosts ;
# create the CurrentUserCurrentHost (default) file, type:
new-item -path $profile -itemtype file -force; notepad2 $PROFILE.CurrentUserCurrentHost ;
#>

write-host -foregroundcolor gray "$((get-date).ToString("HH:mm:ss")):====== EXECUTING: $(Split-Path -Leaf ((&{$myInvocation}).ScriptName) ) ====== " ; 
#           =========
#*======v *** DESKTOP FUNCTIONS *** v======
<# 11:48 AM 10/28/2014 rem-out, it's a PSCX component, non-portable to servers
#*------v Function Test-AdminLocal v------
# Often need to check for admin privilege (localAdmin membership)
# from: https://github.com/neilpa/dotfiles/blob/master/powershell/profile.ps1
function Test-AdminLocal { Test-UserGroupMembership Administrators }
 #*------^ END Function Test-AdminLocal ^------
#>
#*------v Function Out-Excel v------
# http://blogs.technet.com/b/heyscriptingguy/archive/2014/01/10/powershell-and-excel-fast-safe-and-reliable.aspx
# Simple func() to deliver Excel as a out-gridview alternative.
# vers: 1/10/2014
function Out-Excel {

  PARAM($Path = "$env:temp\$(Get-Date -Format yyyyMMddHHmmss).csv")

  $input | Export-CSV -Path $Path -UseCulture -Encoding UTF8 -NoTypeInformation

  Invoke-Item -Path $Path

}#*------^ END Function Out-Excel ^------

#*------v Function Out-Excel-Events v------
# http://blogs.technet.com/b/heyscriptingguy/archive/2014/01/10/powershell-and-excel-fast-safe-and-reliable.aspx
# Simple func() to deliver Excel as a out-gridview alternative, this variant massages array ReplacementStrings with a comma-delimited string. 
# vers: 1/10/2014
function Out-Excel-Events {
    PARAM($Path = "$env:temp\$(Get-Date -Format yyyyMMddHHmmss).csv")
    $input | Select -Property * |
    ForEach-Object {
       $_.ReplacementStrings = $_.ReplacementStrings -join ','
       $_.Data = $_.Data -join ','
       $_
    } | Export-CSV -Path $Path -UseCulture -Encoding UTF8 -NoTypeInformation
    Invoke-Item -Path $Path
}#*------^ END Function Out-Excel-Events ^------

#*------v Function start-ItunesPlaylist v------
Function start-ItunesPlaylist {
  <#
  .SYNOPSIS
      Plays an iTunes playlist.
  .DESCRIPTION
      Opens the Apple iTunes application and starts playing the given iTunes playlist.
  .NOTES
  Author: Frank Peter (http://www.out-web.net/?p=1390)
  .PARAMETER  Source
      Identifies the name of the source.(["Library"|"Internet Radio"]
  .PARAMETER  Playlist
      Identifies the name of the playlist
  .PARAMETER  Shuffle
      Turns shuffle on (else don't care).
  .EXAMPLE
     C:\PS> .\Start-PlayList.ps1 -Source 'Library' -Playlist 'Party'
  .EXAMPLE
     C:\PS> .\Start-PlayList.ps1 -source 'Library' -Playlist "classical-streams"
  .INPUTS
     None
  .OUTPUTS
     None
  #>
  [CmdletBinding()]
  param (
      [Parameter(Mandatory=$true)]
      $Source,
      [Parameter(Mandatory=$true)]
      $Playlist,
      [Switch]$Shuffle
  ) ;
  try {
      $iTunes = New-Object -ComObject iTunes.Application
  } catch {
      Write-Error 'Download and install Apple iTunes'
      return
  } ; 
  <# source options (interegated to get)
  Name
  ----
  Library
  Internet Radio
  #>
  $src = $iTunes.Sources | Where-Object {$_.Name -eq $Source} ; 
  if (!$src) {
      Write-Error "Unknown source - $Source" ; 
      return ; 
  } ;  # if-E
  $ply = $src.Playlists | Where-Object {$_.Name -eq $Playlist} ; 
  if (!$ply) {
      Write-Error "Unknown playlist - $Playlist" ; 
      return ; 
  } # if-E
  if ($Shuffle) {
      if (!$ply.Shuffle) {
          $ply.Shuffle = $true ; 
      } # if-E
  } # if-E
  $ply.PlayFirstTrack() ; 
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$iTunes) > $null ; 
  [GC]::Collect() ; 
} #*------^ END Function start-ItunesPlaylist ^------

#*------v Function Speak-words v------
function Speak-words  {   

<# 
.SYNOPSIS
speak-words - Text2Speech specified words
.NOTES
Written By: Karl Prosser
Website:	http://poshcode.org/835
Change Log
* 2:02 PM 4/9/2015 - added to profile
.PARAMETER  words
Words or phrases to be spoken
.PARAMETER  pause
switch indicating whether to hold execution during speaking 
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
speak-words "here we go now"  ;
.EXAMPLE
speak-words "$([datetime]::now)" ;
Speak current date and time
.EXAMPLE
get-fortune | speak-words ; 
Speak output of get-fortune
.LINK
http://poshcode.org/835
*------^ END Comment-based Help  ^------ #>
<#
  Param(
      [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Specify text to speak")]
      [ValidateNotNullOrEmpty()]
      [string]$words
    ,
      [Parameter(Position=1,Mandatory=$False,HelpMessage="Specify to wait for text to finish speaking")]
      [bool]$pause = $true
  ) # PARAM BLOCK END
  switch pause to a switch #>
  
  Param(
      [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Specify text to speak")]
      [ValidateNotNullOrEmpty()]
      [string]$words
    ,
      [Parameter(Position=1,Mandatory=$False,HelpMessage="Specify to wait for text to finish speaking")]
      [switch]$pause = $true
  ) # PARAM BLOCK END
  # default to no-pause, unless specified
  $flag = 1 ; if ($pause) {$flag = 2}  ; 
  $voice = new-Object -com SAPI.spvoice ; 
  $voice.speak($words, [int] $flag) # 2 means wait until speaking is finished to continue
 
} #*======^ END Function Speak-words ^====== ;
# 10:29 AM 2/19/2016
if(!(get-alias | ?{$_.name -like "speak"})) {Set-Alias -Name 'speak' -Value 'speak-words' ; } ;

#*------v Function Go v------
# 
# vers: 8:49 AM 3/27/2014 - tuned destinations
if( $GLOBAL:go_locations -eq $null ) {
    $GLOBAL:go_locations = @{};
    # 9:54 AM 3/27/2014 make it retain the as-added order
    #$GLOBAL:go_locations =[Ordered]@{};
} # if-block end

function Go ([string] $Location) {
    if ($go_locations.ContainsKey($Location)) {
        #write-output $go_locations[$Location] 
        Set-Location $go_locations[$Location];

        # 10:37 AM 3/27/2014this lists everything in the dir... NAH!
        #Get-ChildItem;
    } else {
        write-verbose -verbose:$true "---";
        write-verbose -verbose:$true "The following locations are defined:";
        write-verbose -verbose:$true $go_locations;
        # 9:06 AM 3/27/2014 sort the current contents - doesn't work
        #write-verbose -verbose:$true 
      #$go_locations = $go_locations | Sort-Object Name
      #$go_locations.GetEnumerator() | Sort-Object Name
      #write-host "boo"
      #$go_locations.GetEnumerator() | Sort-Object Name
    } # if-block end
    
    <#
        .SYNOPSIS
            Go - CD to common system locations
        .EXAMPLE
            PS C:\> Go <location keyword>
        .OUTPUTS
            [none]
    #>
} # function block end

$go_locations.Clear()


#$go_locations.Add("home", "~")


$go_locations.Add("dbx", "c:\dropbox")
$go_locations.Add("profile", $ENV:USERPROFILE)
$go_locations.Add("psp", (Split-Path $PROFILE) )
$go_locations.Add("win", $ENV:WINDIR)

#*------^ END Function Go ^------

# ====== v move-window functions by Sunnyone v ======
# This is qa set of functions that permit you to move-window a window, and even Move-WindowByWindowTitle
# AUTHOR: sunnyone
# URL: https://gist.github.com/sunnyone/7082155
# VERS: 1:36 PM 1/15/2014 commented & formatted by todd kadrie
# VERS: Created 2013-10-21
<# USAGE:
# prereq for the API calls used below
  Define-MoveWindow

  # below matches firefox, even though the WindowTitle doesn't seem to match the string below...
  Move-WindowByWindowTitle -ProcessName Firefox -WindowTitle "Win.ow" -X 100 -Y 200 -Width 640 -Height 480

#>
#*------v Function Define-MoveWindow v------
function Define-MoveWindow {
  $signature = @'
  [DllImport("user32.dll")]
  public static extern bool MoveWindow(
  IntPtr hWnd,
  int X,
  int Y,
  int nWidth,
  int nHeight,
  bool bRepaint);
'@
  Add-Type -MemberDefinition $signature -Name MoveWindowUtil -Namespace MoveWindowUtil
}#*------^ END Function Define-MoveWindow ^------
 
#*------v Function Move-Window v------ 
function Move-Window {
  PARAM ($Handle, [int]$X, [int]$Y, [int]$Width, [int]$Height);
   
  process {
    [void][MoveWindowUtil.MoveWindowUtil]::MoveWindow($Handle, $X, $Y, $Width, $Height, $true);
  } # proceses block end
} #*------^ END Function Move-Window ^------

#*------v Function Move-WindowByWindowTitle v------ 
function Move-WindowByWindowTitle {
  PARAM (
    [string]$ProcessName,
    [string]$WindowTitleRegex,
    [int]$X, [int]$Y, [int]$Width, [int]$Height)
  process {
    $procs = Get-Process -Name $ProcessName | Where-Object { $_.MainWindowTitle -match $WindowTitleRegex }
    foreach ($proc in $procs) {
    Move-Window -Handle $proc.MainWindowHandle -X $X -Y $Y -Width $Width -Height $Height
    } # for-loop end
  } # proceses block end
} #*------^ END Function Move-WindowByWindowTitle ^------
# ====== ^ move-window functions by Sunnyone ^ ======

#*------v Function get-lastsleep() v------
function get-lastsleep {
  # return the last 7 sleep events on the local pc
  # usage: get-lastsleep | ft -auto ;
  # vers: 7:17 AM 8/26/2014 corrected date fmt string
  # ver: 2:48 PM 8/25/2014 - fixed output to display day of week as well
  $nEvt=7;
  write-host -fore yellow ("=" * 10);
  "System" | %{
    write-host -fore yellow "`n===$_ LAST $nEvt Hibe/Sleep===" ;
    # reformat to include dow:
  # @{Name='Time';Expression={[string]::get-date $_.TimeGenerated -format 'ddd MM/dd/yyyy HH:mm tt'}}
    $sleeps=(get-eventlog -logname System -computername localhost -Source Microsoft-Windows-Kernel-Power -EntryType Information -newest $nEvt -message "*sleep*" | select @{Name='Time';Expression={get-date $_.TimeGenerated -format 'ddd MM/dd/yyyy h:mm tt'}},Message);
  } ;
  # return an object
  #return $sleeps;
  # or just dump to cons
  $sleeps| ft -auto;
} #*------^ END Function get-lastsleep() ^------

#*------v GOTO Functions  v------
#function gotoDevExScripts{set-location C:\usr\work\exch\scripts}
#function gotoDevLynScripts{set-location C:\usr\work\exch\scripts}
#function gotoProfile{set-location C:\Users\Accounts\Documents\WindowsPowerShell}
function gotoIncid{set-location c:\usr\work\incid}
# Navigation aliases
Set-Alias p Pop-Location
function ~ { Push-Location (Get-PSProvider FileSystem).Home }
function .. { Set-Location .. }
function ... { ..;.. }
function .... { ...;.. }

<#$AcctAdmn="TORO\AccountS"
$AcctUser="TORO\Account"
#>
# 7:24 AM 5/1/2014 if my non-admin logon and myBox (non-server)
if ($logonType -eq "MyNonAdmn") {
  # non-admin functions
  function gotoDownloads{set-location C:\usr\home\ftp}
  function gotoDbox{set-location c:\usr\home\dropbox}
  function gotoDboxDb{set-location c:\usr\home\dropbox\db}
} # if-block end
#*------^ END GOTO Functions ^------

#*------v OPEN Functions  v------
if ($logonType -eq "MyNonAdmn") {
  # non-admin functions
  function openInput{$sExc=$TextEd + " " + (join-path $binpath input.txt); Invoke-Expression $sExc;}
  function openTmpps1{$sExc=$TextEd + " C:\tmp\tmp.ps1"; Invoke-Expression $sExc;}
} # if-block end
#*------^ END OPEN Functions ^------

#*------v Function Function Set v------
# If an alias exists, remove it.
If (Test-Path ALIAS:set) { Remove-Item ALIAS:set } ; 
Function Set {

<# 
.SYNOPSIS
Set() - Emulate the DOS Set e-vari-handling cmd in PS
.NOTES
Written By: Bill Stewart
Website:	http://windowsitpro.com/powershell/powershell-how-emulating-cmdexes-set-command
Change Log
* 8:42 AM 4/10/2015 reformatted, added help
* Dec 12, 2011 posted
.DESCRIPTION
Note:  You can't use the Set function as part of a PowerShell expression, such as
(Set processor_level).GetType() 
But it has two advantages over Cmd.exe's Set command. First, it outputs DictionaryEntry objects, just like when you use the command...
Get-ChildItem ENV: 
Second, the Set function uses wildcard matching. For example, the command...
Set P 
...matches only a variable named P. Use Set P* to output all evari's beginning with P.
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
No formatting appears to have been put int, results are output to the pipeline.
.EXAMPLE
Set
To list all of the current $env: (equiv to gci $env:)
.LINK
http://windowsitpro.com/powershell/powershell-how-emulating-cmdexes-set-command
*------^ END Comment-based Help  ^------ #>
  If (-Not $ARGS) {
    Get-ChildItem ENV: | Sort-Object Name ; 
    Return ; 
  } ; 
  $myLine = $MYINVOCATION.Line ; 
  $myName = $MYINVOCATION.InvocationName ; 
  $myArgs = $myLine.Substring($myLine.IndexOf($myName) + $myName.Length + 1) ; 
  $equalPos = $myArgs.IndexOf("=") ; 
  # If the "=" character isn't found, output the variables.
  If ($equalPos -eq -1) {
    $result = Get-ChildItem ENV: | Where-Object { $_.Name -like "$myArgs" } |
      Sort-Object Name ; 
    If ($result) { $result } Else { Throw "Environment variable not found" } ; 
  } ElseIf ($equalPos -lt $myArgs.Length - 1) {
    # If the "=" character is found before the end of the string, set the variable.
    $varName = $myArgs.Substring(0, $equalPos) ; 
    $varData = $myArgs.Substring($equalPos + 1) ; 
    Set-Item ENV:$varName $varData ; 
  } Else {
    # If the "=" character is found at the end of the string, remove the variable.
    $varName = $myArgs.Substring(0, $equalPos) ; 
    If (Test-Path ENV:$varName) { Remove-Item ENV:$varName } ; 
  } # if-E
} #*------^ END Function Set ^------

#*------v Function restart-ahk v------
Function restart-ahk {

    <# 
    .SYNOPSIS
    restart-ahk - Close all autohotkey processes, and reslaunch any *.ahk.lnk files in the "$env:APPDATA\...Startup" folder
    .NOTES
    Written By: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka
    Change Log
    # 7:21 AM 6/12/2015 port in support for the -NoRestart switch from restart-itsm
    # 7:11 AM 5/18/2015 port back the restart-ahk enhancements
    # 7:26 AM 4/30/2015 port to profile funct and rename ahk-restart.ps1 to restart-ahk()
    # 8:47 AM 1/15/2015
    # 7:46 AM 12/12/2014

    .DESCRIPTION
    restart-ahk - Close all autohotkey processes, and reslaunch any *.ahk.lnk files in the "$env:APPDATA\...Startup" folder
    .PARAMETER NoRestart
    Parameter to suppress re-open[-NoRestart switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    .\ahk-restart
    .LINK
    *------^ END Comment-based Help  ^------ #>

    Param([Parameter(HelpMessage='NoRestart [$switch]')]
    [switch] $NoRestart) ; 

    $ScriptName=$myInvocation.ScriptName ;
    $StartFldr="$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"; #gci $StartFldr;
    #9:49 AM 5/16/2015 add quick launch fldr
    $QckFldr="$env:userprofile\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch" ;
    $TargAppName="autohotkey"
    #"Firefox or Palemoon"
    $TargAppProc="autohotkey"
    #"^(firefox|palemoon)$"
    
    if ($NoRestart) {
        write-verbose -verbose:$true "$(Get-Timestamp): -NoRestart specified."
    } ; 
    write-verbose -verbose:$true "$(Get-Timestamp): --PASS STARTED:$ScriptName --"
    write-verbose -verbose:$true "killing $($TargAppName)"
    write-verbose -verbose:$true "killing autohotkey"
    # the hammer
    #Invoke-Item c:\path-to\pskill autohotkey;
    # via ps
    #get-process autohotkey -ea silentlycontinue| stop-process 
    $prcs=get-process -ea silentlycontinue| ?{$_.name -match $TargAppProc} ;
    $prcs | select Name ; 
    $prcs | stop-process -verbose;
    start-sleep 2
    $prcs=get-process -ea silentlycontinue| ?{$_.name -like $TargAppProc} 
    if ($prc) {write-verbose -verbose:$true "ZOMBIES:" ;$prc}
    write-verbose -verbose:$true "Launching Startup $TargAppName .lnks..."

    if ($NoRestart) {
      # drop
    } else {
        write-verbose -verbose:$true "Launching Startup $TargAppName .lnks..."
        # ahk code
        $ahks = get-childitem $StartFldr | ?{($_.Extension -like '*.lnk') -AND ($_.Name -like '*.ahk*')} ;
        if($ahks){ $ahks | %{$_.Name ; invoke-item $_.FullName } }
        
     } # if-E No-NoRestart
    write-verbose -verbose:$true "$(Get-Timestamp): --PASS COMPLETED --"
} ; 
# config an alias for it
new-Alias rahk restart-ahk
#*------^ END Function restart-ahk ^------

#*------v Function restart-ff v------
Function restart-ff {

    <# 
    .SYNOPSIS
    restart-ff - Close all autohotkey processes, and reslaunch any *.ahk.lnk files in the "$env:APPDATA\...Startup" folder
    .NOTES
    Written By: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka
    Change Log
    * 7:20 AM 11/23/2015 test $PMpub before invoking
    # 7:21 AM 6/12/2015 port in support for the -NoRestart switch from restart-itsm
    * 6:30 AM 5/20/2015 shifted to the function with a call; matches profile version
    * 6:28 AM 5/20/2015 fixed duped gci $ffpub
    * 6:57 AM 5/18/2015 add ff-public support for work
    * 9:51 AM 5/16/2015: update variables, add qcklaunch fldr support
    * 9:51 AM 5/16/2015: adapt to firefox
    * 7:22 AM 4/21/2015 parallel ff-targeting version initial quick pass - simplest version
    * 8:47 AM 1/15/2015
    * 7:46 AM 12/12/2014
    * kill and relaunch all .ahk files in the Startup menu

    .DESCRIPTION
    restart-ff - Close all autohotkey processes, and reslaunch any *.ahk.lnk files in the "$env:APPDATA\...Startup" folder
    .PARAMETER NoRestart
    Parameter to suppress re-open[-NoRestart switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    .\ahk-restart
    .LINK
    *------^ END Comment-based Help  ^------ #>

    Param([Parameter(HelpMessage='NoRestart [$switch]')]
    [switch] $NoRestart)  ; 
    
    $ScriptName=$myInvocation.ScriptName ;
    #$StartFldr="C:\Users\Account\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup";
    $StartFldr="$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"; #gci $StartFldr;
    #9:49 AM 5/16/2015 add quick launch fldr
    $QckFldr="$env:userprofile\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch" ;
    $TargAppName="Firefox or Palemoon"
    #"autohotkey"
    $TargAppProc="^(firefox|palemoon)$"
    #"autohotkey"
    
    if ($NoRestart) {
          write-verbose -verbose:$true "$(Get-Timestamp): -NoRestart specified."
      } ; 
    write-verbose -verbose:$true "$(Get-Timestamp): --PASS STARTED:$ScriptName --"
    write-verbose -verbose:$true "killing $($TargAppName)"
    $prcs=get-process -ea silentlycontinue| ?{$_.name -match $TargAppProc} ;
    $prcs | select Name ; 
    $prcs | stop-process -verbose;
    start-sleep 2
    $prcs=get-process -ea silentlycontinue| ?{$_.name -like $TargAppProc} 
    if ($prc) {write-verbose -verbose:$true "ZOMBIES:" ;$prc}
    
    if ($NoRestart) {
        # drop
    } else {
          write-verbose -verbose:$true "Launching Startup $TargAppName .lnks..."

          $PMs = get-childitem $QckFldr |?{($_.Extension -eq '.lnk') -AND ($_.Name -match "^Pale\sMoon.*PrivProf.*$")} ;
          $FFs = get-childitem $QckFldr | ?{($_.Extension -eq '.lnk') -AND ($_.Name -match "^Firefox.*PrivProf.*$")}
          if(($env:COMPUTERNAME) -eq "LYN-3V6KSY1") {
              # 11:05 AM 8/17/2015 skip ff use, to force PM, remed $ffpub
              $FFs = $null ; 
              # 6:28 AM 5/20/2015 fixed duped gci $ffpub
              #$FFpub = get-childitem $QckFldr | ?{ ($_.Extension -eq '.lnk') -AND ($_.name -match "Mozilla Firefox.LNK") } ; 
              # 10:57 AM 8/17/2015 abandoning ff, for pm
              $PMpub = get-childitem $QckFldr | ?{ ($_.Extension -eq '.lnk') -AND ($_.name -match "^Pale\sMoon.LNK") } ; 
              
          }  ;  # if-E LYN-3V6KSY1
          
          if($PMs){ 
            $PMs | %{$_.Name ; invoke-item $_.FullName } 
            if($PMpub){ $PMpub | %{$_.Name ; invoke-item $_.FullName } } ; 
          } elseif ($FFs){ 
            $FFs | %{$_.Name ; invoke-item $_.FullName } 
            if ($FFpub){ 
                $FFpub | %{$_.Name ; invoke-item $_.FullName } 
            } # if-E FFPub
          }; # if-E

    } # if-E No-NoRestart

    write-verbose -verbose:$true "$(Get-Timestamp): --PASS COMPLETED --"

} ;  # if-func block
# for profile use: config an alias for it
new-Alias rff restart-ff
# for freestanding restart-ff.ps1 use, just call the function:
#restart-ff
#*------^ END Function restart-ff ^------

#*------v Function get-password v------
Function get-password {
    <# 
    .SYNOPSIS
    get-password - Draw a random number of words from a password .xml file
    .NOTES
    Written By: Todd Kadrie
    Website:    <http://tinstoys.blogspot.com>
    Twitter:    <http://twitter.com/tostka>
    Additional Credits: [REFERENCE]
    Website:    [URL]
    Twitter:    [URL]
    Change Log
    # 9:01 AM 3/16/2016 debugged back to function
    # 8:33 AM 3/16/2016 added optional -times/-t param, specifying number of passpharses to produce
    # 8:19 AM 3/16/2016 add optional -reportLength/-rl 'length' reporting to the output string
    # 9:29 AM 2/22/2016 ported in from uwps get-password.ps1 minior cleanup
    # 2:21 PM 11/20/2015 added Dave Wyatt's Get-CryptoRandom() set
    9:13 PM 12/10/2014
    .DESCRIPTION
    get-password - Draw a random number of words from password text file (one column of source words).
    .PARAMETER  wordscount
    Number of Words to be used in passphrase
    .PARAMETER  lexicon
    Source Words file (in form of an .xml file)[beale|dice|full-path-to-file.xml)]]
    .PARAMETER  Subst
    Switch parameter that indicates to perform stock substitution on the generated passphrase [-subst] 
    .PARAMETER  dialect
    Dialect to convert passphrase [english]
    .PARAMETER  times
    Number of proposed passwords to produce
    .PARAMETER  reportLength
    Propose password and report on length
    .PARAMETER showDebug
    Switch parameter that indicates to display Debugging output[-showdebug] 
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    System.Boolean
    .EXAMPLE
    . get-password -words 3  ; 
    Pull a password combining 2 words from the target file (positional parameter)
    .EXAMPLE
    get-password -words 3 -showdebug 
    Pull a password combining 3 words from the target file, and show debug messages
    .EXAMPLE
    get-password -words 3 -d english -lex beale -showdebug 
    Pull a 3-word passphase from the Beale lex file, using the english dialect, and show debugging output
    .EXAMPLE
    get-password -words 3 -d irish -lex english -showdebug -subst
    Pull a 3-word passphase from the dice lex file, with character substitution, and show debugging output
    .EXAMPLE
    1..10 | %{ .\get-passwords -words 2} ; 
    Pull a set of 10 passwords (now bundled via -times 2 param)
    .EXAMPLE
    1..15|%{  get-password -w 3 |%{ "$($_.length)chars:`t$_" } } ;
    Pull 15 3-word pw's and echo ea's length as well as well (now bundled via '-times 15' & '-reportLength' parameters)
    .EXAMPLE
    .\get-passwords -words 2 -times 10 ; 
    Pull a set of 10 2-word passwords
    .EXAMPLE
    .\get-passwords -words 3 -times 15 -reportLength ; 
    Pull 15 3-word pw's and echo ea's length as well as well (now bundled via '-times 15' & '-reportLength' parameters)

    *----------^ END Comment-based Help  ^---------- #>

    PARAM(
      [alias("w")]
      [alias("words")]
      [int]$wordscount=2
      ,[alias("lex")]
      [string]$lexicon="CUSTOM"
      ,[alias("d")]
      [string]$dialect="english"
      ,[alias("s")]
      [switch]$subst
      ,[alias("rl")]
      [switch]$reportLength
      ,[alias("t")]
      [int]$times=1
      ,[switch]$showDebug  
    ) # PARAM BLOCK END

    if(test-path function:get-cryptorandom){
        if ($showDebug) {
            write-debug "`$showDebug is $true. `nSetting `$DebugPreference = 'Continue'" ; 
            $DebugPreference = "Continue" ; 
            $bDebug=$true ; 
        } else {
          $DebugPreference = "SilentlyContinue";
        };

$sMsg=@"

`$wordscount:$wordscount`t`$lexicon:$lexicon`t`$subst:$subst`t`$reportLength:$reportLength
`$times:$times`t`$showDebug:$showDebug
"@ ; 
        if($bDebug){Write-Debug $sMsg ; }  ; 
        
        
        # stock location for words .xml files
        $LexPath="C:\path-to\" ; 
        if($bDebug){Write-Debug "`$LexPath:$LexPath"}  ; 

        # defined lexicon filename options:
        switch ($lexicon) {
          "CUSTOM" { $wordlist=$(join-path -path $LexPath -childpath "CUSTOMWORDS.xml") }
          "beale" {$wordlist=$(join-path -path $LexPath -childpath "beale-list.xml") }
          "dice" {$wordlist=$(join-path -path $LexPath -childpath "diceware.xml")}
          default {
              # validate that the item passed in is a fully pathed xml file with words in the correct format
              if(test-path $lexicon) {
                  if( get-content $lexicon | Out-String | ?{$_ -match ".*\<string\>.*"} ) {
                      write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Variant Lexicon file specified`nUsing `$lexicon:$lexicon" ;
                  } else {
                      write-error "$((get-date).ToString("HH:mm:ss")):INVALID/MISSING `$wordlist:$($wordlist)"; exit ;
                  } ; 
              } ; 
          }
        } ; 
        if($bDebug){Write-Debug "`$wordlist:$wordlist `n"}  ; 
        if(($bDebug) -AND ($subst)){Write-Debug "`$subst:$True `n"}  ; 

        if(!(test-path $wordlist)){ write-error "$((get-date).ToString("HH:mm:ss")):MISSING `$wordlist:$($wordlist)"; exit ; }
        if($times -gt 1){write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Producing $($times) passphrases" };
        
        [array]$outReport = $null ; 
        $i=0 ; 
        Do {
            $passphrs = $null ; 
            $i++ ; 
            if($bDebug){Write-Debug "`$i:$i"}  ; 
            # just pull x random lines out of the .xml, trim it, and strip the <string></string> tags; titlecase and then concat them, preparse out the non-<string> lines
            $words=Get-Content $wordlist |?{$_ -match ".*\<string\>.*\</string\>"} ;
            if($bDebug){Write-Debug "`$words.count:$(($words | measure).count)"}  ; 

            <# 9:32 AM 11/19/2015 aborted attempt to parse every word out of tags, slow wastes time on text overhead #>
            $a=1 ; 
            Do {
              # do the dialects in here, one word at a time
              $newword=(($words|get-CryptoRandom).trim().replace("<string>","").replace("</string>","")).toLower()
              switch ($dialect) {
                  "english" { 
                    # drop through 
                    } 
                  
              }  # SW-E; 
              
              if(($bDebug) -AND ($dialect -ne "english")){Write-Debug "`$newword:$newword"}
              
              # 9:27 AM 11/19/2015 instead of stripping tags, just work with raw words SLOOOOWWWW, faster to just strip 1/wrd than strip all
              $passphrs+=$((Get-Culture).TextInfo).ToTitleCase($newword)
              $a++ ;
            } Until ($a -gt $wordscount) ; 
            
            # optional subst rules
            if($subst){
                if($bDebug){Write-Debug "Performing Substitution on original passphrase: $passphrs..."}
                # perform standard substitution on the phraze
                # replace 1st vowel, locate it
                $passphrs -match "([REGEX OF CHARS TO BE SUBST-OUT HERE])" ; 
                $v1=$passphrs.indexof($matches[1]) ; 
                if($bDebug){Write-Debug "`$v1:$v1"}  ; 
                # from 0 to $f1
                $ppre=$passphrs.substring(0,$v1) ; 
                if($bDebug){Write-Debug "`$ppre:$ppre"}  ; 
                # from $v1+1 to .length
                $ppost=$passphrs.Substring($v1+1,$passphrs.Length-($v1+1))
                if($bDebug){Write-Debug "`$ppost:$ppost"}  ; 
                $passphrs=$ppre ; 
                <# TRIMMED, SETUP YOUR OWN SUBST SPECIFICATION
                switch ($matches[1]){
                  "X" {$passphrs+="1"}
                } ; 
                #>
                $passphrs+=$ppost ; 
                
                $passphrs=$((Get-Culture).TextInfo).ToTitleCase($passphrs) ; 
            } ;
            if($bDebug){Write-Debug "`$passphrs:$passphrs"}  ;
            
            # 8:19 AM 3/16/2016 add optional -reportLength/-rl 'length' reporting to the output string
            if($reportLength){
                $outReport +="[{0:D2}: {1:D2} chars]:`t{2}" -f $i,$($passphrs.length),$($passphrs) ; 
            } else {
                $outReport +=$passphrs ; 
            } ; 
            if($bDebug){Write-Debug "`$outReport:$($outReport|out-string)"}  ;
        } Until ($i -gt $times) ; # loop-E 
        
        # 8:49 AM 3/16/2016 we've accumulated the loops, dump them to pipeline
        $outReport | write-output 
        
        # wipe the variables
        $varis = "words","lexicon","lexpath","ppre","ppost","passphrs" ; 
        foreach ($vari in $varis) {
          if(get-variable $vari -ea 0) {if($bDebug){Write-Debug "$vari"} ; clear-variable $vari}; 
        } ; 
    } ;  # if-E get-crytorandom test
} #*------^ END Function get-password ^------ ; 

#*======v Generic Desktop Aliases (scripts etc) v======
new-Alias ctw close-taskwindow.ps1
#*======^ END Function Generic Desktop Aliases (scripts etc) ^======

#*======^ *** DESKTOP FUNCTIONS *** ^======
#           =========

#           =========
#*======v SUB MAIN v======

# 12:38 PM 5/24/2016 Sub Main: adding Git support & Posh-Git
Write-Host "Setting up GitHub Environment" ;
. (Resolve-Path "$env:LOCALAPPDATA\GitHub\shell.ps1") ;


#*======^ END SUB MAIN ^======
#           =========

