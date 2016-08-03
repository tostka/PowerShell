# c:\usr\home\db\xxx-prof.ps1

#*======v NON-ADMIN: C:\Users\Account\Documents\WindowsPowerShell\profile.ps1 v======
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\Account\Documents\WindowsPowerShell\profile.ps1
# NON-ADMIN acct $profile.CurrentUserCurrentHost loc
#C:\Users\Account\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1

# ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\Accounts\Documents\WindowsPowerShell\profile.ps1
# notepad2 $profile.CurrentUserAllHosts ;
#*======

#*------V Comment-based Help (leave blank line below) V------ 

<# 
    .SYNOPSIS
xxx-prof.ps1 - included personal profile file

    .NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Additional Credits: [REFERENCE]
Website:	[URL]
Twitter:	[URL]

Change Log
# 1:44 PM 5/24/2016 shifted storage from uhd => C:\sc\powershell\PSProfileUID
# 9:13 AM 4/7/2016 Get-Modules: expand -li to full -ListAvailable, select just the names, too long a baloney output list
# 2:30 PM 4/6/2016 added & then re-moved the persistent history matl back here and then over to xxxsid-incl-servercore.ps1
# 7:38 AM 2/8/2016 updated get-mbxstatus-related aliases. (there's also mbxs in AutoCorrectENG.ahk too)
# 1:38 PM 1/29/2016 moved update-persistenthistory call over to xxxsid-incl-ServerCore.ps1, below the function itself
# 12:25 PM 1/29/2016 adding persistent History material & $HistoryFilePath, w or wo Readline support (below pre-cleans out errored History entries)
# 12:38 PM 11/5/2015 added identifying banner
# 11:50 AM 11/5/2015 cleanedup detection logic for $InclSIDDirL, added regex's for Edge L13 support etc
# * 8:37 AM 10/30/2015 trimmed length of --/==
# 9:05 AM 8/25/2015 major overhaul added PSReadline, get-admincredential etc. Disabled PSReadLine dblquote scriptblock (bugs in psrl)
# 2:15 PM 2/19/2015 added code to detect run in PS ISE (suppress changing ISE win title to console title)
# 11:38 AM 1/7/2015 updated profile doc code
# 12:58 PM 12/31/2014 added missing \ on lync share
# 10:52 AM 12/31/2014 add support for Lync shared scripts on fileshare "\\lynmsv10.global.ad.WORKDOM.com\Lync_FS\scripts"
# 9:03 AM 12/30/2014 sub'd ulb for $binpath, and uhd for $InclPath, added $TextEd vari for notepad2/notepad
# 8:20 AM 12/30/2014 more clean-up, replaced comeent-based
# 7:30 AM 12/30/2014 cleaned up formatting & commenting.
# 2:00 PM 12/23/2014 split out into include files, shifted primary profile into include xxx-prof.ps1 as well
# 12:44 PM 12/23/2014 created c:\usr\home\db\xxx-prof.ps1 to host includes, and inturn be included into local profiles
# 12:23 PM 12/18/2014 added Function play-beep()
# 2:11 PM 12/3/2014 fixed get-timestamp (non-existent get-datetime())
# 9:27 AM 11/11/2014 added enable-lync
# 8:47 AM 11/11/2014 added test-user, and fixed a remote ex bug
# 6:53 AM 11/10/2014 added get-foldersize2 from home
# 2:38 PM 11/6/2014 added html funcs
# 2:29 PM 11/6/2014 ren'd WordWrapStr ->WordWrap-String, WordWrapWindowStr->WordWrap-WindowString
# 8:45 AM 10/31/2014 added get-timestamp, renamed Echo* functs to echo- to conform to verb-noun fmt (easier to remem), whoami->get-whoami
# 12:29 PM 10/28/2014 Edit-ProfileFile added full path to the launches below (won't use PWD automatically)
# 12:11 PM 10/28/2014 fixed LMS prompt, explicit cd $profile in the shorcut; had to edit it out by hand on each machine's LMS desktop shortcut...
# 11:50 AM 10/28/2014 remmed out PSCX-based Test-AdminLocal
# 11:24 AM 10/28/2014 edit-profile: fix for servers, by adding .exe
# 11:21 AM 10/28/2014 updated set-location testing to accommodate server-based EMS/LMS
# 9:38 AM 10/3/2014 added DL aliases
# 7:42 AM 9/8/2014 added get-imdb* aliases
# 7:30 AM 8/26/2014 added comment-based help block to top
# 7:18 AM 8/26/2014 get-lastsleep(): corrected date fmt string
# 2:49 PM 8/25/2014 added get-lastsleep(), also removed redundant mov functions from bottom, and setup EMS/LMS/PS cd's
# 1:04 PM 8/25/2014 added grant-mbx-vscan
# 1:23 PM 8/20/2014 added Read-Host2
# 1:09 PM 8/20/2014 added invoke-flasher()
# 9:59 AM 8/12/2014 added get-IPsettings()
# 8:42 AM 7/24/2014 fixed test-port() failed to returne $false on fail
# 11:48 AM 7/23/2014 added grant/ungrant-mbx()
# 10:45 AM 7/23/2014 added test-port()
# 9:12 AM 7/23/2014 add get-recipient aliases
# 1:17 PM 7/16/2014 organized functions into categories
# 12:32 PM 7/16/2014 renamed test-admin & test-isadmin to test-adminLocal() and test-AdminRunas(), added test-RDP()
# 11:53 AM 7/13/2014 fixed whoami converted to function with a return
# 11:01 AM 7/13/2014 tolower'd whoami, added imdb functions
# 10:30 AM 7/2/2014 added out-clipboard alias for the system clip.exe util
# 2:27 PM 6/4/2014 fix hard coded powershell window titles
# 8:53 AM 6/2/2014 updated to reflect user of CurrentUserAllHosts profile
# 7:50 AM 6/2/2014 added edit-profilefile/ep
# 7:41 AM 6/2/2014 added toggle-forestview & tfv & & profed aliases
# 9:04 PM 5/30/2014 added test-CommandExists()
# 8:21 AM 5/28/2014 validated & fixed go function
# 8:11 AM 5/28/2014 fixed set-location for vanilla ps & EMS et al
# 1:26 PM 5/27/2014 updated prompt() to have bw & color variant (for EMS/LMS etc)
# 12:51 PM 5/27/2014 added a test-path to the PCSX load
# 8:29 PM 5/23/2014 added quote-string & quote-list
# 7:14 PM 5/23/2014 home changes revision, switch on uid & machine, fixed pcx
# 7:18 AM 5/16/2014 added pause,echo & date/timestamp functions
# 11:29 AM 5/14/2014 updated fortune to include wrap etc
# 7:41 AM 5/14/2014 renamed alias pslist to g-proc
# 12:53 PM 5/8/2014 added gmbx, g-adu, gadu
# vers: 2:59 PM 5/1/2014 wrestling wrap function...
# vers: 8:25 AM 5/1/2014 added get-excuse, ren'd get-fortune,added download-file, tweaked admindetect with
# vers: 6:54 AM 4/28/2014 segregated admin-only and added g-proc(get-process)
# vers: 10:32 AM 3/27/2014 played with go() but it's borked UNDER ems, not cd's to it
# vers: 9:09 AM 3/19/2014 added alias mbx-stat & g-mbx,
# VERS: 10:47 AM 3/14/2014
# vers: 10:21 AM 3/14/2014 add cd & open funtions, set-, and more aliases to the mix
# vers: 12:54 PM 2/19/2014 added CheckLastExitCode() & get-Callstack() and cleanedup formatting/commenting
# vers: 3:05 PM 1/15/2014 added move-window functions
# vers: 10:12 AM 1/13/2014: added out-excel()
# vers: 12:01 PM 11/6/2013 fixed prompt to multicolor, added functions
# note: EMS asserts it's own prompt() func, and overrides the below back to stock...
# vers: 3:08 PM 5/15/2013
# vers: 2:34 PM 5/15/2013

Profile-creation cmds
#create the CurrentUserAllHosts file:
New-Item -path $PROFILE.CurrentUserAllHosts -ItemType file -Force ; notepad2 $PROFILE.CurrentUserAllHosts ;
# create the CurrentUserCurrentHost (default) file, type:
New-Item -path $profile -itemtype file -force; notepad2 $PROFILE.CurrentUserCurrentHost ;

To implement, copy the include files to the $LocalInclDir & $LocalInclSIDDir, and then add this to $profile bottom:
#-------
$DomainWork = "WORKDOM";
$DomHome = "HOMEDOM";
$DomLab="WORKDOM-LAB";

$LocalInclDir="c:\usr\home\db" ;
$LocalInclSIDDir = "c:\usr\work\ps\scripts";
# distrib shares:

if (!(Test-Path $LocalInclDir)) {
  if (Test-Path $InclShareCent) {
    $LocalInclDir = $InclShareCent ; Write-Verbose -verbose "Using CentralShare Includes:$LocalInclDir";
  } elseif ($env:USERDOMAIN -eq "WORKDOM-LAB") {
    # add pretest for non-Lync lab
    if(test-path $inclShareLab){
      $LocalInclDir = $inclShareLab ; Write-Verbose -verbose "Using LabShare Includes:$LocalInclDir";
    } elseif (test-path $InclSIDDirL){
      # and then defer into lyncshare lab
      $LocalInclDir = $InclSIDDirL ; Write-Verbose -verbose "Using LabShareL13 Includes:$LocalInclDir";
    } else {
      Write-Warning "No available LAB include source found. Exiting..."; Exit;
    };
  } elseif (Test-Path $InclShareL13) {
    $LocalInclDir = $InclShareL13 ; Write-Verbose -verbose "Using L13 Includes:$LocalInclDir";
  } else {
    Write-Warning "No available include source found. Exiting..."; Exit;
  };
};

#---------

Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":LOADING INCLUDES:")
$sLoad=(join-path -path $LocalInclDir -childpath "xxx-prof.ps1") ; ;if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;};
#-------

.DESCRIPTION
xxx-prof.ps1 - included personal profile file
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
#>

write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):====== EXECUTING: $(Split-Path -Leaf ((&{$myInvocation}).ScriptName) ) ====== " ; 

#*======v CONTSTANTS v======
#$bDebug = $TRUE;
$bDebug = $FALSE;
# standard all machines should be ucase and all accts lower
# 8:25 AM 12/30/2014 don't define these in any other include files! (too confusing)
# define $binpath
if(Test-Path "c:\usr\local\bin"){
  $binpath="c:\usr\local\bin\"
} elseif(Test-Path "d:\scripts"){
  $binpath="d:\usr\local\bin\"
} elseif(Test-Path "c:\scripts"){
  $binpath="c:\scripts\"
} else {
  Write-Error -verbose ((Get-Date).ToString("HH:mm:ss") + ":NO LOCAL BINPATH FOUND. EXITING!:")
  Exit
} # if-E

# Computername contstants
# 8:00 AM 4/7/2016 Contstants: add 3rd home box to $MyBox
$MyBox="BOX1","BOX2"
$MyBoxW="BOX1" ;
$MyBoxH="BOX2"
$rgxProdL13Servers="^$" ; 
$rgxProdServers="^$" ; 
$rgxLabL13Servers="^$" ; 
$rgxLabServers="^$" ; 
$rgxProdL13EdgSrvrs="^$" ;
$rgxLabL13EdgSrvrs="^$" ;

# User ID constants
$AcctWAdmn="WORKDOM\Accounts";
# covers both prod & lab edge server names
$rgxAcctWAdmnEdg="^COMPUTER[0,1]((d)*)\\Accounts$"
$AcctWUser="WORKDOM\Account";
$AcctHAdmin = "XXX";

# Domain contstants
$DomainWork = "WORKDOM";
$DomHome = "HOMEDOM";
$DomLab="WORKDOM-LAB";
$DomL13EdgProd=$rgxProdL13EdgSrvrs ;
$DomL13EdgLab=$rgxLabL13EdgSrvrs ;

#$LocalInclDir="c:\usr\home\db" ;
# 1:46 PM 5/24/2016 shift to github sc locations
$LocalInclDir="C:\sc\powershell\PSProfileUID" ;
$LocalInclSIDDir = "c:\usr\work\ps\scripts";
# distrib shares:

$smtpserver = '' # SMTP server you want to use to send email
$smtpserverport = 25 ;

#*======^ END CONSTANTS  ^======

#*======v CALCULATED VARIABLES v======

switch ($env:USERDOMAIN){
    $DomLab {
        write-host "$DomLab Domain...";
        ($env:computername)
        switch -regex ($env:COMPUTERNAME){
            "$rgxLabL13Servers$" {
                $LocalInclDir = $InclSIDDirL ; Write-Verbose -verbose "Using LabShareL13 Includes:$LocalInclDir";
              };
            "$rgxLabServers" { 
                $LocalInclDir = $inclShareLab ; Write-Verbose -verbose "Using LabShare Includes:$LocalInclDir";
            };
            default { Write-Warning "No available LAB include source found. Exiting Profile load..."; Exit;};
        } # switch-ntry-E ($env:COMPUTERNAME)
    } # switch-ntry lab
    $DomainWork {
        write-host "$DomainWork Domain...";
        ($env:computername)
        switch -regex ($env:COMPUTERNAME){            
            "$rgxProdL13Servers" {
                $LocalInclDir = $InclShareL13 ; Write-Verbose -verbose "Using L13 Includes:$LocalInclDir";
            };
            "$rgxProdServers" { 
                $LocalInclDir = $InclShareCent ; Write-Verbose -verbose "Using CentralShare Includes:$LocalInclDir"; 
            };
            "$MyBoxW" { 
                $LocalInclDir = $LocalInclDir ; Write-Verbose -verbose "Using Existing Local Includes:$LocalInclDir"; 
            };
            default{ Write-Warning "No available PROD include source found. Exiting Profile load..."; Exit;};
        } # switch-ntry-E ($env:COMPUTERNAME)
    } # switch-ntry WORKDOM
    $DomL13EdgProd {
		$LocalInclDir = $binpath ; 
	} ; 
	$DomL13EdgLab {
		$LocalInclDir = $binpath ; 
	} ; 
	default{
        # no dom at home (workgroup), so just test compnames
        if($MyBox -contains $env:COMPUTERNAME){
            "$($env:computername): home network detected";
            # default spec at top will work: $LocalInclDir="c:\usr\home\db" ;
        }; # if-E
    } # switch-ntry-E default
} # switch-E ($env:USERDOMAIN)

#---------

Write-Verbose -verbose "`$LocalInclDir:$LocalInclDir";
Write-Verbose -verbose "`$LocalInclSIDDir:$LocalInclSIDDir";

# 8:56 AM 12/30/2014 define $TextEd vari pointed at notepad.exe or notepad2.exe
# $binpath="c:\usr\local\bin\" ; test-path "(join-path $binpath notepad2.exe)"
if (Test-Path (Join-Path -path $binpath -childpath notepad2.exe)) {
  $TextEd=(Join-Path -path $binpath -childpath notepad2.exe)
} else {
  # 11:24 AM 10/28/2014 fix for servers, by adding .exe
  $TextEd="notepad.exe"
} # if-block end

#*======^ END CALCULATED VARIABLES ^======

#*======v *** SERVER-SIDE CORE FUNCTIONS *** v======
# vers: 8:12 AM 12/30/2014 cleaned up
Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":LOADING INCLUDES:")
# constants file
$sLoad=(join-path -path $LocalInclDir -childpath "xxxsid-incl-ServerCore.ps1") ;if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;}
#*======^ *** SERVER-SIDE CORE FUNCTIONS *** ^======
#                 ==============

#                 ==============
#*======v *** COMMON FUNCTIONS *** v======


#*======^ *** COMMON FUNCTIONS *** ^======
#                 ==============


#*======v FUNCTION SUB MAIN (PRIMARY SCRIPT EXECUTION) v======


# 11:29 AM 7/13/2014 pull get-whoami into a vari
$whoami=(get-whoami).tolower() ; #$whoami
$bRDP = test-RDP ; # validate if this is in an RDP session.

Write-Verbose -verbose ("Acct: " + $whoami)
# 11:15 AM 11/5/2015 on edge boxes, $whoami is lynms5200\Accounts
if (($whoami -ne $AcctWAdmn) -and (!($MyBox -contains $env:COMPUTERNAME))) {
  $logonType -eq "MyNonAdmn"
} elseif (($whoami -eq $AcctWAdmn) -and (($MyBox -contains $env:COMPUTERNAME))) {
  $logonType="MyAdmin"
} elseif ( ( ($whoami -eq $AcctWAdmn) -or ($whoami -match $rgxAcctWAdmnEdg)) -AND (!($MyBox -contains $env:COMPUTERNAME)) ) {
	# Accounts, and not logged into my laptop -> server desktop
  $logonType="ServerAdmin"
} else {
  $logonType ="MyNonAdmn"
} # if-block end
Write-Verbose -verbose ("`logonType: " + $logonType)

# add pcx v3 (only if local)
#Import-Module Pscx -RequiredVersion 3.0.0.0
#12:48 PM 5/27/2014 check for existence before load
# 7:14 PM 5/23/2014 hand path it, above is borked
$pcsxMod='C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\Pscx\Pscx.psd1';
# 10:17 AM 11/5/2015 only load pcx on work workstation
if ((Test-Path $pcsxMod) -AND ($logonType -eq "MyNonAdmn")) {
	Write-Verbose -verbose "Loading PCSX...";Import-Module $pcsxMod
};

# 1:12 PM 1/29/2016 adding Persistent History support
#Update-HistoryPersistant ; 
# 1:31 PM 1/29/2016 not consistently working from here, shift it into firing right below the function declare

#*------v ALIASES v------
if(!(get-alias | ?{$_.name -like "gcid"})) {Set-Alias -Name 'gcid' -Value 'Get-ChildItemDirectory' ; } ;
# 12:07 PM 10/28/2014 remmed, causing errors: set-alias : The AllScope option cannot be removed from the alias 'wget'.
#set-alias wget Get-WebItem
if(!(get-alias | ?{$_.name -like "ss"})) {Set-Alias -Name 'ss' -Value 'Select-String' ; } ;
if(!(get-alias | ?{$_.name -like "ssr"})) {Set-Alias -Name 'ssr' -Value 'select-StringRecurse' ; } ;
#set-alias go Jsh.Go-Path
#set-alias gop Jsh.Push-Path
if(!(get-alias | ?{$_.name -like "script"})) {Set-Alias -Name 'script' -Value 'Jsh.Run-script' ; } ;
if(!(get-alias | ?{$_.name -like "ia"})) {Set-Alias -Name 'ia' -Value 'Invoke-Admin' ; } ;
if(!(get-alias | ?{$_.name -like "ica"})) {Set-Alias -Name 'ica' -Value 'Invoke-CommandAdmin' ; } ;
if(!(get-alias | ?{$_.name -like "isa"})) {Set-Alias -Name 'isa' -Value 'Invoke-ScriptAdmin' ; } ;
if(!(get-alias | ?{$_.name -like "grep"})) {Set-Alias -Name 'grep' -Value 'Select-String' ; } ;

# 8:08 AM 2/23/2016 fix err: Out-Clipboard : The term '$env:SystemRoot\System32\clip.exe' is not recognized as the name of a cmdlet
#if(!(get-alias | ?{$_.name -like "Out-Clipboard"})) {Set-Alias -Name 'Out-Clipboard' -Value '$env:SystemRoot\System32\clip.exe' ; } ;
if(!(get-alias | ?{$_.name -like "Out-Clipboard"})) {
    if(test-path -path "$($env:SystemRoot)\System32\clip.exe"){
        Set-Alias -Name 'Out-Clipboard' -Value "$((get-childitem $env:SystemRoot\System32\clip.exe).fullname)" ; 
    } else {
        write-error "$((get-date).ToString("HH:mm:ss")):Unresolvable clip.exe, skipping out-clipboard alias";
    } ; 
} ;


# alias my functions
if (Test-Path function:Invoke-pause) {if(!(get-alias | ?{$_.name -like "pause"})) {Set-Alias -Name 'pause' -Value 'Invoke-pause' ; } ;}
if(!(get-alias | ?{$_.name -like "g-proc"})) {Set-Alias -Name 'g-proc' -Value 'Get-Process' ; } ;
if(!(get-alias | ?{$_.name -like "notepad"})) {Set-Alias -Name 'notepad' -Value '$TextEd' ; } ;
if (Test-Path function:gotoIncid){if(!(get-alias | ?{$_.name -like "incid"})) {Set-Alias -Name 'incid' -Value 'gotoIncid' ; } ;}
if (Test-Path function:reload-Profile){if(!(get-alias | ?{$_.name -like "reload"})) {Set-Alias -Name 'reload' -Value 'reload-Profile' ; } ;}
# if (test-path function:){}
if (Test-Path function:edit-File){
if(!(get-alias | ?{$_.name -like "edit"})) {Set-Alias -Name 'edit' -Value 'edit-File' ; } ;
if(!(get-alias | ?{$_.name -like "v"})) {Set-Alias -Name 'v' -Value 'edit-File' ; } ;
};
if (Test-Path function:set-ConsoleSmall){if(!(get-alias | ?{$_.name -like "scs"})) {Set-Alias -Name 'scs' -Value 'set-ConsoleSmall' ; } ;}
if (Test-Path function:set-ConsoleWide){if(!(get-alias | ?{$_.name -like "scw"})) {Set-Alias -Name 'scw' -Value 'set-ConsoleWide' ; } ;}
if (Test-Path function:edit-ProfileFile){if(!(get-alias | ?{$_.name -like "ep"})) {Set-Alias -Name 'ep' -Value 'edit-ProfileFile' ; } ;}
# 12:24 PM 1/12/2015 adding support for editing include prof
if (Test-Path function:Edit-ProfileFileInclude){if(!(get-alias | ?{$_.name -like "epi"})) {Set-Alias -Name 'epi' -Value 'Edit-ProfileFileInclude' ; } ;}
if (Test-Path function:Edit-ProfileFileInclude){if(!(get-alias | ?{$_.name -like "ep2"})) {Set-Alias -Name 'ep2' -Value 'Edit-ProfileFileInclude' ; } ;}

if ($logonType -eq "MyNonAdmn") {
  if (!($bRDP)) {
    # non-admin functions
    #set-alias exs gotoDevExScripts
    #set-alias lncs gotoDevLynScripts
    #set-alias dl gotoDownloads
    Set-Alias dbx gotoDbox
    Set-Alias dbxdb gotoDboxDb
    Set-Alias oin openInput
    Set-Alias word ‘C:\Program Files\Microsoft Office\Office14\WINWORD.EXE’
    Set-Alias powerpnt ‘C:\Program Files\Microsoft Office\Office14\POWERPNT.EXE’
    Set-Alias excel ‘C:\Program Files\Microsoft Office\Office14\EXCEL.EXE’
    Set-Alias outlook ‘C:\Program Files\Microsoft Office\Office14\OUTLOOK.EXE’
    Set-Alias firefox ‘C:\Program Files\Mozilla Firefox\firefox.exe’
    # with Call Operator & ; Why: Used to treat a string as a SINGLE command. Useful for dealing with spaces: & 'C:\Program Files\Windows Media Player\wmplayer.exe' "c:\videos\my home video.avi" /fullscreen
    function firefox-private {& 'C:\Program Files (x86)\Mozilla Firefox\firefox.exe' -no-remote -p ToddPriv}
    #Set-Alias firefoxp & 'C:\Program Files (x86)\Mozilla Firefox\firefox.exe' -no-remote -P ToddPriv
    Set-Alias firefoxp firefox-private
    Set-Alias Iexplore 'C:\Program Files\Internet Explorer\iexplore.exe'
    Set-Alias ie 'C:\Program Files\Internet Explorer\iexplore.exe'
    Set-Alias Chrome 'C:\Users\Account\AppData\Local\Google\Chrome\Applicationchrome.exe'

    if (Test-Path function:Resolve-ImdbId){Set-Alias imdbid Resolve-ImdbId -Force}
    if (Test-Path function:Get-ImdbTitle){Set-Alias imdb Get-ImdbTitle -Force}
    if (Test-Path function:Open-ImdbTitle){Set-Alias imdb.com Open-ImdbTitle -Force}
  } ; # if-block end
} # if-block end non-admin
# 6:53 AM 4/28/2014 admin-only aliases
if (get-whoami -eq $AcctWAdmn) {
  Set-Alias g-mbx get-mailbox
  Set-Alias gmbx get-mailbox
  Set-Alias tfv toggle-ForestView
  #new-alias g-prc get-process # 7:40 AM 6/2/2014 seems to consistently error
  Set-Alias g-prc Get-Process
  Set-Alias mbx-stat get-mbxstatus.ps1
  # 8:24 AM 2/12/2015 att verb-noun version:
  Set-Alias get-mbxstat get-mbxstatus.ps1
  Set-Alias gmbxs get-mbxstatus.ps1
  Set-Alias g-adu get-ADUser
  Set-Alias gadu get-ADUser
  # 9:11 AM 7/23/2014 add get-recip shortcuts
  Set-Alias grcp get-recipient
  Set-Alias get-r get-recipient
  # 9:37 AM 10/3/2014 add get-distribution group scs
  Set-Alias gdl Get-DistributionGroup
  Set-Alias g-dl Get-DistributionGroup
  Set-Alias g-dlm Get-DistributionGroupMember
  Set-Alias gdlm Get-DistributionGroupMember
} # if-block end non-admin
#*------^ ALIASES ^------


# 12:03 PM 4/29/2013 forcing powershell's to a small width
# don't use it on EMS
#Set-ConsoleSmall

# 12:05 PM 10/28/2014 the $snapins & $modules varis are only in the prmopt (), force them here too
$ModsLoad=Get-Module ;
$SnapsReg=Get-PSSnapin -Registered ;
$SnapsLoad=Get-PSSnapin ;
# 12:27 PM 12/31/2014 major problem is CAN'T pickup on the Ems under 2010 from mods or snaps - because EMS is REMOTED!. So you need to detect it as a PSSession
# Detect EMS10(remote session):
$EMSSess= (Get-PSSession | where { $_.ConfigurationName -eq 'Microsoft.Exchange' }) ;

if (!($bRDP)) {
  # -------
  # add 3rd party snapins
  $snapins = @(
  "PowerGadgets",
  "NetCmdlets"
  )
  $snapins | ForEach-Object {
    if ( Get-PSSnapin -Registered $_ -ErrorAction SilentlyContinue ) {
      Add-PSSnapin $_
    } # if-block end
  } # for-loop end
  # -------
} ; # if-block end

Write-Verbose -verbose "Your Available Modules are..."
# 9:13 AM 4/7/2016 Get-Modules: expand -li to full -ListAvailable, select just the names, too long a baloney output list
Get-Module -ListAvailable | sort ModuleType,Name | format-table ModuleType,Name -auto | out-default ; 
Write-Verbose -verbose " ";

# Build prompt components & console window config
switch -regex (($env:UserName).toupper()){

  "^(Account|xxx)$" {
      #                 ==============
      #*======v *** TOY FUNCTIONS *** v======
      # moved to c:\usr\home\db\xxx-incl-Toys.ps1 - Toy Function Includes
      $sLoad=(join-path -path $LocalInclDir -childpath "xxx-incl-Toys.ps1") ; if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;}
      #*======^ *** TOY FUNCTIONS *** ^======
      #                 ==============

      #                 ==============
      #*======v *** DESKTOP FUNCTIONS *** v======
      # moved to c:\usr\home\db\xxx-incl-Desktop.ps1
      $sLoad=(join-path -path $LocalInclDir -childpath "xxx-incl-Desktop.ps1") ; if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;}
      #*======^ *** DESKTOP FUNCTIONS *** ^======
      #                 ==============

      #                 ==============
      #*======v *** HOME FUNCTIONS *** v======
      # moved to c:\usr\home\db\xxx-incl-Home.ps1
      # 7:37 AM 12/30/2014 the post echo is coming out between excuse and prompt; suppress it unless debugging
      $sLoad=(join-path -path $LocalInclDir -childpath "xxx-incl-Home.ps1") ; if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;}
      #*======^ *** HOME FUNCTIONS *** ^======
      #                 ==============

      #notepad++ has a SaveAsAdmin plugin that solves the elevation issues (accessing system files)
      $nppPath = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\notepad++.exe' -ErrorAction SilentlyContinue
      if (-not($nppPath)) {
          Write-Warning -Message 'Unable to create npp alias; NotePad++ may not be installed.'
      } Else {
          Set-Alias -Name 'npp' -Value $nppPath.'(default)'
      } ; 


      # 7:56 AM 2/19/2015 this overrides the PS ISE title too, test for it
      if($Host.Name -eq "Windows PowerShell ISE Host"){
            # don't touch the title ISE is: "Administrator: Windows Powershell ISE"
      } elseif ($host.name -eq 'ConsoleHost'){
        # window title
        $host.ui.RawUI.WindowTitle = ("PS Console")
        # customize title bar; adds the - so the module names added by Update-PSTitleBar don't run together
        # used if appending to suit module adds
        #$host.ui.rawui.windowtitle = $host.ui.rawui.windowtitle+" -"
      } # if-E
      
      #region Version specific settings/functions
      If ($psversiontable.psversion -eq '2.0') {
          $PSEmailServer =$smtpserver ;
          $PSEmailServerPort = $smtpserverport ;
          #create functions to add these modules as PowerShell 2.0 doesn't support module auto-loading
          <#
          Function Add-MSSQL {
              Import-Module -Name SQLPS -DisableNameChecking
              Update-PSTitleBar 'MSSQL'
          } ; 
          #>
      } ElseIf ($psversiontable.psversion -ge '3.0') {
          # set default parameters on various commands. See 'Help about_Parameters_Default_Values' for more info
          <#
          $PSDefaultParameterValues=@{
              'Format-Table:AutoSize'=$True;
              'Get-Help:ShowWindow'=$True;
              'Send-MailMessage:SmtpServer'=$smtpserver
          } ; 
          #>
          # prevents the ActiveDirectory module from creating the AD: PSDrive on import
          #$Env:ADPS_LoadDefaultDrive = 0  ; 
          
          If ($host.name -eq 'ConsoleHost') {
              If (Get-Module -ListAvailable -Name PSReadline) {
                  Import-Module -Name 'PSReadLine'  ; 
                  #*------v PSReadline Custom Keyhandlers v------
                  #==up arrow/down arrow search history (if non-blank line)
                  Set-PSReadlineKeyHandler -Key UpArrow -Function HistorySearchBackward
                  Set-PSReadlineKeyHandler -Key DownArrow -Function HistorySearchForward
                  
                  #==dbl quotes
                  <# 8:39 AM 8/25/2015 seems to be the source of this bug when typing any quote:
                    Exception:
                    System.Management.Automation.RuntimeException: Unable to find type [Microsoft.PowerShell.PSConsoleReadline]: make sure t hat the assembly containing this type is loaded.  at System.Management.Automation.ExceptionHandlingOps.CheckActionPreference(FunctionContext funcContext, Exception exc eption)
                  #>
                  <#
                  Set-PSReadlineKeyHandler -Chord 'Oem7','Shift+Oem7' `
                       -BriefDescription SmartInsertQuote `
                       -LongDescription "Insert paired quotes if not already on a quote" `
                       -ScriptBlock {
                    param($key, $arg)
                    $line = $null ; 
                    $cursor = $null ; 
                    [Microsoft.PowerShell.PSConsoleReadline]::GetBufferState([ref]$line, [ref]$cursor) ; 
                    if ($line[$cursor] -eq $key.KeyChar) {
                        # Just move the cursor
                        [Microsoft.PowerShell.PSConsoleReadline]::SetCursorPosition($cursor + 1) ; 
                    } else {
                        # Insert matching quotes, move cursor to be in between the quotes
                        [Microsoft.PowerShell.PSConsoleReadline]::Insert("$($key.KeyChar)" * 2) ; 
                        [Microsoft.PowerShell.PSConsoleReadline]::GetBufferState([ref]$line, [ref]$cursor) ; 
                        [Microsoft.PowerShell.PSConsoleReadline]::SetCursorPosition($cursor - 1) ; 
                    } ; 
                  } ;  # SB-E------- dbl quotes
                  #>
                  #==Mistake: Buffer to Hitory (save cmd unexec'd to History)
                  Set-PSReadlineKeyHandler -Key Alt+w `
                  -BriefDescription SaveInHistory `
                  -LongDescription "Save current line in history but do not execute" `
                  -ScriptBlock {
                      param($key, $arg) ; 
                      $line = $null ; 
                      $cursor = $null ; 
                      [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor) ; 
                      [Microsoft.PowerShell.PSConsoleReadLine]::AddToHistory($line) ; 
                      [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine() ; 
                  } # SB-E-------

                  #-=-=-=-=-=-=-=-=-=-Ctrl+Shift+v is going to conflict with clipomatic's binding=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
                  <# Insert text from the clipboard as a here string
                  Set-PSReadlineKeyHandler -Key Ctrl+Shift+v `
                  -BriefDescription PasteAsHereString `
                  -LongDescription "Paste the clipboard text as a here string" `
                  -ScriptBlock {
                      param($key, $arg) ; 
                      Add-Type -Assembly PresentationCore ; 
                      if ([System.Windows.Clipboard]::ContainsText()) {
                          # Get clipboard text - remove trailing spaces, convert \r\n to \n, and remove the final \n.
                          $text = ([System.Windows.Clipboard]::GetText() -replace "\p{Zs}*`r?`n","`n").TrimEnd() ; 
                          [Microsoft.PowerShell.PSConsoleReadLine]::Insert("@'`n$text`n'@") ; 
                      } else {
                          [Microsoft.PowerShell.PSConsoleReadLine]::Ding() ; 
                      } ; 
                  } # SB-E--Insert CB As Herestring---------
                  #>
                  
                  #==ParenthesizeSelection=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
                  # Sometimes you want to get a property of invoke a member on what you've entered so far
                  # but you need parens to do that. This binding will help by putting parens around the current selection,
                  # or if nothing is selected, the whole line.
                  Set-PSReadlineKeyHandler -Key 'Alt+(' `
                  -BriefDescription ParenthesizeSelection `
                  -LongDescription "Put parenthesis around the selection or entire line and move the cursor to after the closing parenthesis" `
                  -ScriptBlock {
                      param($key, $arg) ; 
                      $selectionStart = $null ; 
                      $selectionLength = $null ; 
                      [Microsoft.PowerShell.PSConsoleReadLine]::GetSelectionState([ref]$selectionStart, [ref]$selectionLength) ; 
                      $line = $null ; 
                      $cursor = $null ; 
                      [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor) ; 
                      if ($selectionStart -ne -1) {
                          [Microsoft.PowerShell.PSConsoleReadLine]::Replace($selectionStart, $selectionLength, '(' + $line.SubString($selectionStart, $selectionLength) + ')') ; 
                          [Microsoft.PowerShell.PSConsoleReadLine]::SetCursorPosition($selectionStart + $selectionLength + 2) ; 
                      } else {
                          [Microsoft.PowerShell.PSConsoleReadLine]::Replace(0, $line.Length, '(' + $line + ')') ; 
                          [Microsoft.PowerShell.PSConsoleReadLine]::EndOfLine() ; 
                      } ; 
                  } # SB-E----ParenthesizeSelection-------
                  
                  #==ExpandAliases
                  Set-PSReadlineKeyHandler -Key "Alt+%" `
                  -BriefDescription ExpandAliases `
                  -LongDescription "Replace all aliases with the full command" `
                  -ScriptBlock {
                      param($key, $arg) ; 
                      $ast = $null ; 
                      $tokens = $null ; 
                      $errors = $null ; 
                      $cursor = $null ; 
                      [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$ast, [ref]$tokens, [ref]$errors, [ref]$cursor) ; 
                      $startAdjustment = 0 ; 
                      foreach ($token in $tokens) {
                          if ($token.TokenFlags -band [System.Management.Automation.Language.TokenFlags]::CommandName) {
                              $alias = $ExecutionContext.InvokeCommand.GetCommand($token.Extent.Text, 'Alias') ; 
                              if ($alias -ne $null) {
                                  $resolvedCommand = $alias.ResolvedCommandName ; 
                                  if ($resolvedCommand -ne $null) {
                                      $extent = $token.Extent ; 
                                      $length = $extent.EndOffset - $extent.StartOffset ; 
                                      [Microsoft.PowerShell.PSConsoleReadLine]::Replace($extent.StartOffset + $startAdjustment, $length, $resolvedCommand) ; 
                                      # Our copy of the tokens won't have been updated, so we need to
                                      # adjust by the difference in length
                                      $startAdjustment += ($resolvedCommand.Length - $length) ; 
                                  } # if-E ; 
                              } # if-E ; 
                          } # if-E ; 
                      } # loop-E ; 
                  } # SB-E----ExpandAliases-------
                  #*------^ END PSReadline Custom Keyhandlers ^------
              } ; 
          } ; 
      }#endregion  ; 

      # fortune
      # 11:13 AM 5/14/2014 updated to dewrap, strip redund spaces & trim fortune and wrap70
      if (Test-Path function:get-fortune){
        $phrase=(get-fortune);
        $exc=wrap-text (($phrase -creplace "\s+"," ") -creplace "`n"," ").trim() 70 ;
        Write-Host -fore green ("`n" + $exc + "`n");
        if (Test-Path function:speak-words){
          $phrase.replace("`n"," ") | speak-words ; 
        }  # if-E; 
      } # if-E;


      # 8:23 AM 5/1/2014 add an excuse too
      if (Test-Path function:get-excuse){

          Write-Host -fore DarkGray "Handy Excuse: " -nonewline
            $phrase="Handy Excuse: "
            $phrase+=(get-excuse);
            $nChars=70 ; if ($phrase.length -gt $nChars) {$phrase="`n"+(wrap-text $phrase $nChars)};Write-Host -fore red ($phrase);
            if (Test-Path function:speak-words){
              $phrase.replace("`n"," ") | speak-words ; 
            } ; 
      } # if-E ; 

  }  # switch-ntry-E "^(Account|xxx)$" ; 

  "^(AccountS)$" {
        # 1:28 PM 8/24/2015 issue here, is if we go with prompted $cred (for Lync https), the add-EMSRemote/add-LMSRemote()s are in here, and won't load
        $host.ui.RawUI.WindowTitle = ("PS ADMIN - " + [Environment]::UserDomainName)
        Write-Verbose -verbose "admin"
        # customize title bar; adds the - so the module names added by Update-PSTitleBar don't run together -- this will let us append module names into a delimited string to indicate what's loaded
        $host.ui.rawui.windowtitle = $host.ui.rawui.windowtitle+" -"


        #                 ==============
        #*======v *** SERVER-APP FUNCTIONS *** v======

        $sLoad=(join-path -path $LocalInclDir -childpath "xxxsid-incl-ServerApp.ps1") ;if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;}
        #*======^ *** SERVER-APP FUNCTIONS *** ^======
        #                 ==============
  } # switch-ntry-E "^(AccountS)$"; 

  default{
		# other
  } # switch-ntry-E ; 
} # switch block end (($env:UserName)

switch -regex (($env:COMPUTERNAME).toUpper()){
  "$MyBoxW" {
      # 8:03 AM 5/28/2014 if not exch/ems/lync, cd it c:\u\l\b
      # split lync off from exch
     if ( (Get-PSSession | where {$_.ConfigurationName -eq 'Microsoft.Exchange'}) -or (($SnapsLoad | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"})) ){
      Set-Location "c:\usr\work\exch\scripts"
    } elseif ((($ModsLoad | where { $_.Name -eq "Lync" }))) {
      Set-Location "c:\usr\work\lync\sscripts"
    } else {
		  #stock powershell
		  Set-Location "$binpath"
      }

  } # switch block end "$MyBoxW" 

  # 8:01 AM 4/7/2016 added tin-box to rgx
  "(BOX1|BOX2)" {
      Set-Location "$binpath"
  }  ; 

  default{
      if(Test-Path "c:\scripts") {Set-Location "c:\scripts"};
      if(Test-Path "$binpath") {Set-Location "$binpath"};
  } # switch block end default
} # switch block end (($env:COMPUTERNAME)

# set to working script dir
# ahh yes... this would be so nice if it was a built in variable
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

if ((Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"})) {
  # 12:03 PM 10/28/2014 try the code from prompt func
} elseif ($ModsLoad | where {$_.Name -eq "Lync"}){
  if($MyBox -contains $env:COMPUTERNAME.ToUpper()){
    Set-Location c:\usr\work\lync\scripts\ ;
  } else {
    # server LMS
    if(Test-Path "c:\scripts") {Set-Location "c:\scripts"};
  } # if-E

} elseif (Get-PSSession | where {$_.ConfigurationName -eq 'Microsoft.Exchange'}) {
  if($MyBox -contains $env:COMPUTERNAME.ToUpper()){
    Set-Location c:\usr\work\exch\scripts\ ;
  } else {
    # server EMS
    if(Test-Path "c:\scripts") {Set-Location "c:\scripts"};
  } # if-E
} else {
  # stock PS
  # set window title
  if ( Test-RunAsElevated ) { 
    Write-Host -foregroundcolor red "`n[* RunAsAdmin *]"
  } ;

  # 1:58 PM 8/25/2014 cd to ulb
  if(Test-Path "c:\scripts") {Set-Location "c:\scripts"};
  if(Test-Path "$binpath") {Set-Location "$binpath"};
} # if-block end

#*======^ END SUB MAIN (PRIMARY SCRIPT EXECUTION) ^======


