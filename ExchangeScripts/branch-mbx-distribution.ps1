# branch-mbx-distribution.ps1
# prompts for mbx, and runs folder count
# 9:08 AM 7/17/2015: reworking, adding transcript and optimizing, replaced all dupe qrys with recycled objects
# vers: 11:41 AM 2/28/2014 TOR port, switched org name to dynamic variable
# vers: 11:45 AM 7/12/2012 fixed formatting issues, and filtered out users on opposing version from test-mapi's
# vers: 1:34 PM 6/20/2012 funtional
# vers: 8:08 AM 6/18/2012 adapting to run branch mbx distribution reports in response to 'slow branch' questions
# vers: 10:47 AM 6/15/2012, adapted to folder count report
# vers: 9:26 AM 12/15/2011 add support for user spec'd as commandline arg
# vers: 12:28 PM 9/23/2011 added detection for whether E2010 or E2007 on basis of OS vers (R2=E2010)
# vers: 2:08 PM 9/15/2011: added test-mapiconnectivity and quote for test-exchangesearch
# vers: 11:35 AM 9/1/2011 pulled latency out, wildly innacurate compared to exmon
# vers: 9:41 AM 9/1/2011 added logonstats console dump
# vers: 8:16 AM 11/5/2010

#   - 8:16 AM 11/5/2010 added quota numbers in kb, too easy to misread the mb #'s and apply them 
#     to a mbx and push it into disabled.
#   - 2:25 PM 7/13/2010: added title from get-user
#   - 10:03 AM 6/10/2010: added forwarding fields (DeliverToMailboxAndForward,ForwardingAddress),
#       res fields (IsResource, IsLinked, IsShared,ResourceType)
#       added fields: GrantSendOnBehalfTo, HiddenFromAddressListsEnabled, WhenCreated
#       corrected a typo displaying Warn quota for prohsendreceive quota
#       added conditional color bolding to some fields if important status
#   - 11:53 AM 5/21/2010: added UseDatabaseQuotaDefaults to the get-mailbox
#   - 7:24 AM 5/3/2010: added size ref (mb) as trailing comment
#   - 8:15 AM 3/22/2010: initial

#*======v GLOBAL SPECS v======
$MbxHardLimit = "500" # don't pull more than this number of mbxs
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
$TimeStampNow = get-date -uformat "%Y%m%d-%H%M"
#*======^ END  GLOBAL SPECS  ^======


#*======v FUNCTIONS v======
#*----------v Function EMSLoadLatest v----------
  #Checks local machine for registred E20[10|07] EMS, and then loads the newest one found
  #Returns the string 2010|2007 for reuse for version-specific code

function EMSLoadLatest {
  # check registred vS loaded ;
  $SnapsReg=Get-PSSnapin -Registered ;
  $SnapsLoad=Get-PSSnapin ;
  # check/load E2013, E2010, or E2007, stop at newest (servers wouldn't be running multi-versions)
  if (($SnapsReg |
      ?{$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
  {
    if (!($SnapsLoad |
      ?{$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
    {
      Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue ; 
      return "2010" ;
    } else {
      return "2010" ;
    }
  } elseif (($SnapsReg |
      ?{$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"}))
  {
    if (!($SnapsLoad |
      ?{$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"}))
    {
      Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue ; 
      return "2007" ;
    } else {
      return "2007" ;
    }
  }
}
#*----------^END Function EMSLoadLatest ^----------

<# Preload EMS plugin:
#testing for OS revision to figure if loading e2010 or e2007
# or do it on R2 in the Name string, less likely to break on upgrades & patches
#((Get-WmiObject Win32_OperatingSystem).Name).indexof(" R2 ")  
if (((Get-WmiObject Win32_OperatingSystem).Name).indexof(" R2 ") -ne -1) {
  # it's E2010
    . \\chr100\pool\mis\messaging\exch2007\install-scripts\EMS-Preload10.ps1
    # 9:42 AM 7/12/2012 added E14 server version record
    $LocalServerVers = "E14"
} else {
  # it's e2007
. \\chr100\pool\mis\messaging\exch2007\install-scripts\EMS-Preload.ps1
  EMSPreload
  # ExchangeVersion : 0.1 (8.0.535.0)
   $LocalServerVers = "E8"
   write-host "LocalServerVers: " $LocalServerVers
}
#>

#*------v Function Test-Transcribing v------
#requires -version 2.0
# Author: Oisin Grehan
# URL: http://poshcode.org/1500
# Tests for whether transcript (start-transcript) is already running
# usage:
#   if (Test-Transcribing) {stop-transcript} ;
function Test-Transcribing {
  $externalHost = $host.gettype().getproperty("ExternalHost",
        [reflection.bindingflags]"NonPublic,Instance").getvalue($host, @())

  try {
    $externalHost.gettype().getproperty("IsTranscribing",
        [reflection.bindingflags]"NonPublic,Instance").getvalue($externalHost, @())
  } catch {
     write-warning "This host does not support transcription."
  }
} #*------^ END Function Test-Transcribing ^------

#*======^ END FUNCTIONS ^======

#*======v SUB MAIN v======



$EMSVers = EMSLoadLatest  # returns: "2013"|"2010"|"2007"
write-host ((get-date).ToString("HH:mm:ss") + ":`$EMSVers: " + $EMSVers)

$ExOrgName = (get-organizationconfig).Name

write-host " "

# 9:26 AM 12/15/2011 pretest for user as commandline arg
if ($Args.count -ge 1)
{
  # $Args[0] = $Args[0].ToSTring()
  $user = $Args[0].ToSTring()
  write-host -foregroundcolor "green" "User specified: " $user
} else {
  $user = Read-Host "Enter target user Acct,EmailAddr,etc"
  Write-Host -foregroundcolor "green" "Mailbox: " $user
  write-host " "
  
}

$outtransfile=$ScriptDir + "logs"
if (!(test-path $outtransfile)) {Write-Host "Creating dir: $outtransfile" ;mkdir $outtransfile ;} ;
#$outtransfile+="\" + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.log" ;
$outtransfile+="\" + $($user.ToUpper()) + "-" + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.log" ;
if (Test-Transcribing) {stop-transcript} ;
start-transcript -path $outtransfile ;



# echo the timestamp
#$TimeStampNow = get-date -uformat "%Y%m%d-%H%M"
#$TimeStampNow = get-date -uformat "%m/%d/%Y-%H:%M"
write-host -foregroundcolor "green"  (get-date -uformat "%m/%d/%y-%H:%M")
#06/19/12-13:31

write-host -foregroundcolor green "$((get-date).tostring("hh:mm:ss")):====== v PASS STARTED v ======";

write-host -foregroundcolor "green" "Pulling specified user's hosting db, server & branch"
# (Get-Mailbox kadrtod).database
$MbxUser = (Get-Mailbox $user) ;
#$db = (Get-Mailbox $user).database.name
$db = $MbxUser.database.name
#$server = (Get-Mailbox $user).servername
$server = $MbxUser.servername
# (get-user kadrtod).office
$branch = (get-user $user).office

$BranchMbxs=Get-Mailbox -Filter "Office -eq '$branch'" -resultsize $MbxHardLimit ;
if(($BranchMbxs | measure).Count -ge 1000) {
  write-host -foregroundcolor green "NOTE: LARGE BRANCH DETECTED - Only the first $($MbxHardLimit) mailboxes are being analyzed in these reports."; 
}  ; 

$bRet = $user + " details:"

write-host -foregroundcolor "green"  $bRet 
write-host "db: " $db
write-host "server: " $server
write-host "branch: " $branch
write-host " "

$MbxServerVers = ((get-mailbox $user).exchangeversion.exchangebuild.major)
# 9:38 AM 7/12/2012 record user mbx-revision:
#((get-mailbox $user).exchangeversion.exchangebuild.major)
if ($MbxServerVers -eq 14) {
  $MbxServerVersStr = "E14"
} elseif ($MbxServerVers -eq 8) {
  $MbxServerVersStr = "E8"
}

# test Ex revision:, can't run Ex2010 tests from Ex2007 server
if ((get-exchangeserver ($env:COMPUTERNAME)).admindisplayversion.major -ne ((get-mailbox $user).exchangeversion.exchangebuild.major)) {
  write-warning "E2010 mbx must test on E2010 EMS. E2007 mbx on E2007"
  write-warning "Exiting. Run the script from the same EMS revision as hosts the mailbox"
  write-host " "
  Exit 32
}


write-host -foregroundcolor "green" "Verify Outlook MAPI access locally..."
Test-MAPIConnectivity $user | select Server, Database,Mailbox, Result, @{label="Latency(ms)";expression={($_.Latency.Milliseconds)}} | ft -auto|out-default

write-host " "

# pull back logonstats
write-host -foregroundcolor "green" "Checking user's level of load being generated..."
$MbxUser | get-logonstatistics | select ClientMode, ClientVersion, CurrentOpenAttachments, CurrentOpenFolders, CurrentOpenMessages, LastAccessTime, @{label="Latency(ms)";expression={($_.Latency.Milliseconds)}} , LogonTime|ft –wrap|out-default ;
write-host " "

write-host -foregroundcolor "green" "Verify User's Exchange Search performance for the mbx..."
#Test-ExchangeSearch kadrtod | select Database, Server, Mailbox, ResultFound, SearchTimeInSeconds,Error | ft -auto
Test-ExchangeSearch $user | select Database, Server, Mailbox, ResultFound, SearchTimeInSeconds,Error | ft -auto |out-default

write-host " "

# 10:46 AM 7/12/2012 pull number of other users in the users database

write-host " "
write-host -foregroundcolor "yellow" "ARE ALL USERS IN DATABASE " $db " REPORTING ISSUES?"

$DbMbxs = (get-mailbox -database ((Get-Mailbox $user).database) -resultsize $MbxHardLimit ) ;
write-host -foregroundcolor "yellow" "The database hosts  $(($DbMbxs | measure).Count) _different_ users across $($ExOrgName)"
if(($DbMbxs | measure).Count -ge 1000) {
  write-host -foregroundcolor green "NOTE: LARGE DATABASE DETECTED - Only the first $($MbxHardLimit) mailboxes are being analyzed in these reports."; 
}  ; 
#write-host -foregroundcolor "yellow" ("The database hosts  $() _different_ users across " + $ExOrgName)
write-host " "

# Get-Mailbox -Filter "Office -eq '$OfficeName'" | sort Database 
# 7:54 AM 7/17/2015 retrieve into stock 1-qry obj
# 8:19 AM 7/17/2015 moved up to the branch block
#$BranchMbxs=Get-Mailbox -Filter "Office -eq '$branch'" -resultsize $MbxHardLimit ;

#write-host -foregroundcolor "green" "Number Users in '" ($branch) "':" (Get-Mailbox -Filter "Office -eq '$branch'" | measure).count
# 7:54 AM 7/17/2015
write-host -foregroundcolor "green" "Number Users in '" ($branch) "':$(($BranchMbxs | measure).count) `n" ;
#write-host " "

write-host -foregroundcolor "green" "Mailbox DATABASE distribution for users in '" ($branch) "':"

$DbsTtl = $BranchMbxs | group Database # | select Count,Name,Group | ft -auto |out-default

write-host " "
write-host -foregroundcolor "yellow" "ARE ALL PROBLEM USERS ON SAME DATABASE ABOVE?"

write-host -foregroundcolor "yellow" "Output above shows the branch is distributed across $(($DbsTtl | measure).Count)  _different_ databases:"
$DbsTtl | select Count,Name,Group | ft -auto |out-default

write-host " "
write-host -foregroundcolor "green" "Mailbox SERVER distribution for users in '" ($branch) "':"

$SrvrsTtl = $BranchMbxs| group servername | sort name | select Count,Name,Group #| ft -auto |out-default
$SrvrsTtl | ft -auto |out-default ;
write-host " "
write-host -foregroundcolor "yellow" "ARE ALL PROBLEM USERS ON SAME SERVER ABOVE?"
#write-host -foregroundcolor "yellow" "Output above shows the branch is distributed across " (Get-Mailbox -Filter "Office -eq '$branch'" | group servername | measure).Count " _different_ servers"
write-host -foregroundcolor "yellow" "Output above shows the branch is distributed across $( $SrvrsTtl | measure).Count) _different_ servers`n"
#write-host " "
write-host -foregroundcolor "yellow" "ARE ALL USERS IN BRANCH HAVING ISSUES?" 
write-host -foregroundcolor "yellow" "IF USERS CAN'T BE ISOLATED TO A SINGLE DATABASE OR SERVER," 
write-host -foregroundcolor "yellow" "THE ISSUES ARE LOCAL TO THE BRANCH -> CHECK WORKSTATION/LOCAL-NET/WAN (BITS/WAN GROUPS)" 

write-host " "
# check user's mbx contents
#write-host -foregroundcolor "green" "If Outlook performance, checking user-content that can slow down Outlook:"
write-host -foregroundcolor "green" "Outlook Load: Checking users's KEY performance-impacting folders for excessive message counts:"

# 8:47 AM 7/17/2015
$MbxUser| Get-MailboxFolderStatistics -FolderScope all | where {$_.name -match '^(Inbox|Deleted Items|Sent Items|Outbox|Calendar|Tasks)$' } | select @{Name='FolderPath';Expression={"..." + $_.Identity.tostring().substring($_.Identity.ToString().length -20,20)}},name,ItemsInFolderAndSubfolders | ft -auto | out-default ;


#write-host " "
write-host -foregroundcolor "green" "*Limits above: Inbox|SentItems|DeletedItems < 100k"
write-host -foregroundcolor "green" "*Limits above: Calendar|Tasks < 10k"


# finally loop the whole branch and run the Exchange tests for each user
write-host " "
write-host -foregroundcolor "green" "Validating Outlook Access for " $MbxServerVersStr " users in " ($branch) "..."

# just run all into one table

  # shorten latency.
  
  #$BranchMbxs | where {$_.exchangeversion.exchangebuild.major -eq $MbxServerVers} | measure | select count ;
  
  $BranchMbxs | where {$_.exchangeversion.exchangebuild.major -eq $MbxServerVers} | Test-MAPIConnectivity | select Mailbox,Database,Server, Result, @{label="Latency(ms)";expression={($_.Latency.Milliseconds)}} ,Error | ft -auto |out-default
  
  #write-host -foregroundcolor "green" "Checking client load from users in " ($branch) "..."
  # 10:23 AM 7/12/2012 just the matching users
  write-host -foregroundcolor "green" "Checking client load from " $MbxServerVersStr " users in " ($branch) "..."
  write-host " "
  #$BranchMbxs| get-logonstatistics | select UserName, ClientMode, ClientVersion, CurrentOpenAttachments, CurrentOpenFolders, CurrentOpenMessages, LastAccessTime, Latency, LogonTime|ft –wrap|out-default
  # drop the latency
  $BranchMbxs| get-logonstatistics | select UserName, ClientMode, ClientVersion, CurrentOpenAttachments, CurrentOpenFolders, CurrentOpenMessages | ft –wrap |out-default

write-host -foregroundcolor green "$((get-date).tostring("hh:mm:ss")):====== ^ PASS COMPLETED ^ ======";
stop-transcript ;
#*======^ END SUB MAIN ^======

