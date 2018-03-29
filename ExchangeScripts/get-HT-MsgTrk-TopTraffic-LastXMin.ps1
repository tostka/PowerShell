# get-HT-MsgTrk-TopTraffic-LastXMin.ps1

#*----------V Comment-based Help (leave blank line below) V---------- 

<# 
    .SYNOPSIS
get-HT-MsgTrk-TopTraffic-LastXMin.ps1 - Runs msgtracking on the local HubTransports for the designated Site & Exchange Revision, for the designated minutes of recent of traffic, then runs logparser to produce top-20 addresses in as either Sender, or Recipient in four categories of traffic.

    .NOTES
Written By: Todd Kadrie
Website: http://tinstoys.blogspot.com
Twitter: http://twitter.com/tostka

Change Log
  # 2:00 PM 10/2/2013 initial revision

    .DESCRIPTION
get-HT-MsgTrk-TopTraffic-LastXMin.ps1 - Runs msgtracking on the local HubTransports for the designated Site & Exchange Revision, for the designated minutes of recent of traffic, then runs logparser to produce top-20 addresses in as either Sender, or Recipient in four categories of traffic.

    .INPUTS
None. Does not accepted piped input.

    .OUTPUTS
None. Returns no objects or output.

    .EXAMPLE
.\get-HT-MsgTrk-TopTraffic-LastXMin.ps1 -SiteName US -ExTargVers 2007 -TrkTarget RECIPIENT -TrkMins 60

*----------^ END Comment-based Help  ^---------- #>

#------------- Declaration and Initialization of Global Variables ------------- 

PARAM(
  [alias("s")]
  [string] $SiteName=$(Read-Host -prompt "Specify Site to analyze [SiteName: US | EU | AP ]"),
  [alias("v")]
  [string] $ExTargVers=$(Read-Host -prompt "Specify Exchange Revision to target: [ExTargVers: 2007 | 2010 ]"),
  [alias("e")]
  [string] $TrkTarget=$(Read-Host -prompt "Specify 'Top' event type to filter: `n[TrkTarget: SENDER | RECIPIENT | BOUNCE | DEFERRED]"),
  [alias("m")]
  [string] $TrkMins=$(Read-Host -prompt "Specify minutes of recent traffic to filter: [TrkMins: [integer]]")
) # PARAM BLOCK END


## immed after param, setup to stop executing on first error
trap { break; }

# debugging output control
#$bDebug = $TRUE
$bDebug = $FALSE
If ($bDebug -eq $TRUE) {write-host ((get-date).ToString("HH:mm:ss") + ": *** DEBUGGING MODE ***")}

# *----------------v DELIMITED CONSTANTS v----------------
$SitesNameList = "US;AP;EU" ; $SitesNameList=$SitesNameList.split(";") ; 
$SitesList = "*US*;*AP*;*EU*" ; $SitesList=$SitesList.split(";") ; 
$ServersMbxNA2010="US-E2010EXCH1;US-E2010EXCH2" ; $ServersMbxNA2010=$ServersMbxNA2010.split(";") ; 
$ServersHCNAF="US-E2010HUBCAS1;US-E2010HUBCAS2" ; $ServersHCNAF=$ServersHCNAF.split(";") ; 
$ServersMbxNA="US-EXCH10;US-EXCH7;US-EXCH8" ; $ServersMbxNA=$ServersMbxNA.split(";") ; 
$ServersHCNA="US-HUBCAS1;US-HUBCAS2;US-HUBCAS3" ; $ServersHCNA=$ServersHCNA.split(";") ; 
$ServersMbxAP="AP-EXCH3" ; $ServersMbxAP=$ServersMbxAP.split(";") ; 
$ServersHCAP="AP-HUBCAS1;AP-HUBCAS2" ; $ServersHCAP=$ServersHCAP.split(";") ; 
$ServersMbxEU="EU-EXCH5;EU-EXCH6" ; $ServersMbxEU=$ServersMbxEU.split(";") ; 
$ServersHCEU="EU-HUBCAS1;EU-HUBCAS2" ; $ServersHCEU=$ServersHCEU.split(";") ; 
$IisLogsHCNA="\\US-HUBCAS1\e$\Weblogs\W3SVC1\;\\US-HUBCAS2\e$\Weblogs\W3SVC1\;\\US-HUBCAS3\e$\Weblogs\W3SVC1\" ; $IisLogsHCNA=$IisLogsHCNA.split(";") ; 
$IisLogsHCNAF="\\US-E2010HUBCAS1\e$\IIS Weblog\W3SVC1\;\\US-E2010HUBCAS2\e$\IIS Weblog\W3SVC1\" ; $IisLogsHCNAF=$IisLogsHCNAF.split(";") ; 
$IisLogsHCAP="\\AP-HUBCAS1\E$\WEBLOGS\W3SVC1\;\\AP-HUBCAS2\E$\WEBLOGS\W3SVC1\" ; $IisLogsHCAP=$IisLogsHCAP.split(";") ; 
$IisLogsHCEU="\\EU-HUBCAS1\E$\WEBLOGS\W3SVC1\;\\EU-HUBCAS1\E$\WEBLOGS\W3SVC1\" ; $IisLogsHCEU=$IisLogsHCEU.split(";") ; 
$MsgTrkLogsHCNA="\\US-HUBCAS1\F$\Program Files\Microsoft\Exchange Server\Logs\MessageTracking\;\\US-hubcas2\F$\Program Files\Microsoft\Exchange Server\Logs\MessageTracking\;\\US-HUBCAS3\e$\Weblogs\W3SVC1\" ; $MsgTrkLogsHCNA=$MsgTrkLogsHCNA.split(";") ; 
$MsgTrkLogsHCNAF="\\US-E2010HUBCAS1\e$\IIS Weblog\W3SVC1\;\\US-E2010HUBCAS2\e$\IIS Weblog\W3SVC1\" ; $MsgTrkLogsHCNAF=$MsgTrkLogsHCNAF.split(";") ; 
$MsgTrkLogsHCEU="\\EU-HUBCAS1\E$\Program Files\Microsoft\Exchange Server\Logs\MessageTracking\;\\EU-HUBCAS2\E$\Program Files\Microsoft\Exchange Server\Logs\MessageTracking\" ; $MsgTrkLogsHCEU=$MsgTrkLogsHCEU.split(";") ; 
$MsgTrkLogsHCAP="\\AP-HUBCAS1\E$\Program Files\Microsoft\Exchange Server\Logs\MessageTracking\;\\AP-HUBCAS2\E$\Program Files\Microsoft\Exchange Server\Logs\MessageTracking\" ; $MsgTrkLogsHCAP=$MsgTrkLogsHCAP.split(";") ;
$MsgTrkLogsHCNATemplate="\\US-HUBCASX\E$\WEBLOGS\W3SVC1\" ; 
$MsgTrkLogsHCNAFTemplate="\\US-E2010HUBCASX\E$\IIS WEBLOG\W3SVC1\" ;
$MsgTrkLogsHCAPTemplate="\\AP-HUBCASX\E$\Program Files\Microsoft\Exchange Server\Logs\MessageTracking\" ;
$MsgTrkLogsHCEUTemplate="\\EU-HUBCASX\E$\Program Files\Microsoft\Exchange Server\Logs\MessageTracking\" ;
# *----------------^ DELIMITED UNI CONSTANTS ^----------------

#*================v FUNCTION LISTINGS v================

#*----------v Function EMSLoadLatest v----------
<#
  Checks local machine for registred E20[13|10|07] EMS, and then loads the newest one found
  Returns the string 2013|2010|2007 for reuse for version-specific code
#>
function EMSLoadLatest {
  # check registred vS loaded ;
  $SnapsReg=Get-PSSnapin -Registered ;
  $SnapsLoad=Get-PSSnapin ;
  # check/load E2013, E2010, or E2007, stop at newest (servers wouldn't be running multi-versions)
  if (($SnapsReg| where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2013"}))
  {
    if (!($SnapsLoad | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2013"}))
    {
      Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2013 -ErrorAction SilentlyContinue ; return "2013" ;
    } else {
      return "2013" ;
    }
  } elseif (($SnapsReg| where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"})) 
  { 
    if (!($SnapsLoad | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"})) 
    {
      Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue ; return "2010" ;
    } else {
      return "2010" ;
    }
  } elseif (($SnapsReg| where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"}))
  {
    if (!($SnapsLoad | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"}))
    {
      Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue ; return "2007" ;
    } else {
      return "2007" ;
    }
  }
}
#*----------^END Function EMSLoadLatest ^----------

#---------------------------v Function Cleanup v---------------------------
Function Cleanup {
  # clear all objects and exit
  # Clear-item doesn't seem to work as a variable release 

  exit
} #*----------------^ END Function Cleanup ^----------------

#*----------------v Function ProcessLogs() v----------------
Function ProcessLogs {

 
PARAM(
  [alias("s")]
  [string] $SiteName=$(Read-Host -prompt "Specify Site to analyze [SiteName: US | EU | AP ]"),
  [alias("v")]
  [string] $ExTargVers=$(Read-Host -prompt "Specify Exchange Revision to target: [ExTargVers: 2007 | 2010 ]"),
  [alias("e")]
  [string] $TrkTarget=$(Read-Host -prompt "Specify 'Top' event type to filter: `n[TrkTarget: SENDER | RECIPIENT | BOUNCE | DEFERRED]"),
  [alias("m")]
  [string] $TrkMins=$(Read-Host -prompt "Specify minutes of recent traffic to filter: [TrkMins: [integer]")
) # PARAM BLOCK END
  
  # switch as per specified settings
  switch ($SiteName)
  {
    "US" {
      switch ($ExTargVers) 
      {
        "2007" {$Hubs = $ServersHCNA ; $MTLogs = $MsgTrkLogsHCNA }
        "2010" {$Hubs = $ServersHCNAF ;$MTLogs = $MsgTrkLogsHCNAF}
      } # switch block end
    } 
    "AP" {$Hubs = $ServersHCAP ; $MTLogs = $MsgTrkLogsHCAP}
    "EU" {$Hubs = $ServersHCEU ; $MTLogs = $MsgTrkLogsHCEU}
    defAPlt {write-warning "Invalid Site spec '$SiteName' Exiting." ; exit 1}
  } # switch block end
 
  write-host -ForegroundColor White ((get-date).ToString("HH:mm:ss") + ": Site: " + $SiteName)

  foreach ($Hub in $Hubs) {
    write-host -foregroundcolor White ("-" * 10)
    If ($bDebug -eq $TRUE) {write-host ((get-date).ToString("HH:mm:ss") + ": `$Hub: " + $Hub)}
    $ServerName = $Hub
    write-host -foregroundcolor yellow ((get-date).ToString("HH:mm:ss") + ": Parsing Server: " + $ServerName + " for the Top " + $TopUsers + " MsgTracking users in the last " + $TrkMins + " minutes (" + (get-date -format "yyMMdd") + ")...")
    
    #1. Setup search parameters and file names
    # Setup the messagetracking end time string (now):
    $TimeStampEnd= get-date -uformat "%m/%d/%Y %I:%M:%S %p"
    # strip any zeropadding
    if (($TimeStampEnd.ToSTring()).StartsWith('0')) {$TimeStampEnd=$TimeStampEnd.substring(1,$TimeStampEnd.length-1)}
    # setup a filename-compatible variant string
    $TimeStampEndFN= get-date -uformat "%Y%m%d-%H%M"
    
    # calculate a string for the designated number of minutes prior
    $TimeStampStart=(get-date).AddMinutes((-1 * $TrkMins))
    # and format it into the format necessary for the get-messagetrackinglog cmdlet
    $TimeStampStart=get-date ($TimeStampStart.ToString()) -uformat "%m/%d/%Y %I:%M:%S %p"
    
    if ($bDebug) { write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ":`$TimeStampStart: " + $TimeStampStart + "`n`$TimeStampEnd: " + $TimeStampEnd ) }
    # strip any zeropadding
    if (($TimeStampStart.ToSTring()).StartsWith('0')) {$TimeStampStart=$TimeStampStart.substring(1,$TimeStampStart.length-1)}
    # setup a filename-compatible variant string
    $TimeStampStartFN=get-date ($TimeStampStart.ToString()) -uformat "%Y%m%d-%H%M"
    # gen filename from script 
    $TimeStampNow = get-date -uformat "%Y%m%d-%H%M"
    # construct the temporary MessageTracking output file name
    $outTrkFile=$HTTempDir + $ServerName + "-TRACK-" + $TimeStampStartFN + "to" + $TimeStampEndFN + ".csv"
    # construct the Logparser output report file name
    $outRptFile = $HTTempDir + $ServerName + "-" + $TimeStampStartFN + "to" + $TimeStampEndFN + "-top" + $TopUsers + "-" + $TrkTarget + "Addrs.csv"
    
    # echo the settings to console
    write-host ((get-date).ToString("HH:mm:ss") + ": TimeStampStart:" + $TimeStampStart)
    write-host ((get-date).ToString("HH:mm:ss") + ": TimeStampEnd:" + $TimeStampEnd)
    write-host ((get-date).ToString("HH:mm:ss") + ": `$outTrkFile:" + $outTrkFile)
    write-host ((get-date).ToString("HH:mm:ss") + ": `$outRptFile:" + $outRptFile)

    #2. Use msgtrack to pull out the time range of receive events (make sure to use no-type):
    # events: DEFER | DELIVER | DSN | FAIL | POISONMESSAGE | RECEIVE | SEND
    # build the Message Tracking syntax as a string 
    $sExcCmd = 'get-messagetrackinglog -ResultSize Unlimited -server ' + $ServerName + ' -Start "' + $TimeStampStart + '" -End "' + $TimeStampEnd + '" |  select Timestamp,ClientIp,ClientHostname,ServerIp,ServerHostname,SourceContext,ConnectorId,Source,EventId,InternalMessageId,MessageId,@{Name="Recipient-Addrs";Expression={$_.recipients}},RecipientStatus,TotalBytes,RecipientCount,RelatedRecipientAddress,Reference,MessageSubject,Sender,ReturnPath,MessageInfo | export-csv ' + $outTrkFile + ' â??notype' 
    
    write-host -foregroundcolor Yellow "---"
    write-host -foregroundcolor Yellow ((get-date).ToString("HH:mm:ss") + ": Running initial messagetracking query(collect " + $TrkMins + "min traffic)...")
    
    # start a stopwatch to time the the process
    $sw = [Diagnostics.Stopwatch]::StartNew()
    write-host " " 
    write-host ((get-date).ToString("HH:mm:ss") + ": get-messagetracking query at Invoke:" ); $sExcCmd 
    
    # clear any existing errors
    $error.clear() 
    # Invoke the get-messagetrackinglog command to retrieve the raw traffic for the specified number of minutes
    Invoke-Expression $sExcCmd -verbose

    # check for errors
    if($error.count -gt 0){
      write-host -foregroundcolor red ((get-date).ToString("HH:mm:ss") + ": " + $error )
      if ($error -like '*The task has failed with an exception. The error message is: The parameter is incorrect*') {write-host -foregroundcolor yellow ((get-date).ToString("HH:mm:ss") + ": *** You cannot run Exch2010 MessageTracks from Exch2007 EMS.`nMove execution of this script to a machine with Exch2010 EMS installed***") ; exit 1}
    } # if-block end

    # confirm that the CSV tracking output file was produced, before moving on to logparsing
    if (!(test-path $outTrkFile)){
      write-warning ("NO " + $outTrkFile + " file written! ABORTING logparser pass on this system")
    } else {
      If ($bDebug -eq $TRUE) {write-host -ForegroundColor Yellow ((get-date).ToString("HH:mm:ss") + ": `$outTrkFile: " + $outTrkFile + " file present") }
      
      #3. Logparse out the top-20 items from the messagetracking csv output:
      
      # Switch matching the type of traffic we want to evaluate this pass; build suitable LogParser SQL query syntax to suit...
      switch ($TrkTarget) {
        "SENDER" {
          $sSQL =  "SELECT TOP 10 Sender, Count(*) AS Msgs INTO " + $outRptFile + " FROM " + $outTrkFile + " WHERE eventid = 'RECEIVE' GROUP BY Sender ORDER BY Msgs DESC"
        }
        "RECIPIENT" {
         # 12:50 PM 10/1/2013 using the new Recipient-Addrs custom field from the msgtrack
         $sSQL =  "SELECT TOP 20 Recipient-Addrs, Count(*) AS Msgs INTO " + $outRptFile + " FROM " + $outTrkFile + " WHERE eventid = 'DELIVER' OR eventid = 'SEND' OR eventid = 'FAIL' OR eventid = 'DSN' GROUP BY Recipient-Addrs ORDER BY Msgs DESC"
          
        }
        "BOUNCE" {
         # 12:50 PM 10/1/2013 using the new Recipient-Addrs custom field from the msgtrack
         $sSQL = "SELECT TOP 20 Recipient-Addrs, Count(*) AS Msgs INTO " + $outRptFile + " FROM " + $outTrkFile + " WHERE eventid = 'FAIL' OR eventid = 'DSN' GROUP BY Recipient-Addrs ORDER BY Msgs DESC"
        }
        "DEFERRED"  {
          # 12:50 PM 10/1/2013 using the new Recipient-Addrs custom field from the msgtrack
          $sSQL =  "SELECT TOP 20 Recipient-Addrs, Count(*) AS Msgs INTO " + $outRptFile + " FROM " + $outTrkFile + " WHERE eventid = 'DEFER' GROUP BY Recipient-Addrs ORDER BY Msgs DESC"
        }
      } # if-block end
       
      # build the LogParser commandline syntax, wrapping the SQL sytnax in double-quotes
      $sExcCmd =  $LOGPAREXE + ' -headerRow ON -iDQuotes APto ' + $sQot + $sSQL + $sQot + ' -i:CSV -o:csv'
      
      write-host -foregroundcolor Yellow "---"
      write-host -foregroundcolor Yellow ((get-date).ToString("HH:mm:ss") + ": Running secondary LogParser query(Top " + $TopUsers + ",on " + $TrkTarget + ")...")
      
      write-host ((get-date).ToString("HH:mm:ss") + ": " + $TrkTarget + " Query at Invoke:" ); $sExcCmd 
      
      # Invoke the LogParser command
      Invoke-Expression $sExcCmd -verbose

      # evaluate that the proper Logparser output csv file was produced (exit the script if no output)
      if (!(test-path $outRptFile)){
          write-warning ("NO " + $outRptFile + " file written!")
          exit 1
      } else {
        If ($bDebug -eq $TRUE) {write-host -ForegroundColor Yellow ((get-date).ToString("HH:mm:ss") + ": `$outRptFile: " + $outRptFile + " file present")}
      } # if-block end
      
      # if we're using CSV output of the report, pull the report csv in and redisplay it to console...
      if ($bOutputToCSV -eq $TRUE) {
        import-csv $outRptFile | ft -wrap | out-defAPlt
      } # if-block end
      
    } # if-block end (logparser block)
    
    # then move the processing files ($outTrkFile & $outRptFile csv's to the permanent reporting $DestDir
    write-host -foregroundcolor Yellow "---"
    write-host ((get-date).ToString("HH:mm:ss") + ": Pass Results..." )
    write-host ((get-date).ToString("HH:mm:ss") + ": input file: " + $outTrkFile )
    write-host ((get-date).ToString("HH:mm:ss") + ": output file: " + $outRptFile)
    write-host ((get-date).ToString("HH:mm:ss") + ": Moving output file to : " + $DestDir)
    
    # Only copy if not already running on the reporting archive server 
    if ($ComputerName -eq $runsServer) {
      write-host ((get-date).ToString("HH:mm:ss") + ": Script is executing on " + $runsServer + " skipping move of result files.")
    } else {
      # remote server pass, move results & processing files to central storage location.
        $error.clear() 
        move-item $outTrkFile $DestDir #-whatif
        move-item $outRptFile $DestDir #-whatif
    } # if-block end
    
    
    write-host -foregroundcolor green ((get-date).ToString("HH:mm:ss") + ": PASS COMPLETED")
   # stop stopwatch & echo time
    $sw.Stop()
    # simple output
    write-host -foregroundcolor green ((get-date).ToString("HH:mm:ss") + ": Elapsed Time: (HH:MM:SS.ms)" + $sw.Elapsed.ToString())
    write-host -foregroundcolor Yellow ("=" * 10)
        
  } # for-loop end HUBS

} #*----------------^ END Function ProcessLogs() ^----------------

#*================^ END FUNCTION LISTINGS ^================

#--------------------------- Invocation of SUB_MAIN ---------------------------

$EMSVers = EMSLoadLatest  # return's a string ["2013"|"2010"|"2007"]
write-host ((get-date).ToString("HH:mm:ss") + ": `$EMSVers: " + $EMSVers)
# then test $EMSVers for approp code in multi-version scripts (2007|2010|2013)

$ComputerName = ($env:COMPUTERNAME)

# assign specs per params: 
$SiteName=$SiteName.ToUpper()
$ExTargVers=$ExTargVers.ToUpper()
$TrkTarget=$TrkTarget.ToUpper()

# 12:35 PM 8/28/2013 validate the inputs (uses regular expression matching)
if (!($SiteName -match "^(US|EU|AP)$")) {
  write-warning ("INVALID SiteName SPECIFIED: " + $SiteName + ", EXITING...")
  exit 1
} # if-block end
if (!($ExTargVers -match "^(2007|2010)$")) {
  write-warning ("INVALID ExTargVers SPECIFIED: " + $ExTargVers + ", EXITING...")
  exit 1
} # if-block end
# DEFERRED | DELIVER | DSN | FAIL | POISONMESSAGE | RECEIVE | SEND
if (!($TrkTarget -match "^(SENDER|RECIPIENT|BOUNCE|DEFERRED)$")) {
  write-warning ("INVALID TrkTarget SPECIFIED: " + $TrkTarget + ", EXITING...")
  exit 1
} # if-block end
# TrkMins integer range 0-200
if (!($TrkMins -match "^([1-9]|[1-9][0-9]|[1][0-9][0-9]|20[0-0])$")) {
  write-warning ("INVALID TrkMins SPECIFIED: " + $TrkMins + ", EXITING...")
  exit 1
} # if-block end
# specify the number of 'Top ##' entries to return via logparser
$TopUsers = 20

write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ": `$ExTargVers:" + $ExTargVers)

# Switch as per specified Exchange version - test to ensure we don't try to run Ex2010 EMS cmdlets from Ex2007 EMS...
switch ($EMSVers){
  "2007" {
    if ($ExTargVers -ne "2007") {
      write-warning ("There is a mismatch between the specified Exch Rev (`$ExTargVers:" + $ExTargVers + ") and the detected EMS version (`$EMSVers:" + $EMSVers + "). `nEMS2007 can only be used for 2007 queries. `nExiting...")
     exit 1
    }
  }
  "2010" {
    write-host "2010"
  }
  "2013" {
    write-host "2010"
  }
} # switch block end

<# $TrkTarget Exchane Message Tracking EventID options:
DEFER Message delivery delayed
DELIVER Message delivered to a mailbox
DSN "A delivery status notification was generated.
Messages quarantined by the Content Filter are also delivered as DSNs. the recipients field has the SMTP address of the quarantine mailbox."
EXPAND Distribution Group expanded. The RelatedRecipientAddress field has the SMTP address of the Distribution Group.
FAIL Delivery failed. The RecipientStatus field has more information about the failure, including the SMTP response code. You should also look at the Source and Recipients fields when inspecting messages with this event.
POISONMESSAGE Message added to or removed from the poison queue
RECEIVE "Message received. The Source field is STOREDRIVER for messages submitted by Store Driver (from a Mailbox server), or  SMTP for messagesâ?¦
a) received from another Hub/Edge
b) received from an external (non-Exchange) host using SMTP
c) submitted by SMTP clients such as POP/IMAP users."
REDIRECT Message redirected to alternate recipient
RESOLVE Generally seen when a message is received on a proxy address and resolved to the defAPlt email address. The RelatedRecipientAddress field has the proxy address the message was sent to. The recipients field has the defAPlt address it was resolved (and delivered) to.
SEND Message sent by SMTP. The ServerIP and ServerHostName parameters have the IP address and hostname of the SMTP server.
SUBMIT "The Microsoft Exchange Mail Submission service on a Mailbox server successfully notified a Hub Transport server that a message is awaiting submission (to the Hub). These are the events you'll see on a Mailbox server.
The SourceContext property provides the MDB Guid, Mailbox Guid, Event sequence number, Message class, Creation timestamp, and Client type. Client type can be User (Outlook MAPI), RPCHTTP (Outlook Anwhere), OWA, EWS, EAS, Assistants, Transport."
TRANSFER Message forked because of content conversion, recipient limits, or transport agents
#>

# specify the location of temporary processing files
$HTTempDir = "e:\scripts\logs\"
# specify the UNC path to the archive location for the resulting reports
$DestDir = "\\US-MAILUTILS\e$\scripts\logs"

# flag to indicate whether we're LogParsing to CSV (used to pull the csv back in and display it to console)
$bOutputToCSV = $TRUE

$runsServer = "US-MAILUTILS"
# shift to dynamic on the local box

# Derive paths from script path support
$ScriptDir = (Split-Path -parent $MyInvocation.MyCommand.Definition) + "\"
$ScriptBaseName = [system.io.path]::GetFilename($MyInvocation.InvocationName)
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName)  


# log parser is assumed to be in the directory from which the script is being run
$LOGPAREXE = ".\logparser.exe"
# 9:20 AM 9/27/2013 test logpar loc

if (!(test-path $LOGPAREXE)) {
  if (test-path .\logparser.exe) {
    $LOGPAREXE = ".\logparser.exe"
  } else {
    write-warning ("NO COPY OF LOGPARSER FOUND AT EITHER OF:`n" + $LOGPAREXE + "`n" + (Get-Location).path)
    write-warning "Exiting"
    exit 1
  }  # if-block end
  write-host ((get-date).ToString("HH:mm:ss") + ": Creating " + $HTTempDir)
  New-Item $HTTempDir â??type directory    
} # if-block end

$TimeStampNow = get-date -uformat "%Y%m%d-%H%M" 

write-host " "

# 2:37 PM 9/25/2013 it's making it to here, but no further, just drops out

# confirm log dir exists
$testPath= $ScriptDir + "logs\" 
if (!(test-path $testPath)) {
  write-host ((get-date).ToString("HH:mm:ss") + ": Creating " + $testPath)
  New-Item $testPath -type directory    
} # if-block end

If ($bDebug -eq $TRUE) {
  write-host -foregroundcolor darkgray (((get-date).ToString("HH:mm:ss") + ": `$ScriptDir: " + $ScriptDir))
  write-host -foregroundcolor darkgray (((get-date).ToString("HH:mm:ss") + ": `$ScriptNameNoExt: " + $ScriptNameNoExt))
  write-host -foregroundcolor darkgray (((get-date).ToString("HH:mm:ss") + ": `$ScriptBaseName: " + $ScriptBaseName))
  write-host -foregroundcolor darkgray (((get-date).ToString("HH:mm:ss") + ": `$LOGPAREXE: " + $LOGPAREXE))
  write-host -foregroundcolor darkgray (((get-date).ToString("HH:mm:ss") + ": `$HTTempDir: " + $HTTempDir))
} # if-block end

# 8:49 AM 5/8/2013 check logparser.exe
if (!(test-path $LOGPAREXE)) {write-warning ("ERROR LOGPARSER NOT FOUND AT " + $LOGPAREXE + ". EXITING") ; exit 32}

# setup variables to hold Quote and Single-Quote characters
$sQot = [char]34 ; $sQotS = [char]39 ;

# description of the script's purpose
$sAppDesc = "Running msgtracking on the specified regional HTs, for the last 30mins of traffic, then running logparser to produce top-" + $TopUsers + " " + $TrkTarget + " traffic sources." + "..."
write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ": " + $sAppDesc)
# LAUNCH THE PROCESSLOGS FUNCTION
ProcessLogs -Site $SiteName -ExTargVers $ExTargVers -TrkTarget $TrkTarget -TrkMins $TrkMins

# call the generic cleanup function  
Cleanup

#*================^ END SCRIPT BODY ^================

 