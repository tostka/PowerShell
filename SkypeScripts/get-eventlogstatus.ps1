# get-EventLogStatus.ps1

#*----------V Comment-based Help (leave blank line below) V---------- 

## 
#     .SYNOPSIS
# get-EventLogStatus.ps1 - Quick server status script to confirm System Health: Displays the last 4hrs of App & Sys Err|Warn events, and then displays App & Sys last 5 Errors
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
# # Logs tlog content to standard log file
# 
# # AUTHOR: Todd Kadrie
# # 
# # Script to collect tlog gen counts per specified server and log the count to a local text file
# # queries all sg's on the specified server, creates an output .log file per sg (or db on ex2010)
# 
# 
# Change Log
#   8:03 AM 5/23/2014 added transcribed output
# # 7:53 AM 5/23/2014 added grouped output summary
#   # vers: 9:12 AM 1/31/2014 initial build
# 

#$bDebug = $TRUE
$bDebug = $FALSE
If ($bDebug -eq $TRUE) {write-host "*** DEBUGGING MODE ***"}


#*================v FUNCTION DECLARATIONS v================

#*----------v Function EMSLoadLatest v----------
  #Checks local machine for registred E20[10|07] EMS, and then loads the newest one found
  #Returns the string 2010|2007 for reuse for version-specific code

function EMSLoadLatest {
  # check registred vS loaded ;
  $SnapsReg=Get-PSSnapin -Registered ;
  $SnapsLoad=Get-PSSnapin ;
  # check/load E2013, E2010, or E2007, stop at newest (servers wouldn't be running multi-versions)
  if (($SnapsReg |
      where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
  {
    if (!($SnapsLoad |
      where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
    {
      Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue ; 
      return "2010" ;
    } else {
      return "2010" ;
    }
  } elseif (($SnapsReg |
      where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"}))
  {
    if (!($SnapsLoad |
      where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"}))
    {
      Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue ; 
      return "2007" ;
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
  # vers: 12:35 PM 1/31/2014 - added test-transcribing
  
  Write-Host ((get-date).ToString("HH:mm:ss") + "Exiting Script...")
  # 7:43 AM 1/24/2014 always stop the running transcript before exiting
  if (Test-Transcribing) {stop-transcript} ;
  exit
} #*----------------^ END Function Cleanup ^----------------

#*----------------v Function Test-Transcribing v----------------
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
} #*----------------^ END Function Test-Transcribing ^----------------

#*================^ FUNCTION DECLARATIONS ^================

#*----------------v SUB_MAIN  v----------------

$EMSVers = EMSLoadLatest  # returns: "2013"|"2010"|"2007"
write-host "`$EMSVers: " $EMSVers
# then test $EMSVers for approp code in multi-version scripts

# the ever-current version not used, file locks prevent overwrites
#$staticoutfile = "store-mbx-counts-current.csv"
#1:37 PM 10/25/2013 incl
#$ScriptDir = (Split-Path -parent $MyInvocation.MyCommand.Definition) + "\"
$ScriptBaseName = [system.io.path]::GetFilename($MyInvocation.InvocationName)
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName)  

$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName)

# 8:35 AM 1/28/2014 add transcript to find out why it's never completing/exiting:
# gen filename from script and start/stop
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
#$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;


if (!(test-path ($ScriptDir + "logs"))) {
  write-host "Creating " $($ScriptDir + "logs")
  New-Item ($ScriptDir + "logs\") -type directory    
}
$TimeStampNow = get-date -uformat "%Y%m%d-%H%M" ;
$outransfile=$ScriptDir + "logs\" + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.log" ;
#$outransfile=$ScriptDir + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.txt" ;
#stop-transcript -ErrorAction SilentlyContinue
# stop transcript,trap any error & eat complaint
# note, this will suppress all errors coming out of the transcript commands - even one's you WANT to see:
# 9:12 AM 1/28/2014 not working and crashes out the script!
#Trap {Continue} Stop-Transcript | Out-Null ;
if (Test-Transcribing) {stop-transcript} ;
write-host -foreground yellow ("Transcribing output to: " + $outransfile)
start-transcript -path $outransfile ;

# run out the eventlog summary
write-host -fore yellow ("=" * 10); "Application","System" | %{write-host -fore yellow "`n===$_ 4HRS===" ; Get-Eventlog -after ([DateTime]::Now.AddHours(-4)) -logname $_ | ?{$_.EntryType -eq 'Error' -OR $_.EntryType -eq 'Warning'} | select TimeGenerated,EntryType,Source,EventID,Message} ; "Application","System" | %{write-host -fore yellow "`n===$_ LAST 5 ERRORS===" ; Get-EventLog -LogName $_ -EntryType Error -Newest 5 | select TimeGenerated,EntryType,Source,Message | ft -auto | out-default} ; write-host -fore yellow ("=" * 10);

<# grouped
write-host -fore yellow ("=" * 10); "Application","System" | %{write-host -fore yellow "`n===$_ 4HRS===" ; Get-Eventlog -after ([DateTime]::Now.AddHours(-4)) -logname $_ | ?{$_.EntryType -eq 'Error' -OR $_.EntryType -eq 'Warning'} | select TimeGenerated,EntryType,Source,EventID,Message} ; "Application","System" | %{write-host -fore yellow "`n===$_ LAST 5 ERRORS===" ; Get-EventLog -LogName $_ -EntryType Error -Newest 5 | select TimeGenerated,EntryType,Source,Message | ft -auto | out-default} ; write-host -fore yellow ("=" * 10);
#>
# 7:51 AM 5/23/2014 added grouped id's output:
write-host -fore yellow ("=" * 10); "Application","System" | %{write-host -fore yellow "`n===$_ 4HRS GROUPED EVENTID's===" ; Get-Eventlog -after ([DateTime]::Now.AddHours(-4)) -logname $_ | ?{$_.EntryType -eq 'Error' -OR $_.EntryType -eq 'Warning'} | group EventID | sort Count -desc | select Count,Name | ft -auto | out-default} ;

stop-transcript ;


Cleanup

#*----------------^ END SUB_MAIN  ^----------------
