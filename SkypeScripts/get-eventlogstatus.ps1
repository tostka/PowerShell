# get-EventLogStatus.ps1

<# 
.SYNOPSIS
get-EventLogStatus.ps1 - Quick server status script to confirm System Health: Displays the last 4hrs of App & Sys Err|Warn events, and then displays App & Sys last 5 Errors
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
* 9:47 AM 9/28/2017 updated ExHub example to drop a swath of FOPE engine update errors from SITE
* 12:45 PM 9/20/2017 had to create 'uniq' obj containing summary id & providernames, to run through in sample evts blocks for Sys,App & Crim logs
* 9:26 AM 9/1/2017 add timegenerated & providername filter on the event sample summaries (was pulling the first id match, from any provider)
* 7:50 AM 8/21/2017 switch get-winevents -ea from Stop to Continue
* 10:05 AM 8/4/2017 added machine name to the sample reports, looped out the App/Sys & Crimlog grouped events and pulled a single sample of each if not -Dump
upper the $Computer vari, sub'd out & with 'and' now Server0 is complaing, after changing invalid -AppLog fail code, updated Distrib script
tested and working multi server SITEc validation: .\get-EventLogStatus.ps1 -ComputerName "Server7" -Dump ; , added remote param: $ComputerName, sub'd all $env:computernames with it, 
added -dump param, SITEc common excludes in last example, debugging OK on Server0, shift ,'s back to trailing (from leading) ise on l13 boxes complaining, also swap out 
' for " in params, ran clean on L13 , 
* 10:25 AM 8/3/2017 added distrib script to pshelp
* 1:25 PM 10/26/2016 typo fix: $rrSrc => $rSrc
* 12:24 PM 10/10/2016 fixed extraneous } @ln 384
* 9:11 AM 9/30/2016 added pretest if(get-command -name set-AdServerSettings -ea 0)
* 10:09 AM 8/24/2016 added -ExcludeSource param, to supporess noisey sources, added csv export of the raw events. xml is good for powershell filtering (retains hierarchy & all
 metadata), but csv is about the only thing well-readable by excel for GUI analysis. 
* 8:51 AM 9/1/2015 save out and update to be get-EventLogStatus-4h.ps1
* 8:03 AM 5/23/2014 added transcribed output
* 7:53 AM 5/23/2014 added grouped output summary
* 9:12 AM 1/31/2014 initial build
#-=-=-UPDATE VERS DISTRIBUTION SCRIPT=-=-=-=-=-=
ALL EXCHANGE & SITEC SERVERS: 8:40 AM 8/4/2017 updated
$whatif=$true ; $files = (gci -path "\\$env:COMPUTERNAME\c$\usr\work\exch\scripts\get-eventlogstatus.*" | ?{$_.Name -match "(?i:(.*\.(PS1|CMD)))$" }) ;[array]$srvrs = get-exchangeserver | ?{(($_.IsMailboxServer) -OR ($_.IsHubTransportServer))} | select  @{Name='COMPUTER';Expression={$_.Name }} ;$srvrs +=(get-cspool| ?{($_.Services -like '*Registrar*') -AND ($_.Site -match 'SITE((_)*)SiteName((_Lab)*)')} | select -expand computers | select @{Name='COMPUTER';Expression={$_ }}) ;$srvrs = $srvrs|?{$_.computer -ne $($env:COMPUTERNAME) } ;"`$Srvrs:$($srvrs.computer -join ';')" ; $srvrs | % { write-host "$($_.computer)" ;  copy-item -path  $files -Destination \\$($_.computer)\c$\scripts\ -whatif:$($whatif) ; } ; get-date ;
#-=-=-=-=-=-=-=-=


.DESCRIPTION
.PARAMETER  Hours
Hours of Eventlogs to process (default=4)[integer]
.PARAMETER  AppLog
Name of Application & Services sub-log [-AppLog "SITEc Server"] 
.PARAMETER  ComputerName
Computer Name (defaults local) [-ComputerName server]
.PARAMETER  ExcludeEvents
Comma delimited string of EventIds to be excluded from results[string]
.PARAMETER  ExcludeSource
Comma delimited string of Sources to be excluded from results[string]
.PARAMETER  Dump
Dump outputs to console (as well as export) [-dump]
ParaHelpTxt
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.\get-EventLogStatus.ps1 -hours 24 -ExcludeEvents "12014,1000" ;  
Run 24 hour pass, excluding 2 eventids.
.EXAMPLE
.\get-EventLogStatus.ps1 -Hours 24 -ExcludeEvents "12014,1000,6020,2003" -ExcludeSource "Microsoft Forefront Protection,GetEngineFiles" ;  
Run 24 hour pass, excluding 4 eventids, and 2 sources.
.EXAMPLE
.\get-EventLogStatus.ps1 -Hours 3 -AppLog 'SITEc Server'
Run 3 hour pass on SITEc Crimson log.
.EXAMPLE
.\get-EventLogStatus.ps1 -Hours 3 -AppLog 'SITEc Server' -ExcludeEvents "14507" ; 
Run 3hr pass on SITEc Server Crimson log, and exclude eventid 14507 events
.EXAMPLE
.\get-EventLogStatus.ps1 -Hours 3 -AppLog 'SITEc Server' ; 
Run 3hr pass on both the 'SITEc Server' Crimson log
.EXAMPLE
.\get-EventLogStatus.ps1 -ComputerName "Server7" -Dump ; 
For 3 SITEc FE's run 24hr pass on the 'SITEc Server' Crimson log dump matching events to console, Exclude 36888,44028 & 4098's
.EXAMPLE
"Server4 ; } ;
Run a list of hosts 24hr
.EXAMPLE
get-exchangeserver |?{$_.ishubtransportserver -AND $_.site -like '*SiteName'} |%{ "===$($_):" ; .\get-EventLogStatus.ps1 -computername $_ -Hours 2  -ExcludeEvents "12014,36887,36888,6020,6012,7004,7001" ; } ;
Run dynamic Exchange hub query, 4hr filter with common HubCas eventid exclusions (including FOPE)
.EXAMPLE
cls ; get-exchangeserver |?{$_.IsMailboxServer -AND $_.site -like '*SiteName'} |%{ "===$($_):" ; .\get-EventLogStatus.ps1 -computername $_ -Hours 2   ; } ;
Run dynamic Exchange hub query, 4hr filter with common HubCas eventid exclusions
.LINK
#>

# 2:56 PM 8/3/2017 untype AppLog, see if it addresses fail on 2nd applog loop: [string]$AppLog
# 7:32 AM 8/4/2017 shift ,'s back to trailing (from leading) ise on l13 boxes complaining, also swap out ' for " in params
# 8:03 AM 8/4/2017 added remote option param for $ComputerName, sub'd out all $env:computernames with it, 
Param(
    [Parameter(Position=0,HelpMessage="Hours of Eventlogs to process (default=4)[integer]")][ValidateNotNullOrEmpty()]
    [alias("hrs")]
    [int]$Hours=4,
    [Parameter(Position=1,Mandatory=$false,HelpMessage="Name of Application & Services sub-log [-AppLog 'SITEc Server']")]
    [alias("Src")]
    $AppLog,
    [Parameter(Mandatory=$false,HelpMessage='Computer Name [-ComputerName server]',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [Alias('__ServerName', 'Server', 'Computer', 'Name')]
    [string[]]$ComputerName = $env:COMPUTERNAME, 
    [Parameter(Mandatory=$false,HelpMessage="Comma delimited string of EventIds to be excluded from results[string]")]
    [alias("exclude")]
    [string]$ExcludeEvents,
    [Parameter(Mandatory=$false,HelpMessage="Comma delimited string of Sources to be excluded from results[string]")]
    [alias("NoSrc")]
    [string]$ExcludeSource,
    [Parameter(Mandatory=$false,HelpMessage="Dump outputs to console (as well as export) [-dump]")]
    [switch]$Dump
) # PARAM BLOCK END



#region INIT; # ------ 
# vsc dump the $args parameter
if ($Args.count -ge 1){
    write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):`$args:" ;    
    foreach ($i in $args) {$i}
} ; 

#*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
# pick up the bDebug from the $ShowDebug switch parameter
# SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
if ($ShowDebug) {$bDebug=$true ; $DebugPreference = "Continue" ; write-debug "(`$ShowDebug:$ShowDebug ;`$bDebug:$bDebug ;`$DebugPreference:$DebugPreference)" ; };
if ($Whatif){$bWhatif=$true ; Write-Verbose -Verbose:$true "`$Whatif is $true (`$bWhatif:$bWhatif)" ; };
if($bdebug){$ErrorActionPreference = 'Stop' ; write-debug "(Setting `$ErrorActionPreference:$ErrorActionPreference;"};
# scriptname with extension
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ; 
#$ComputerName = $env:COMPUTERNAME ; # 8:03 AM 8/4/2017 rem'd favor of param defaulted here
$sQot = [char]34 ; $sQotS = [char]39 ; 
$NoProf=[bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
$MyBox="MyComputer","MyComputer" ; 
$DomainWork = "DOMAIN";
$DomHome = "DOMAIN";
$DomLab="DOMAIN";
#$ProgInterval= 500 ; # write-progress wait interval in ms
# 12:23 PM 2/20/2015 add gui vb prompt support
#[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null ; 
# 11:00 AM 3/19/2015 should use Windows.Forms where possible, more stable

# Clear error variable
$Error.Clear() ; 
<##-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# SCRIPT-CLOSE MATERIAL TO CLEAR THE UNDERLYING $DBGPREF & $EAPREF TO DEFAULTS:
if ($ShowDebug -OR ($DebugPreference = "Continue")) {
        Write-Verbose -Verbose:$true "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
        $bDebug=$false
        # 8:41 AM 10/13/2015 also need to enable write-debug output (and turn this off at end of script, it's a global, normally SilentlyContinue)
        $DebugPreference = "SilentlyContinue" ;
} # if-E ; 
if($ErrorActionPreference -eq 'Stop') {$ErrorActionPreference = 'Continue' ; write-debug "(Restoring `$ErrorActionPreference:$ErrorActionPreference;"};
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#>
#*======^ SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE ^======
#endregion INIT; # ------ 

## immed after param, setup to stop executing the script on the first error
#trap { break; }


#$bDebug = $TRUE
$bDebug = $FALSE
If ($bDebug -eq $TRUE) {write-host "*** DEBUGGING MODE ***"}




#*================v FUNCTION DECLARATIONS v================


#---------------------------v Function Cleanup v---------------------------
Function Cleanup {
    # clear all objects and exit
    # Clear-item doesn't seem to work as a variable release 
    # vers: 12:35 PM 1/31/2014 - added test-transcribing
      
    Write-Host ((get-date).ToString("HH:mm:ss") + "Exiting Script...")
    # 7:43 AM 1/24/2014 always stop the running transcript before exiting
    if(Test-TranscriptionSupported){
        if (Test-Transcribing) {stop-transcript} ;
    } ; 
  
  exit
} #*----------------^ END Function Cleanup ^----------------

#*------v Function Test-TranscriptionSupported v------
function Test-TranscriptionSupported {
  <#
  .Synopsis
  Tests to see if the current host supports transcription.
  .Description
  Powershell.exe supports transcription, WinRM and ISE do not.
  .Example
  #inside powershell.exe
  Test-Transcription
  #returns true
  Description
  ------
  Returns a $true if the host supports transcription; $false otherwise
  #>
  
  #($Host.Name -eq 'ServerRemoteHost')
  $externalHost = $host.gettype().getproperty("ExternalHost",
  [reflection.bindingflags]"NonPublic,Instance").getvalue($host, @())
  try {
    [Void]$externalHost.gettype().getproperty("IsTranscribing",
    [Reflection.BindingFlags]"NonPublic,Instance").getvalue($externalHost, @())
    $true
  } catch {
    $false
  } # try-E
}#*------^ END Function Test-TranscriptionSupported ^------


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

#*----------------v SUB MAIN  v----------------

# the ever-current version not used, file locks prevent overwrites
#$staticoutfile = "store-mbx-counts-current.csv"

if (!(test-path ($ScriptDir + "logs"))) {
  write-host "Creating " $($ScriptDir + "logs")
  New-Item ($ScriptDir + "logs\") -type directory    
}; 
# 8:13 AM 8/4/2017 shifting to possible array $Computername, means get-winevent won't accept an array, so we need to always forloop it
foreach ($Computer in $ComputerName){
    $Computer = $Computer.toUpper() ; 
    $TimeStampNow = get-date -uformat "%Y%m%d-%H%M" ;
    #$outransfile=$ScriptDir + "logs\" + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.log" ;
    $outransfile=$ScriptDir + "logs\$($ScriptNameNoExt)-$($Computer)-$($Hours)HRS-$($TimeStampNow)-trans.log" ;

    # 11:08 AM 8/3/2017: echo params
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):#*======v PROCESSING: $($Computer) v======" ; 

    if($Computer){write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$Computer:$($Computer)" ; } ; 
    if($Hours){write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$Hours:$($Hours)" ; } ; 
    if($AppLog){write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$AppLog:$($AppLog)" ; } ; 
    if($ExcludeEvents){write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$ExcludeEvents:$($ExcludeEvents)" ; } ; 
    if($ExcludeSource){write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$ExcludeSource:$($ExcludeSource)" ; } ; 


    $ErrorActionPreference = 'Stop' ; 
    # Clear error variable
    $Error.Clear() ; 
    # try/catch/finally pattern
    TRY {
        if($ExcludeEvents){
            # split the string into an array
            $RmvEvts=$ExcludeEvents.split(",")
        } ; 
        # 10:00 AM 8/24/2016 $ExcludeSource
        if($ExcludeSource){
            $RmvSrcs=$ExcludeSource.split(",")
        } ;     
        # 12:56 PM 10/10/2016 add & chck AppLog
        if($AppLog){
            if( !(Get-EventLog -list |?{$_.Log -eq $AppLog})){ 
              write-host -foregroundcolor red "$((get-date).ToString('HH:mm:ss')):NO LOCAL SOURCE:$($APPLOG) FOUND, ABORTING PASS";
              throw "NO LOCAL SOURCE:$($APPLOG) FOUND, ABORTING PASS" ; 
            } ; 
        } ; 

        if(Test-TranscriptionSupported){
            if (Test-Transcribing) {stop-transcript} ;
            write-host -foreground yellow ("Transcribing output to: " + $outransfile)
            start-transcript -path $outransfile ;
        } else { 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Note: This host DOES NOT support transcription! Skipping Tscript file" ;
        }

        # run out the eventlog summary
        write-host -fore yellow ("=" * 10);
    
        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):`n=== $($Hours) HRS===" ;
    
        $tLog="Application"  ; 
        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Processing $($tLog) Error and Warning Evts" ;
    
        # 11:45 AM 8/3/2017 levels: Warning = 3 Error = 2 Critical = 1 
        # 8:19 AM 8/4/2017 add ea 4 stop
        # 7:44 AM 8/21/2017 add -ea 0, throws error if no matches: NoMatchingEventsFound, which fails try/catch
        $AppEvts=Get-WinEvent -computername $Computer -FilterHashtable @{logname="$($tLog)"; Level=1,2,3 ;  StartTime=$(([DateTime]::Now.AddHours(-1 * $Hours))) ;} -ErrorAction SilentlyContinue   | 
            select TimeCreated,Level,LevelDisplayName ,ProviderName,Id,Message ;  

        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Raw Evts Returned:$(($AppEvts |measure).count)" ;
        if($RmvEvts){
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Excluding following Events: $($ExcludeEvents)" ;
            foreach ($rEvt in $RmvEvts){
                #$AppEvts = $AppEvts |?{$_.EventID -ne $rEvt} ; 
                # TimeCreated,Level,LevelDisplayName ,ProviderName,Id,Message
                $AppEvts = $AppEvts |?{$_.Id -ne $rEvt} ; 
            } ; 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Net $($tLog) Evts:$(($AppEvts |measure).count)" ;
        } ; 
    
        # $RmvSrcs
        if($RmvSrcs){
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Excluding following Sources: $($ExcludeSource)" ;
            foreach ($rSrc in $RmvSrcs){
                #$AppEvts = $AppEvts |?{$_.Source -ne $rSrc} ; 
                $AppEvts = $AppEvts |?{$_.ProviderName -ne $rSrc} ; 
            } ; 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Net $($tLog) Evts:$(($AppEvts |measure).count)" ;
        } ; 

        $aOutFile=$outransfile.replace("-trans","-$($tLog.substring(0,3))").replace(".log",".xml") ; 
        if($AppEvts.count -gt 0){
            if($Dump){
                write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):#*======v Dumping $($tlog) Events to Console v======" ;
                $AppEvts | out-string | ft -auto | out-default ; 
                write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):#*======^ Dumping $($tlog) Events to Console ^======" ;
            } ; 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Exporting $($tlog) Events to $($aOutFile)" ;
            $AppEvts | export-clixml -path $aOutFile ; 
            # 9:47 AM 8/24/2016export a raw csv for excel use
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Exporting $($tlog) Events to $($aOutFile.replace(".xml",".csv"))" ;
            $AppEvts | export-csv -notype -path $aOutFile.replace(".xml",".csv") ; 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Generating EventID grouping for $($tLog) events:" ;
            $AppEvtsGrpd = $AppEvts| group ID | sort Count -desc | select Count, @{Name='ID';Expression={$_.Name}}  ; 
            $AppEvtsGrpd | out-string | ft -auto | out-default ; 
            $AppEvtsGrpd | export-csv -notype -path $aOutFile.replace("$($tLog.substring(0,3))","$($tLog.substring(0,3))-GroupedEvents").replace(".xml",".csv") ; 
            # 12:39 PM 9/20/2017 need a separate data obj for unique id & providername (the grouped only puts out the ids)
            $AppEvtsUniqs = $appevts | select -COque providername,id ; 
            # 9:24 AM 8/4/2017 if not dumping, pull group samples
            if(!$Dump){
                "#*======v SAMPLE $($Computer) $($tLog) EVENTS v======" ; 
                foreach ($item in $AppEvtsUniqs){
                    #12:41 PM 9/20/2017 updated to $item.Providername
                    write-host "$((Get-WinEvent -computername $Computer -FilterHashtable @{logname="$($tLog)"; ID=$($item.ID) ; ProviderName=$($item.Providername) ; StartTime=$(([DateTime]::Now.AddHours(-1 * $Hours))) ;} -ErrorAction SilentlyContinue -MaxEvents 1 | fl ID,LevelDisplayName ,ProviderName,Message,TimeGenerated | out-string).trim())`n-----" ; 
                } ; 
                "#*======^ SAMPLE $($Computer) $($tLog) EVENTS ^======" ; 
            } ; 
        } else { 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):No matching $($tlog) Events: Creating placemarker export file:$($aOutFile)" ;
            $dummy=@{Results="(NO MATCHES RETURNED)"};
            New-Object PSObject -Property $dummy | export-clixml -path $aOutFile ;
        } ; 
    
    
        # -----------
    
    
        $tLog="System"  ; 
        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Processing $($tLog) Error and Warning Evts" ;
        <#$SysEvts =Get-Eventlog -after ([DateTime]::Now.AddHours(-1 * $Hours)) -logname $tLog | ?{$_.EntryType -eq 'Error' -OR $_.EntryType -eq 'Warning'} | 
            select TimeGenerated,EntryType,Source,EventID,Message ; 
        #>
        $SysEvts=Get-WinEvent -computername $Computer -FilterHashtable @{logname="$($tLog)"; Level=1,2,3 ;  StartTime=$(([DateTime]::Now.AddHours(-1 * $Hours))) ;} -ErrorAction Continue  | 
            select TimeCreated,Level,LevelDisplayName ,ProviderName,Id,Message ; 

        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Raw Evts Returned:$(($SysEvts |measure).count)" ;
        if($RmvEvts){
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Excluding following Events: $($ExcludeEvents)" ;
            foreach ($rEvt in $RmvEvts){
                #$SysEvts = $SysEvts |?{$_.EventID -ne $rEvt} ; 
                $SysEvts = $SysEvts |?{$_.ID -ne $rEvt} ; 
            } ; 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Net $($tLog) Evts:$(($SysEvts |measure).count)" ;
        } ; 
        # $RmvSrcs 10:03 AM 8/24/2016
        if($RmvSrcs){
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Excluding following Sources: $($ExcludeSource)" ;
            foreach ($rSrc in $RmvSrcs){
                #$SysEvts = $SysEvts |?{$_.Source -ne $rSrc} ; 
                $SysEvts = $SysEvts |?{$_.ProviderName -ne $rSrc} ; 
            } ; 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Net $($tLog) Evts:$(($SysEvts |measure).count)" ;
        } ;
    
        $sOutFile=$outransfile.replace("-trans","-$($tLog.substring(0,3))").replace(".log",".xml") ; 
        if($SysEvts.count -gt 0){
            # 7:47 AM 8/4/2017 add console dump
            if($Dump){
                write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):#*======v Dumping $($tlog) Events to Console v======" ;
                $SysEvts | out-string | ft -auto | out-default ; 
                write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):#*======^ Dumping $($tlog) Events to Console ^======" ;
            } ; 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Exporting $($tlog) Events to $($sOutFile)" ;
            $SysEvts | export-clixml -path $sOutFile ; 
            # 9:47 AM 8/24/2016export a raw csv for excel use
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Exporting $($tlog) Events to $($sOutFile.replace(".xml",".csv"))" ;
            $SysEvts | export-csv -notype -path $sOutFile.replace(".xml",".csv") ; 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Generating EventID grouping for $($tLog) events:" ;
            $SysEvtsGrpd = $SysEvts| group ID | sort Count -desc | select Count, @{Name='ID';Expression={$_.Name}}  ; 
            $SysEvtsGrpd | out-string | ft -auto | out-default ; 
            $SysEvtsGrpd | export-csv -notype -path $sOutFile.replace("$($tLog.substring(0,3))","$($tLog.substring(0,3))-GroupedEvents") ; 
            # 12:39 PM 9/20/2017 need a separate data obj for unique id & providername (the grouped only puts out the ids)
            $SysEvtsUniqs = $SysEvts | select -COque providername,id ; 

            # 9:24 AM 8/4/2017 if not dumping, pull group samples
            if(!$Dump){
                "#*======v SAMPLE $($Computer) $($tLog) EVENTS v======" ; 
                foreach ($item in $SysEvtsUniqs){
                    # 12:31 PM 9/20/2017 typo fix, mising $ on $ProviderName
                 write-host "$((Get-WinEvent -computername $Computer -FilterHashtable @{logname="$($tLog)"; ID=$($item.ID) ; ProviderName=$($item.ProviderName) ; StartTime=$(([DateTime]::Now.AddHours(-1 * $Hours))) ;} -ErrorAction SilentlyContinue -MaxEvents 1 | fl ID,LevelDisplayName ,ProviderName,Message,TimeGenerated | out-string).trim())`n-----" ; 
                } ; 
                "#*======^ SAMPLE $($Computer) $($tLog) EVENTS ^======" ; 
            } ; 
        } else { 
            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):No matching $($tlog) Events: Creating placemarker export file:$($sOutFile)" ;
            $dummy=@{Results="(NO MATCHES RETURNED)"};
            New-Object PSObject -Property $dummy | export-clixml -path $sOutFile ;
        } ; 
    
        # 1:00 PM 10/10/2016 add Crimson log pass using -AppLog
        if($AppLog){
            if (get-command get-winevent){
                $CrimLog = $AppLog ; 
                $AppEvents=@{} ; 

                    $tLog=$CrimLog  ; 
                    write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Processing CrimsonLog '$($CrimLog)' Error and Warning Evts" ;
                    $Returns = Get-WinEvent -computername $Computer -FilterHashtable @{logname="$($CrimLog)"; Level=1,2,3 ;  StartTime=$(([DateTime]::Now.AddHours(-1 * $Hours))) ;} -ErrorAction Continue  | 
                        select TimeCreated,Level,LevelDisplayName ,ProviderName,Id,Message ; 
                    if($Returns){
                        $AppEvents.add("$($CrimLog)",$Returns) ; 
                    }; 

                    write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Raw Evts Returned:$(($AppEvents["$($CrimLog)"] |measure).count)" ;
                    if($RmvEvts){
                        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Excluding following Events: $($ExcludeEvents)" ;
                        foreach ($rEvt in $RmvEvts){
                            #$AppEvents["$($CrimLog)"] = $AppEvents["$($CrimLog)"] |?{$_.EventID -ne $rEvt} ; 
                            #$AppEvents["$($CrimLog)"] = $AppEvents["$($CrimLog)"] |?{$_.ID -ne $rEvt} ;
                            # $h."$($_.Name)" = $_.Value 
                            $AppEvents."$($CrimLog)" = $AppEvents."$($CrimLog)" |?{$_.ID -ne $rEvt} ; 
                        } ; 
                        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Net $($tLog) Evts:$(($AppEvents."$($CrimLog)" |measure).count)" ;
                    } ; 
                    # $RmvSrcs 10:03 AM 8/24/2016
                    if($RmvSrcs){
                        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Excluding following Sources: $($ExcludeSource)" ;
                        foreach ($rSrc in $RmvSrcs){
                            #$AppEvents["$($CrimLog)"] = $AppEvents["$($CrimLog)"] |?{$_.Source -ne $rSrc} ; 
                            #$AppEvents["$($CrimLog)"] = $AppEvents["$($CrimLog)"] |?{$_.ProviderName -ne $rSrc} ; 
                            $AppEvents."$($CrimLog)" = $AppEvents."$($CrimLog)"  |?{$_.ProviderName -ne $rSrc} ; 
                        } ; 
                        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Net $($CrimLog) Evts:$(($AppEvents."$($CrimLog)" |measure).count)" ;
                    } ;

                    $sOutFile=$outransfile.split(".")[0].replace("trans",$CrimLog.replace(" ","")) ;
                    $sOutFile+=".xml" ; 

                    if($AppEvents."$($CrimLog)".count -gt 0){
                        # 7:47 AM 8/4/2017 add console dump
                        if($Dump){
                            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):#*======v Dumping $($tlog) Events to Console v======" ;
                            $AppEvents."$($CrimLog)"| out-string | ft -auto | out-default ; 
                            write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):#*======^ Dumping $($tlog) Events to Console ^======" ;
                        } ; 
                        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Exporting $($CrimLog) Events to $($sOutFile)" ;
                        $AppEvents."$($CrimLog)" | export-clixml -path $sOutFile ; 
                        # 9:47 AM 8/24/2016export a raw csv for excel use
                        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Exporting $($CrimLog) Events to $($sOutFile.replace(".xml",".csv"))" ;
                        $AppEvents."$($CrimLog)" | export-csv -notype -path $sOutFile.replace(".xml",".csv") ; 
                        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Generating EventID grouping for $($CrimLog) events:" ;
                        #$SysEvtsGrpd["$($CrimLog)"] = $AppEvents."$($CrimLog)"| group EventID | sort Count -desc | select Count, @{Name='EventID';Expression={$_.Name}}  ; 
                        $CrimEvtsGrpd = $AppEvents."$($CrimLog)"| group ID | sort Count -desc | select Count, @{Name='ID';Expression={$_.Name}}  ; 
                        $CrimEvtsGrpd| out-string | ft -auto | out-default ; 
                        $CrimEvtsGrpd | export-csv -notype -path $sOutFile.replace("$($CrimLog.substring(0,3))","$($CrimLog.substring(0,3))-GroupedEvents") ; 
                        # 12:39 PM 9/20/2017 need a separate data obj for unique id & providername (the grouped only puts out the ids)
                        $CrimEvtsUniqs = $appevts | select -COque providername,id ; 
                        
                        # 9:24 AM 8/4/2017 if not dumping, pull group samples
                        if(!$Dump){
                            "#*======v SAMPLE $($Computer) $($tLog) EVENTS v======" ; 
                            foreach ($item in $CrimEvtsUniqs){
                                write-host "$((Get-WinEvent -computername $Computer -FilterHashtable @{logname="$($tLog)"; ID=$($item.ID) ; ProviderName=$($item.ProviderName) ; StartTime=$(([DateTime]::Now.AddHours(-1 * $Hours))) ;} -ErrorAction SilentlyContinue -MaxEvents 1 | fl ID,LevelDisplayName ,ProviderName,Message,TimeGenerated | out-string).trim())`n-----" ; 
                            } ; 
                            "#*======^ SAMPLE $($Computer) $($tLog) EVENTS ^======" ; 
                        } ; 
                    } else { 
                        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):No matching $($CrimLog) Events: Creating placemarker export file:$($sOutFile)" ;
                        $dummy=@{Results="(NO MATCHES RETURNED)"};
                        New-Object PSObject -Property $dummy | export-clixml -path $sOutFile ;
                    } ; 

                    # still getting strange failure completing loop, manually break out before last iteration
                    #If(($CrimArray | measure).count -eq 1){ break }

                } else {
                    throw "The underlying Get-WinEvent command is not present locally (requires Psv2+)" ; 
                } ; # if get-winevent
                # ====
            } ; # if applog



        write-host -fore yellow ("=" * 10);




        #stop-transcript ;
        write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):`$outransfile:$outransfile (opening...)";
        #invoke-item $outransfile ;
        #Cleanup # 8:23 AM 8/4/2017 move cleanup outside of computer loop
    } 
    CATCH {
        $msg=": Error Details: $($_)" ;
        Write-Error "$(get-date -format "HH:mm:ss"): FAILURE!" ;
        Write-Error "$(get-date -format "HH:mm:ss"): Error in $($_.InvocationInfo.ScriptName)." ; 
        Write-Error "$(get-date -format "HH:mm:ss"): -- Error information" ;
        Write-Error "$(get-date -format "HH:mm:ss"): Line Number: $($_.InvocationInfo.ScriptLineNumber)" ;
        Write-Error "$(get-date -format "HH:mm:ss"): Offset: $($_.InvocationInfo.OffsetInLine)" ;
        Write-Error "$(get-date -format "HH:mm:ss"): Command: $($_.InvocationInfo.MyCommand)" ;
        Write-Error "$(get-date -format "HH:mm:ss"): Line: $($_.InvocationInfo.Line)" ;
        #Write-Error "$(get-date -format "HH:mm:ss"): Error Details: $($_)" ;
        $msg=": Error Details: $($_)" ;
        Write-Error  "$(get-date -format "HH:mm:ss"): $($msg)"  ; 
        #if(get-command -name set-AdServerSettings -ea 0){ set-AdServerSettings -ViewEntireForest $false ;
        Exit ;
    } ; 
    <# 8:25 AM 8/4/2017 move transcript code outside of finally 
    Finally {
	    Stop-Transcript ; 
    } ; 
    #>
    if (Test-Transcribing) {stop-transcript} ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):#*======^ END PROCESSING: $($Computer) ^======" ; 

} # $Computer loop-E ; 

Cleanup 

#*----------------^ END SUB_MAIN  ^----------------

