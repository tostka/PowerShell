#Check-DriveStatusHP.ps1

<#
.SYNOPSIS
Check-DriveStatusHP.ps1 - Leverage HP Array Configuration Utility CLI to output drive status info (and drive mappings). Emails regional admins for SITED/SITE, and SITE Messaging for DAG-NA
.NOTES
Written By: Todd Kadrie
Website:	http://toddomation.com
Twitter:	http://twitter.com/tostka
CheckSmartArray() By: ramazancan (original concept to parse the hpacucli.exe output for Fail's)   
Website:	https://ramazancan.wordpress.com/2014/07/04/powershellhow-to-monitor-hp-smartarray-disk-status/
Updated By: Todd Kadrie
Website:	http://toddomation.com
Twitter:	http://twitter.com/tostka
Change Log
# 4:15 PM 3/24/2018 updated pshhelp
* 7:40 AM 10/31/2017 updated pshelp desc citing specif notification addresses (regional direct)
* 8:24 AM 8/29/2017: Updated pshelp to reflect only the checksmartarray() & concept is rmazancan's code. Rest is all rewritten (part of validating for SITE use). So far working OK in SITED
* 12:16 PM 8/25/2017 successfully dupe suppresses, had to decrement the 'everyxcycles' by 1, as the 0 cycle doesn't get counted (so 4=3 as a variable) ; added regional notfication DLs and put the distr admins in the loop on the notices - saves time to reso when they're notified first. 
Also added copy status/mount confirmation on the problem db, and the edbfilepath to the problem db. ; start porting in semaphore (suppress alerts ever x hrs), and pull replica health 
-> aim is to make this deliverable to the admins on the ground and Server's, to get the drive replacement started IMMEDIATELY! without my invovlement
* # 12:00 PM 8/23/2017 failed drive on bcc641 let me live debug it, added my own smtp code, dyn smtpserver (via get-exchangeserver), works 
* # 11:07 AM 5/1/2017 put the hostname in the $LOGFILE
* 10:48 AM 5/1/2017 tshot - pulled the automail - I think hobbit/solarwinds 
    have that down, I mainly want this for the port/slot/bay specs for recommending 
    proactive drive replacements 
* 10:12 AM 5/1/2017 tweaked and extended example code
* 2014/07/04 posted version

#-=TASK WRAPPERLESS CREATION COMMAND-=-=-=-=-=-=-=
# Run every five minutes from the specified start time to end time: /SC MINUTE /MO 5 /ST 12:00 /ET 14:00
schtasks /CREATE /TN "Check-DriveStatusHP" /TR "powershell.exe -noprofile -executionpolicy Unrestricted -file c:\scripts\Check-DriveStatusHP.ps1" /SC MINUTE /MO 60 /ST 12:15 /ET 23:59 /ru "$($env:userdomain)\MyAccount" /rp "<ru pw>" /s Server0 
#-=-=-=-=-=-=-=-=
#-=-=-=-=-=-=-=-=
TO RUN TASK -WHATIF FOR TESTING, YOU NEED TO USE -COMMAND vs -FILE:
Program/Script: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
Add argument (optional): -noprofile -executionpolicy Unrestricted -file c:\scripts\Check-DriveStatusHP.ps1

#-=DISTRIBUTION CMD:-=-=-=-=-=-=-=
[array]$files = (gci -path "\\$env:COMPUTERNAME\c$\usr\work\exch\scripts\Check-DriveStatusHP*" | ?{$_.Name -match "(?i:(.*\.(PS1|CMD|XML)))$" })  ;[array]$srvrs = get-exchangeserver | ?{(($_.IsMailboxServer))} | select  @{Name='COMPUTER';Expression={$_.Name }} ;$srvrs = $srvrs|?{$_.computer -ne $($env:COMPUTERNAME) } ; $srvrs | % { write-host "$($_.computer)" ; copy $files -Destination \\$($_.computer)\c$\scripts\ -whatif ; } ; get-date ;
#-=-=-=-=-=-=-=-=
.DESCRIPTION
Check-DriveStatusHP.ps1 - Leverage HP Array Configuration Utility CLI to output drive status info (and drive mappings)
9:56 AM 5/1/2017: Server6)\Compaq\Hpacucli\Bin\hpacucli.exe"
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.EXAMPLE
.LINK
#>


#region INIT; # ------ 
#*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
# pick up the bDebug from the $ShowDebug switch parameter
# SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
if ($ShowDebug) {$bDebug=$true ; $DebugPreference = "Continue" ; write-debug "(`$ShowDebug:$ShowDebug ;`$bDebug:$bDebug ;`$DebugPreference:$DebugPreference)" ; };
if ($Whatif){$bWhatif=$true ; Write-Verbose -Verbose:$true "`$Whatif is $true (`$bWhatif:$bWhatif)" ; };
if($bdebug){$ErrorActionPreference = 'Stop' ; write-debug "(Setting `$ErrorActionPreference:$ErrorActionPreference;"};
# If using WMI calls, push any cred into WMI:
#if ($Credential -ne $Null) {$WmiParameters.Credential = $Credential }  ; 

# scriptname with extension
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ; 
$ComputerName = $env:COMPUTERNAME ;
$sQot = [char]34 ; $sQotS = [char]39 ; 
$NoProf=[bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
$MyBox="MyComputer","MyComputer" ; 
$DomainWork = "DOMAIN";
$DomHome = "DOMAIN";
$DomLab="DOMAIN";
#endregion INIT; # ------ 

#region LOCALCONSTANTS ; # ------ 
$targControllerSlot = 1 # 0=integrated (C?), 1=Ex add-on JBOD drives
$SMTPTo=@("SITEDLISMessagingReports@DOMAIN.com") ; 
if($env:computername -like 'SITED*'){
    $SMTPTo+="SITEDDLISMessagingReports@DOMAIN.com"
} ; 
if($env:computername -like 'SITE*'){
    $SMTPTo+="SITEDLISMessagingReports@DOMAIN.com"
} ; 

# dir in which the status semaphore/flag files will be stored
$FlagsDir="c:\scripts\flags" ; 
# we'll just store vari.xml files in same dir with the script, but distinctive name $ScriptNameNoExt-varis.xml
if (!(test-path $FlagsDir)) {Write-Host "Creating dir: $($FlagsDir)" ;mkdir $FlagsDir ;} ;
#endregion LOCALCONSTANTS ; # ------ 

# Clear error variable
$Error.Clear() ; 


#*======v FUNCTIONS v======
#*----------------v Function Start-IseTranscript v----------------
Function Start-IseTranscript {
    <#
    .SYNOPSIS
    This captures output from a script to a created text file
    .NOTES
    NAME:  Start-iseTranscript
    EDITED BY: Todd Kadrie
    AUTHOR: ed wilson, msft
    REVISIONS: 
    * 8:40 AM 3/11/2015 revised to support PSv3's break of the $psise.CurrentPowerShellTab.consolePane.text object
        and replacement with the new...
            $psise.CurrentPowerShellTab.consolePane.text
        (L13 FEs are PSv4, SITE650 is PSv2)
    * 9:22 AM 3/5/2015 tweaked, added autologname generation (from script loc & name)
    * 09/10/2010 17:27:22 - original
    TYPICAL USAGE: 
        Call from Cleanup() (or script-end, only populated post-exec, not realtime)
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        if($host.Name -eq "Windows PowerShell ISE Host"){
                # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
                $Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
                write-host "`$Logname: $Logname";
                Start-iseTranscript -logname $Logname ;
                # optional, normally wouldn't archive ISE debugging passes
                #Archive-Log $Logname ;
            } else {
                if($bdebug){ write-host -ForegroundColor Yellow "$((get-date).ToString('HH:mm:ss')):Stop Transcript" };
                Stop-TranscriptLog ; 
                if($bdebug){ write-host -ForegroundColor Yellow "$((get-date).ToString('HH:mm:ss')):Archive Transcript" };
                Archive-Log $transcript ; 
            } # if-E
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


    KEYWORDS: Transcript, Logging, ISE, Debugging
    HSG: WES-09-25-10    
    .DESCRIPTION
    use if($host.Name -eq "Windows PowerShell ISE Host"){ } to detect and fire this only when in ISE
    .PARAMETER Logname
    the name and path of the log file.
    .INPUTS
    [string]/fso file object
    .OUTPUTS
    [io.file]
    .EXAMPLE
    Archive-Log $Logname ;
    Archives specified file to Archive
    .Link
        Http://www.ScriptingGuys.com
    #Requires -Version 2.0
    #>

    Param(
    [string]$Logname
    )
  
    if (!($scriptDir) -OR !($scriptNameNoExt)) {
        throw "`$scriptDir & `$scriptNameNoExt are REQUIRED values from the main script SUBMAIN. ABORTING!"
        # can't interpolate from here, because the script because the invokation is the function, rather than the script
    } else {
        if (!($Logname)) {
            # build from script if nothing passed in
            $Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans-log.txt")) ;
            write-host "ISE Trans `$Logname: $Logname";
        };
$TranscriptHeader = @"
**************************************
Windows PowerShell ISE Transcript Start
Start Time: $(get-date)
UserName: $env:username
UserDomain: $env:USERDNSDOMAIN
ComputerName: $env:COMPUTERNAME
Windows version: $((Get-WmiObject win32_operatingsystem).version)
**************************************
Transcript started. Output file is $Logname
"@
        $TranscriptHeader | out-file $Logname -append
        if (($host.version) -lt "3.0") {
            # use legacy obj
            $psISE.CurrentPowerShellTab.Output.Text | out-file $Logname -append
        } else {
            # use the new object
            $psISE.CurrentPowerShellTab.ConsolePane.text | out-file $Logname -append
        } # if-E
    }  # if-E
} #*----------------^ END Function start-iseTranscript ^---------------- 

#---------------------------v Function Cleanup v---------------------------
function Cleanup {
    <# 
    .SYNOPSIS
    Cleanup() - Cleanup Function - cleans up objects before script exit 
    .NOTES
    Written By: Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	http://twitter.com/tostka

    Change Log
    # 1:17 PM 4/26/2017 customized to cover transciptise variant names, etc for send-emailnotification()
    # 8:49 AM 6/21/2016 added test for $bWhatif, and switched logname to include -WHATIF- indicator
    # 11:33 AM 4/7/2016 added support for $Script:ExPSSPersist to to permit EMS Machine: sessions to retain their permanent existance
    # 11:28 AM 3/31/2016 validated that latest round of updates are still functional, minor cleanup rem'd code
    #2:25 PM 3/24/2016 cleanup, working
    1:03 PM 3/23/2016 initial version
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Cleanup ;
    .EXAMPLE
    .LINK
    #>

    # clear all objects and exit
    # Clear-item doesn't seem to work as a variable release 
   
    # # 8:46 AM 3/11/2015 at some time from then to 1:06 PM 3/26/2015 added ISE Transcript
    # 8:39 AM 12/10/2014 shifted to stop-transcriptLog function
    # 7:43 AM 1/24/2014 always stop the running transcript before exiting
    if ($bdebug) {"CLEANUP"}

    #*------v EMS/AD CLOSE Block v------
    # 1:40 PM 3/16/2016 working version port
    if(($script:ExPSS -ne $false) -AND ($script:ExPSS -ne $true)) {
        # reset EMS search settings
        if(get-command -name set-AdServerSettings -ea 0){ set-AdServerSettings -ViewEntireForest $false -ea 0 ; } ; 
        # close any existing open $script:ExPSS
        # 11:33 AM 4/7/2016 leak to permit EMS Machine: sessions to retain their permanent sessions
        if($script:ExPSSPersist) {
            if($bDebug){Write-Debug "$((get-date).ToString("yyyyMMdd HH:mm:ss")):`$Script:ExPSSPersist set, Exch Tools EMS running, leaving ExPSS session intact ID#$($script:ExPSS.ID)"; } ;
        } else { 
            if($bDebug){Write-Debug "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Closing existing `$script:ExPSS session ID#$($script:ExPSS.ID)"; } ;
            if($script:ExPSS){Remove-PSSession -ID $($script:ExPSS).ID ; } ; 
        } # if-E
    }  ; 
    #11:54 AM 4/26/2017 in this case, we need to expliclty close EMs
    $sName="Microsoft.Exchange.Management.PowerShell*"; if ((Get-PSSnapin | where {$_.Name -eq $sName})) {Remove-PSSnapin $sName -ea Continue};
    if($script:AdPSS) {
        # close any existing open $script:AdPSS   
        if($bDebug){Write-Debug "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Closing existing `$script:AdPSS Module "; } ;
        $script:AdPSS | remove-Module ; 
    }  ; 
    if ($ShowDebug -OR ($DebugPreference = "Continue")) {
            Write-Verbose -Verbose:$true "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
            $bDebug=$false
            $DebugPreference = "SilentlyContinue" ;
    } # if-E ; 
    #11:54 AM 4/26/2017 in this case, we need to expliclty close ADMS Remove-module ActiveDirectory 
    $sName="activedirectory"; if ((get-module | where {$_.Name -eq $sName})) {Remove-Module $sName -ea Continue};

    if($ErrorActionPreference -eq 'Stop') {$ErrorActionPreference = 'Continue' ; write-debug "(Restoring `$ErrorActionPreference:$ErrorActionPreference;"};

    if($host.Name -eq "Windows PowerShell ISE Host"){
        write-host "`$transcript: $transcript";
        #Start-iseTranscript
        Start-iseTranscript -logname $transcript ;

    } else {
        if($bdebug){ write-host -ForegroundColor Yellow "$((get-date).ToString('HH:mm:ss')):Stop Transcript" };
        #Stop-TranscriptLog ; 
        Stop-Transcript
    } # if-E
    
    $attachments+=$transcript ; 
    
    # 10:32 AM 8/23/2017 use $bAlert semaphore to trigger mailing
    # 10:33 AM 8/24/2017a always email
    #if ($bAlert ){
    if($SendMsg){
        # 12:09 PM 4/26/2017 need to email transcript before archiving it
        if($bdebug){ write-host -ForegroundColor Yellow "$((get-date).ToString('HH:mm:ss')):Mailing Report" };
        if($FailedDriveCount ){
            $SmtpBody += "`n`r`$FailedDriveCount:$($FailedDriveCount)"
        } ; 
        $SmtpBody += "`n`rPass Completed $([System.DateTime]::Now)`nResults Attached:($transcript)" ;
        $SmtpBody += "`n`r'" +  ('-'*50) ;
        $email.body = $SmtpBody ; 
        if($attachments){ 
            $email.attachments=$attachments ; 
        } ; 
        
        if($whatif) {$SMTPSubject="WHATIF-"+$SMTPSubject ;  } 
        write-verbose -verbose:$true  $SMTPSubject ;
        $email.Subject = $SMTPSubject ; 
        $email|out-string ; 
        $email.body | out-string ; 
        Send-MailMessage @email 
        $SentNotifications++ ; 
        $CyclesSinceLastNotif = 0 ; 
    } else {
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):NOMINAL OR SUPPRESSED ALERT PASS, NO MAILED REPORT GENERATED" ;
        $CyclesSinceLastNotif++ ; 
    }; 

    # always strip StackTrace breakpoint if it's set
    if(Get-PSBreakpoint -Variable StackTrace){ Get-PSBreakpoint -Variable StackTrace | Remove-PSBreakpoint ; } ; 

    # export variables to xml
    <#$SentNotifications
    $CyclesSinceLastNotif
    $bAlert
    $FailedDriveCount 
    #>
    $rgxVariStore="^(SentNotifications|CyclesSinceLastNotif|bAlert)$" ; 
    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Exporting Variables:$($rgxVariStore) " ;
    #get-variable semaphore,count,xray* | export-clixml $varixml ;
    get-variable |?{$_.name -match $rgxVariStore} |%{"`nvari:$($_.Name):'$($_.Value)'" } ;
    get-variable |?{$_.name -match $rgxVariStore} | export-clixml $VariXml ;

    
    #11:10 AM 4/2/2015 add an exit comment
    Write-Verbose "END ==== $scriptBaseName ====" -Verbose:$verbose
    Write-Verbose "---------------------------------------------------------------------------------" -Verbose:$verbose

    exit
} #*----------------^ END Function Cleanup ^----------------

#*------v Function CheckSmartArray v------
function CheckSmartArray {
    Write-Host -foregroundcolor green "`nChecking SmartArray on system $($env:COMPUTERNAME), controller slot $($targControllerSlot)..." ; 
    C:\Windows\System32\cmd.exe /c "C:\Program Files (x86)\Compaq\Hpacucli\Bin\hpacucli.exe" controller slot=$($targControllerSlot) physicaldrive all show 
} ; #*------^ END Function CheckSmartArray ^------

#*======^ END FUNCTIONS ^======

#*======v SUB MAIN v======
$transcript=$ScriptDir + "logs"
if (!(test-path $transcript)) {Write-Host "Creating dir: $transcript" ;mkdir $transcript ;} ;
$transcript+="\" + $ScriptNameNoExt + "-" + "$(get-date -format 'yyyyMMdd-HHmmtt')" + "-trans.log.txt" ;

$logfile = "$(join-path -path (join-path -path $ScriptDir -childpath "logs") -childpath $ScriptNameNoExt)-$($env:COMPUTERNAME)-$(get-date -format 'yyyyMMdd-HHmmtt').log.txt" ; 
# purge prior logs with variant of the scriptnamename
gci "$(join-path -path (join-path -path $ScriptDir -childpath "logs") -childpath $ScriptNameNoExt)-*.log.txt" -ea 0 | Remove-Item -ea 0 ; 


$FlagFile = $FlagsDir+="\" + $ScriptNameNoExt + "-ALERTING.flag" ;
$VariXml = join-path -path $ScriptDir -childpath "$($ScriptNameNoExt)-$($env:COMPUTERNAME)-varis.xml"
# read in variables from prior
if(test-path -path $VariXml){
    Import-Clixml $varixml | %{ 
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Populating vari:$($_.Name):'$($_.Value)'" ;
        Set-Variable $_.Name $_.Value  ; 
    } ;
} ; 

$EmailEveryXCycles= 4 ; 
$EmailEveryXCycles= $EmailEveryXCycles -1 ; # doesn't count the first cycle, so 4 would end up = every 5
[array]$attachments=$null ;

write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):#*======v START PASS v======" ; 

$Error.Clear() ; 
TRY {
    if($host.name -ne 'Windows PowerShell ISE Host'){
        start-transcript -path $transcript ;
    } else { "ISE detected, skipping transcript" } ; 

    If(get-service|?{$_.name -like 'MSExchange*'}){
        $sName="Microsoft.Exchange.Management.PowerShell.E2010"; if (!(Get-PSSnapin | where {$_.Name -eq $sName})) {$EMSSnap = Add-PSSnapin $sName -ea Stop} ; 
    } ; 

    if (Get-PSSnapin | where {$_.Name -eq $sName}){
        # Dyn SMTP Server finder: from Ex server desktops with EMS loaded
        if((get-exchangeserver $env:computername).IsHubTransportServer) { 
          $SMTPServer = "$($env:computername)" ; 
        } else {
          $SMTPServer = (get-transportserver "$(($env:computername).substring(0,3))*" | select -expand name | get-random );
        }; 

        $email=@{
            SmtpServer=$SMTPServer ;
            From=$(($ScriptBaseName.replace(".","-")) + "@DOMAIN.com")  ;
            To=$SMTPTo ; 
            Priority = "Normal" ; 
            Body=$null ; 
        } ; 

        $tDag = Get-DatabaseAvailabilityGroup dag-au | ?{$_.servers -contains $($env:COMPUTERNAME)} | select -expand name ;
        
        # run status pass, all phys drives on slot 1
        CheckSmartArray | out-file -filepath $logfile -append ; 

        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):DETAILED OUTPUT FOR CTRL IN SLOT $($targControllerSlot):`n(blank if none)" ;
        [array]$logcontent=gc $logfile ; 
        $logcontent | out-string ; 
        write-host "`n$((get-date).ToString('HH:mm:ss')):CHECKING FOR 'FAILED'S IN OUTPUT:`n(blank if none)" ;
        $FailedDriveCount = 0 ; 
        foreach ($line in $logcontent) {
            if ($line -match "Failed") {
                $FailedDriveCount ++ ; 
                $SmtpSubject = "FAILED DISK FOUND ON $($env:COMPUTERNAME) ($((get-date -format 'yyyyMMdd-HHmmtt')))" ; 
                write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):FAILED DISK ON $($env:COMPUTERNAME)" ;
                $SmtpBody += "`n`rFAILURE:`n$($line)`n`rdetailed logs can be found $($logfile.replace("C:\","\\$($env:COMPUTERNAME)\c$\"))" ; 
                write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):$($SmtpBody)" ; 
                
                $FailedDbs=Get-DatabaseAvailabilityGroup $tDAG | select -expand servers | 
                    % {Get-MailboxDatabaseCopyStatus -Server $_ | where {$_.Status -match '^(Failed|FailedAndSuspended)'}}
                $faileddbs|%{
                    $fdbName= $_.identity.tostring().split("\")[0] ; 
                    $sMsg = "===STATUS OF DBS IN $($tDAG) WITH FAILED REPLICAS:" ; 
                    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ;
                    $SmtpBody += "`n`r" + $sMsg ; 
                    $CopyStat= Get-MailboxDatabaseCopyStatus -identity $fdbName | select name,status,LatestFullBackupTime,LatestIncrementalBackupTime ; 
                    $SmtpBody += "`n`r$(($CopyStat| out-string).trim())" ; 
                    $(($CopyStat| out-string).trim()) ; 
                    $sMsg = "===EdbFilePath of affected dbs" ; 
                    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ;
                    $SmtpBody += "`n`r" + $sMsg ; 
                    $EdbFilePath = Get-MailboxDatabase -identity $fdbName | select name,EdbFilePath ; 
                    $SmtpBody += "`n`r$(($EdbFilePath| out-string).trim())" ; 
                    $(($EdbFilePath| out-string).trim()) ; 
                } ; 
            };  # if-E fail
        } # loop-E ; 
        
        $attachments+=$logfile ; 
        
        if($FailedDriveCount -eq 0 -AND $bAlert ){
            # reset semaphore if present but no longer in Fail state
            $SmtpSubject = "DISK RECOVERY ON $($env:COMPUTERNAME) ($((get-date -format 'yyyyMMdd-HHmmtt')))" ; 
            $SmtpBody = "Returning to NOMINAL status`n`r$($line)`n`rdetailed logs can be found $($logfile.replace("C:\","\\$($env:COMPUTERNAME)\c$\"))" ; 
            $SendMsg=$true ; 
            $bAlert = $false ; 

        } elseif($FailedDriveCount -ge 0 -AND -not $bAlert){
            # new alert
            $SendMsg=$true ; 
            $bAlert = $true ; 
        } elseif($FailedDriveCount -AND $bAlert ) {
            # existing alert, only send on after $EmailEveryXCycles
            if($CyclesSinceLastNotif -ge $EmailEveryXCycles){
                $SendMsg=$true ; 
                write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):TRIGGERING ALERT NOTIFICATION" ;
            } else {
                $SendMsg=$false ;
                write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):ALERT STATUS, SENT $($SentNotifications++) NOTIFICATIONS, `n`r$($CyclesSinceLastNotif) CYCLES SINCE LAST NOTIFICATION`r`nWAITING UNTIL $($EmailEveryXCycles) CYCLES TO SEND A NEW NOTIFICATION" ;
            }
        } else { 

        } ; 

        
    } else { write-error "$((get-date).ToString('HH:mm:ss')):NO EMS FOUND, FAILED LOAD, OR SCRIPT NOT RUNNING ON AN EXCHANGE ROLE SERVER! ABORTING!"; } ; 

    Cleanup ; 

} CATCH {
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
    Write-Error  "$(get-date -format "HH:mm:ss"): $($msg)" ; 
    <#
    set-AdServerSettings -ViewEntireForest $false ;
    Exit ;
    #>
    Cleanup
} ; 
#*======^ END SUB MAIN  ^======


