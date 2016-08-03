# backup-Lync13Configs.ps1
# debug: cls ; .\backup-Lync13Configs.ps1 -showdebug -whatif ;
<# 
.SYNOPSIS
backup-Lync13Configs.ps1 - Backup a variety of Lync2013 config items to text files, and then robocopy the collected content to DR Server
.NOTES
Based on code Written By: Richard Brynteson Avtex 2013
Website:	http://masteringlync.com/2013/03/31/backup-for-lync-server-2013/
Updated by Todd Kadrie
Website: http://tinstoys.blogspot.com/
Change Log
* 11:41 AM 8/3/2016 moved transcribing up higher
* 11:26 AM 8/3/2016 update attributions, added comment-based help for -showdebug param
* 11:15 AM 8/3/2016 fixed typo in the RGS status echo code
* 11:12 AM 8/3/2016 added schtasks cmd docs, duped the robocopy exitcodes into the whatif block, added bracketing ===^ PASS STARTED ^ === to confirm transcript runs to completion of script.
* 10:30 AM 8/3/2016 fundemental rewrite, retained Richards core CS , commands, but better logic (pre-testing of dumpable objects), logging and redundancy in use, also added -showdebug, -whatif & transcript logging, and capture of robocopy error status
* 7:33 AM 8/2/2016 - TSK reformatted, cleanedup, added pshelp
* 3/31/13 - posted version
Task-creation command: (run at 1am daily, as account logon)
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
schtasks /create /tn "backup-Lync13Configs.ps1" /tr "powershell.exe -noprofile -file 'c:\scripts\backup-Lync13Configs.ps1'" /sc DAILY /ST 01:00 /ru domain\account /rp [PASSWORD] ;
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
.DESCRIPTION
backup-Lync13Configs.ps1 - Backup a variety of Lync2013 config items to text files, and then robocopy the collected content to DR Server
.PARAMETER showDebug
showDebug switch flag
.PARAMETER whatif
Whatif switch flag
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Outputs a folder of config dump files
.EXAMPLE
.LINK
http://masteringlync.com/2013/03/31/backup-for-lync-server-2013/
#>

# 9:12 AM 8/3/2016: defaulted $whatif to false
Param(
    [Parameter(HelpMessage='Debugging Flag [-showDebug]')]
    [switch] $showDebug
    ,[Parameter(HelpMessage='Whatif Flag  [-whatIf]')]
    [switch] $whatIf=$false
) # PARAM BLOCK END

#region INIT; # ------ 
if ($ShowDebug) {$bDebug=$true ; $DebugPreference = "Continue" ; write-debug "(`$ShowDebug:$ShowDebug ;`$bDebug:$bDebug ;`$DebugPreference:$DebugPreference)" ; };
if ($Whatif){$bWhatif=$true ; Write-Verbose -Verbose:$true "`$Whatif is $true (`$bWhatif:$bWhatif)" ; }
else {$bWhatif=$false} ; 
#if($bdebug){$ErrorActionPreference = 'Stop' ; write-debug "(Setting `$ErrorActionPreference:$ErrorActionPreference;"};
# scriptname with extension
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ; 
$ComputerName = $env:COMPUTERNAME ;
$sQot = [char]34 ; $sQotS = [char]39 ; 
$NoProf=[bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
#$ProgInterval= 500 ; # write-progress wait interval in ms
#endregion INIT; # ------ 


#region SETUP
$TimeStampNow = get-date -uformat "%Y%m%d-%H%M"
$outtransfile=$ScriptDir + "logs"
if (!(test-path $outtransfile)) {Write-Host "Creating dir: $outtransfile" ;mkdir $outtransfile ;} ;
# timestamped log file
#$outtransfile+="\" + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.log" ;
# 10:32 AM 8/3/2016 recycled logfile every day
$outtransfile+="\" + $ScriptNameNoExt + "-trans.log" ;

start-transcript -path $outtransfile #-se 0 ;
write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):===v PASS STARTED v ===" ;

# variables/contstants
$dumppath = "D:\BackupFromSITE" ; 
$RetainDays = "-10" ; 
switch ($($env:userdomain)) {
  "domain" {
      $poolFQDN = "pool.domain.com" ; 
      $drFolderPath = "\\LyncFileShare\FileShare\scripts\cfg-backups\FromSITE" ; 
  }
  "domain-LAB"{
      $poolFQDN = "pool.domainlab.com" ; 
      $drFolderPath = "\\LyncFileShare\FileShare\scripts\cfg-backups\FromSITE" ; 
  }
  default{ throw "UNRECOGNIZED DOMAIN:$($env:userdomain)!" }
} ; 


if($bDebug){
    Write-Debug "$((get-date).ToString("HH:mm:ss")):`$outtransfile:$outtransfile`n`$ComputerName:$ComputerName`n`$poolFQDN:$poolFQDN`n`$drFolderPath:$drFolderPath"; 
} ;
#endregion SETUP

# Clear error variable
$Error.Clear() ; 
TRY {
    # 8:03 AM 8/2/2016 TSK, test for running on a Lync FE
    if(!(Get-Service |?{$_.name -eq 'RtcSrv'})){ throw "$((get-date).ToString("HH:mm:ss")):Missing RTCSRV service:`nThis script can only be run locally from a Lync Front End server "} ; 
    # 7:37 AM 8/2/2016 TSK, test for registered, but not loaded
    $mName="Lync" ; 
    if(get-Module -ListAvailable | where {$_.Name -eq $mName}) {
        if (!(get-Module  | where {$_.Name -eq $mName})) {
            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):Loading Module: $mName"; } ;
            Import-Module $mName -ea Stop ;
        } ; 
    } else {
        throw "$((get-date).ToString("HH:mm:ss")):Missing Key-requirement: No registered Lync Module.`nThis script can only be run locally from a Lync server "; 
    } ; 

    # validate dirs present, create if missing
    if (!(test-path $dumppath)) {Write-Host "Creating dir: `$dumppath" ;mkdir $dumppath -WhatIf:$whatif ;} 
    if (!(test-path $drFolderPath)) {Write-Host "Creating dir: `$drFolderPath" ;mkdir $drFolderPath -WhatIf:$whatif ;} 
    
    # clean up local out of date files & empty folders
    $dfiles=get-childitem -path $dumppath -recurse | ?{ ($_.LastWriteTime -lt (get-date).adddays($RetainDays)) -AND (!$_.psiscontainer)} ; 
    if ($dfiles) {
        if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):Purging old files"; } ;
        $dfiles | Remove-Item  -path $_.fullname -force -WhatIf:$whatif; 
    } ;
    $subfolders = get-childitem $dumppath|?{$_.psiscontainer -AND $_.GetFiles().Count -eq 0} ;
    if ($subfolders) {
        if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`Purging Empty Subfolders"; } ;
        $subfolders | Remove-Item -WhatIf:$whatif; 
    } ; 
    
    # Create local dump subfolder
    $currDate = (get-date -format "yyyyMMdd-HHmmtt") ;  
    New-Item -path "$dumppath\$currDate" -Type Directory -WhatIf:$whatif ; 

    # clean up $drFolderPath out of date files & empty folders
    $dfiles=get-childitem  -path $drFolderPath -recurse | ?{ ($_.LastWriteTime -lt (get-date).adddays($RetainDays)) -AND (!$_.psiscontainer)} ; 
    if ($dfiles) {
        if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):Purging old files"; } ;
        $dfiles | Remove-Item  -path $_.fullname -force -WhatIf:$whatif; 
    } ; 
    $subfolders = get-childitem $drFolderPath |?{$_.psiscontainer -AND $_.GetFiles().Count -eq 0} ;
    if ($subfolders) {
        if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`Purging Empty Subfolders"; } ;
        $subfolders | Remove-Item -WhatIf:$whatif; 
    } ; 
    
    if(!$bWhatif){
        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Backup to local $($dumppath) in progress..." ;

        #Export CMS/XDS and LIS
        Export-CsConfiguration -FileName $dumppath\$currDate\XdsConfig.zip ; 
        Export-CsLisConfiguration -FileName $dumppath\$currDate\LisConfig.zip ; 
        #Export Voice Information
        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Exporting Voice Configuration..." ;
        Get-CsDialPlan | Export-Clixml -path $dumppath\$currDate\DialPlan.xml ; 
        Get-CsVoicePolicy | Export-Clixml -path $dumppath\$currDate\VoicePolicy.xml ; 
        Get-CsVoiceRoute | Export-Clixml -path $dumppath\$currDate\VoiceRoute.xml ; 
        Get-CsPstnUsage | Export-Clixml -path $dumppath\$currDate\PSTNUsage.xml ; 
        Get-CsVoiceConfiguration | Export-Clixml -path $dumppath\$currDate\VoiceConfiguration.xml ; 
        Get-CsTrunkConfiguration | Export-Clixml -path $dumppath\$currDate\TrunkConfiguration.xml ; 
        #Export RGS Config - includes agent groups, queues and workflows.
        # fails if non-configed, precheck & run conditionally
        if(Get-CsApplicationEndpoint -Filter {OwnerUrn -like "*RGS*"} | ?{$_.Registrarpool -like $poolFQDN}){
            if((Get-CsRgsAgentGroup -identity "service:ApplicationServer:$poolFQDN") -OR (Get-CsRgsQueue  -identity "service:ApplicationServer:$poolFQDN") -OR (Get-CsRgsWorkflow  -identity "service:ApplicationServer:$poolFQDN")){
            Export-CsRgsConfiguration -Source "service:ApplicationServer:$poolFQDN" -FileName $dumppath\$currDate\RgsConfig.zip ; 
            } else { write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):(No Response-Group Agent groups, Queues or Workflows are configured in pool:" ;}
        } else { write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):(No Response-Group EndPoints configured in pool $($poolFQDN))" } ; 
        #Export User/Contact Info, pool level
        Export-CsUserData -PoolFqdn $poolFQDN -FileName $dumppath\$currDate\UserData.zip ; 
    } else { 
        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):-WHATIF in use, skipping export commands (-whatif not supported n export-/get-)." ;
    } ; 
    write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):XDS, LIS, User and RGS backup to server is completed.  Files are located at $dumppath\$currDate" ;

    #Copy Files to DR Server
    # syntax: ROBOCOPY source destination [file [file]...] [options]
    # copyflags in use: D=Data, A=Attributes, T=Timestamps, S=Security=NTFS ACLs, O=Owner info
    # /S :: copy Subdirectories, but not empty ones.
    #robocopy $dumppath $drFolderPath /COPY:DATSO /S ; 
    
    $rCopyFlags="/COPY:DATSO /S"  ; 
    write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Robocopying $($dumppath)\$($currDate) to $($drFolderPath)...  " ;
    if(!$bWhatif){
        # capture and echo back errors from robocopy.exe, via redirected streams into tempfiles
        $soutf= [System.IO.Path]::GetTempFileName() ; 
        $serrf= [System.IO.Path]::GetTempFileName()  ; 
        $process = Start-Process -FilePath ROBOCOPY.exe -ArgumentList "$($dumppath) $($drFolderPath) $($rCopyFlags)" -NoNewWindow -PassThru -Wait -RedirectStandardOutput $soutf -RedirectStandardError $serrf ;  
        $process | out-string ;  
        <# for non-robocopy, use this Exittest
        switch($process.ExitCode){
            1 {write-host "ExitCode 1 returned (Error)"} ; 
            0 {write-host "ExitCode 0 returned (No Error)"} ; 
            default {write-host "other ExitCode returned $($process.ExitCode)"} ; 
        } ;  ; 
        #>
        <# But robocopy.exe returns non-standard ExitCodes: 
        ExitCode 0 = No Errors, No Files Copied 
        ExitCode 1 = No Errors, New Files Copied. actually a good thing.
        ExitCode > 1 = There were errors, or at least unexpected results
        #>
        switch($process.ExitCode){
            1 {write-verbose -verbose:$true "$((get-date).ToString("HH:mm:ss")):ExitCode 1 returned (No Errors, New Files Copied)"} ; 
            0 {write-verbose -verbose:$true "$((get-date).ToString("HH:mm:ss")):ExitCode 0 returned (No Errors, No Files Copied)"} ; 
            default {write-verbose -verbose:$true "$((get-date).ToString("HH:mm:ss")):ERROR during RoboCopy: Non-1/0 ExitCode returned $($process.ExitCode)"} ; 
        } ;  ; 
        if((get-childitem $soutf).length){ (gc $soutf) | out-string } ; 
        remove-item $soutf ;  ; 
        if((get-childitem $serrf).length){(gc $serrf) | out-string} ;  ; 
        remove-item $serrf  ; ; 
    } else { 
        # capture and echo back errors from robocopy.exe
        $soutf= [System.IO.Path]::GetTempFileName() ; 
        $serrf= [System.IO.Path]::GetTempFileName()  ;
        # if whatif, run /L = List-only mode 
        $process = Start-Process -FilePath ROBOCOPY.exe -ArgumentList "$($dumppath) $($drFolderPath) /COPY:DATSO /S /L" -NoNewWindow -PassThru -Wait -RedirectStandardOutput $soutf -RedirectStandardError $serrf ;  
        $process | out-string ;  
        switch($process.ExitCode){
            1 {write-verbose -verbose:$true "$((get-date).ToString("HH:mm:ss")):ExitCode 1 returned (No Errors, New Files Copied)"} ; 
            0 {write-verbose -verbose:$true "$((get-date).ToString("HH:mm:ss")):ExitCode 0 returned (No Errors, No Files Copied)"} ; 
            default {write-verbose -verbose:$true "$((get-date).ToString("HH:mm:ss")):ERROR during RoboCopy: Non-1/0 ExitCode returned $($process.ExitCode)"} ; 
        } ;  ; 
        if((get-childitem $soutf).length){ (gc $soutf) | out-string } ; 
        remove-item $soutf ;  ; 
        if((get-childitem $serrf).length){(gc $serrf) | out-string} ;  ; 
        remove-item $serrf  ; ; 
    }
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
    Write-Error  "$(get-date -format "HH:mm:ss"): $($msg)" ; 
    Exit ;
}  
FINALLY {
  write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):===^ PASS COMPLETED ^ ===" ;
	Stop-Transcript ; 
    if ($ShowDebug -OR ($DebugPreference = "Continue")) {
        Write-Verbose -Verbose:$true "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
        $bDebug=$false
        # 8:41 AM 10/13/2015 also need to enable write-debug output (and turn this off at end of script, it's a global, normally SilentlyContinue)
        $DebugPreference = "SilentlyContinue" ;
    } # if-E ; 
    if($ErrorActionPreference -eq 'Stop') {$ErrorActionPreference = 'Continue' ; write-debug "(Restoring `$ErrorActionPreference:$ErrorActionPreference;"};
} ; 
