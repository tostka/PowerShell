#check-ExchSvcStatusRpt.ps1
# clean out the set-csuserpin overhead code

# *** REGION SETUP MARKER
#region SETUP ; 

#*------V Comment-based Help (leave blank line below) V------ 

<# 
.SYNOPSIS
check-ExchSvcStatusRpt.ps1 - mailed-report version of the svcstatus script
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Change Log
# 8:45 AM 10/3/2017 added post test-mapi to random mbx per store. Fixed post svc-remediation report output 
    (was using | format-table vs | select) in the fragment. Debugged on lab-Server1. 
# 1:33 PM 10/2/2017 retooling for remedial svc start, have it back to html function
# 10:05 AM 12/21/2015 added mbxrole local db-named queue rpt
# 1:40 PM 12/17/2015 the Win2008R2 task delay doesn't seem to work for '3 mins', so lets detect and hand set it.
# 2:21 PM 12/16/2015 #1086: added 'blank' queues returned option, to avoid blank space in report  ;
tested & functional on Server0 too (had to change layout of some reports to -LIST)
* 9:42 AM 12/14/2015 updated Send-EmailNotification, tested for & tweaked psv2 params, which Psv3 has, and v2 does not. Also updated the attachment code to properly  test for status. And finally switched @$email to base model, and tested to add extra attributes, rather than keeping multiple concurrent full versions
* 12:51 PM 3/20/2015 2nd vers
.PARAMETER showDebug
Show Debugging Output.
.PARAMETER  NoWait
Override default post-reboot wait (for testing)
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Outputs a text transcript and emails the report to $rptAdmin (emailadmin@DOMAIN.com)
.EXAMPLE
.\check-ExchSvcStatusRpt.ps1
.LINK
*------^ END Comment-based Help  ^------ #>

[CmdletBinding()]
Param(
    [Parameter(HelpMessage='Debugging Flag [$switch]')]
    [switch] $showDebug,
    [Parameter(HelpMessage='Skip run delay[$switch]')]
    [switch] $NoWait
) # PARAM BLOCK END

$sPollDelay=300 ; # # seconds to wait to check services
# 12:50 PM 12/16/2015 svcs that are normally down per role
$rgxExHubNonRunSvc="^(msftesql-Exchange)$" ; 
$rgxExMbxNonRunSvc="^(wsbexchange)$" ; 
$rgxExCasNonRunSvc="" ; 

# 12:07 PM 10/2/2017 disable poll delay when -NoWait
if($NoWait){ 
    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):NoWait specified: Disabling `$sPollDelay$($sPollDelay)" ; 
    $sPollDelay=0 
} ; 

#*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
# 2:10 PM 2/4/2015 shifted to here to accommodate include locations
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName)) ;
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ; 
$ComputerName = $env:COMPUTERNAME ;
# 12:48 PM 3/11/2015 detect -noprofile runs (in case you need to add profile content/functions to get to function)
# pick up the bDebug from the $ShowDebug switch parameter
if ($ShowDebug) {$bDebug=$true};
if($bdebug){
  $ErrorActionPreference = 'Stop';
};
If ($Whatif){write-host "`$Whatif is $true" ; }; 
#$ProgInterval= 500 ; # write-progress wait interval in ms
# 12:23 PM 2/20/2015 add gui vb prompt support
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null ; 
# 11:00 AM 3/19/2015 should use Windows.Forms where possible, more stable

# Clear error variable
$Error.Clear()

#*======^ SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE ^======

#*======v HTML-EMAIL-DECLARE-BOILERPLATE v======
# 2:20 PM 3/25/2015 added optional report inline-in email body
$BodyAsHtml=$true ;
# 12:37 PM 12/16/2015 send-mailmessage param has a different name, rename $bodyashtml to match
# 12:15 PM 2/9/2015 add an SMTP retry limit (per user attempted)
[int]$WelcomeRetryLimit=4;
[int]$RetryDelay=20;    # wait time after failure
#*======^ HTML-EMAIL-DECLARE-BOILERPLATE ^======

#*======v HTML-CSS-DECLARE-BOILERPLATE v======
# 2:59 PM 3/24/2015 functional
# edit the following to suit:
# -------
$Author = "Todd Kadrie" ;
#$PageTitle="Exchange Server Health Report";
$PageTitle="Exchange Server Health Report ($($env:COMPUTERNAME))";
#$MainCaption = $PageTitle + "<br>Processing Results $((get-date).ToString("HH:mm:ss"))" ;
# 2:00 PM 3/25/2015 make it a 2-line, and add date
$MainCaption = $PageTitle + "<br>Processing Results: <br>$((get-date -uformat "%m/%d/%Y %H:%M"))" ;
#region HTMLASSEMBLY ; # ------- 
$sFileDate = get-date -uformat "%Y%m%d-%H%M" ;
$sRptDateTime=(get-date -uformat "%m/%d/%Y-%H:%M") ;
$ExecUser=($env:USERNAME) ;
#$FileBaseName = "Processes_$sFileDate" ;
# 2:06 PM 3/25/2015 explicit report names in this script
# recycle and illegal-strip the $PageTitle+date:
$FileBaseName = "$($PageTitle + "-" + $sFileDate)" ;
$FileBaseName = [RegEx]::Replace($FileBaseName, "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())), '') ;
# 12:19 PM 12/16/2015 need to path it

$RptFile =(join-path -path $ScriptDir -childpath "logs")
if (!(test-path $RptFile)) {Write-Host "Creating dir: $RptFile" ;mkdir $RptFile ;} ;
$RptFile=(join-path -path $RptFile -childpath ($FileBaseName + ".html"))
if (test-path $RptFile){ remove-item -path $RptFile } ;
if (test-path $RptFile.replace(".html",".csv")){ remove-item -path $RptFile.replace(".html",".csv") } ;

# embedded CSS version
#$CSS = get-content c:\usr\work\SITEc\scripts\tor-incl-base-graybar-borders.css ;
# pull it from the scriptpath
$CSS = get-content (join-path -path $ScriptDir -childpath "tor-incl-base-graybar-borders.css")
$sHTMLhead = "<title>$PageTitle</title><style>$CSS</style>" ; 
# 2:44 PM 3/24/2015 DOMAIN logo titleblock
$sHtmlPreLogo = "<table border='0' cellpadding='3' cellspacing='3'><tbody><td><h2><P>$MainCaption</P></h2></td><td><IMG SRC='http://www.DOMAIN.com/Style%20Library/DOMAIN/images/DOMAINlogo.gif' ALIGN=RIGHT></td></tbody></table>" ;
$sHtmlPreTextLogo = "<table border='0' cellpadding='3' cellspacing='3'><tbody><tr><td><h2><p>$MainCaption</p></h2></td><td style='text-align: center;'><big><big><big><span style='font-family: Times New Roman,Times,serif; font-weight: bold; color: rgb(204, 0, 0);'>&nbsp;&nbsp;&nbsp;DOMAIN</span></big></big></big><br></td></tr></tbody></table>" ;
if($BodyAsHtml){$sHtmlPre=$sHtmlPreTextLogo}
else{$sHtmlPre=$sHtmlPreLogo};
# 2:45 PM 3/24/2015 stock footer
#$sHtmlFooter="<h3>For details, contact $Author<br>Report Created on $sRptDateTime by $ExecUser<br>executed on $($env:COMPUTERNAME)</h3>" ;
# 2:15 PM 3/25/2015 add a footer that quotes executed script $ScriptBaseName
$sHtmlFooter="<h3>For details, contact $Author<br>Report Created on $sRptDateTime by $ExecUser<br>$($ScriptBaseName) executed on $($env:COMPUTERNAME)</h3>" ;
#endregion HTMLASSEMBLY ; # -------
#*======^ HTML-CSS-DECLARE-BOILERPLATE ^======


#*======v FUNCTIONS  v======
#*------------v Function Test-TranscriptionSupported v------------
function Test-TranscriptionSupported {
  <#
  .SYNOPSIS
  Tests to see if the current host supports transcription.
  .DESCRIPTION
  Powershell.exe supports transcription, WinRM and ISE do not.
  .Example
  #inside powershell.exe
  Test-Transcription
  #returns true
  Description
  -------
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
}#*------------^ END Function Test-TranscriptionSupported ^------------

#*------------v Function Test-Transcribing v------------
function Test-Transcribing {
    <#.SYNOPSIS
    Tests for whether transcript (Start-Transcript) is already running
    .NOTES
    Author: Oisin Grehan
    URL: http://poshcode.org/1500
    requires -version 2.0

    Change Log
    10:13 AM 12/10/2014

    .INPUTS
    .OUTPUTS
    Outputs $TRUE/FALSE reflecting transcribe status
    .EXAMPLE
    if (Test-Transcribing) {Stop-Transcript} 
    #>

    
    if(!(test-path function:Test-TranscriptionSupported  ))
    {write-error "$((get-date).ToString("HH:mm:ss")):MISSING DEPENDENT FUNCTION:Test-TranscriptionSupported" ; CleanUp ; } ; 
    
    # 8:56 AM 2/24/2015 move the test outside of the try
    if($Host.Name -ne "Windows PowerShell ISE Host"){
        $externalHost = $host.gettype().getproperty("ExternalHost",
           [reflection.bindingflags]"NonPublic,Instance").getvalue($host, @())
        try {
            if (Test-TranscriptionSupported) {
                $externalHost.gettype().getproperty("IsTranscribing",
                    [reflection.bindingflags]"NonPublic,Instance").getvalue($externalHost, @())
            } else {
                
            };  # if-E

        } catch {
            Write-Warning "Tested: This host does not support transcription."
            <# 9:19 AM 2/24/2015 the externalhost call is failing and dumping into here
                $Transcript:C:\scripts\logs\set-CSUserPins-20150224-0918-trans.log
                WARNING: Tested: This host does not support transcription.
            #>
        } # try-E
    } else {
        write-host "Test-Transcribing:SKIP PS ISE does not support transcription commands [returning `$true]";
        return $true ; 
    } # if-E
} #*------------^ END Function Test-Transcribing ^------------

#*------------v Function Stop-TranscriptLog v------------
function Stop-TranscriptLog {
  <#.SYNOPSIS
  Stops & ARCHIVES a transcript file (if no archive needed, just use the stock Stop-Transcript cmdlet)
  .NOTES
  #Written By: Todd Kadrie
  #Website:	http://tinstoys.blogspot.com
  #Twitter:	http://twitter.com/tostka
  Requires test-transcribing() function
  
  Change Log
  # 10:19 AM 12/16/2015 repl: get-timestamp() with explicit $(Get-Date -Format "HH:mm:ss")
  # 1:18 PM 1/14/2015 added SITEc fs rpt share support
  # 10:54 AM 1/14/2015 added lab support (Server0\d$)
  # 10:11 AM 12/10/2014 tshot stop-transcriptlog archmove, for existing file clashes
  9:04 AM 12/10/2014 shifted more into the try block
  12:49 PM 12/9/2014

  .INPUTS
  leverages the global $transcript variable (must be set in the root script; not functions)

  .OUTPUTS
  Outputs $TRUE/FALSE reflecting successful archive attempt status

  .EXAMPLE
  Stop-TranscriptLog 
  #>
  
  #can't define $Transcript as a local param/vari, without toasting the main vari!
  if ($bDebug) {"SUB: stop-transcriptlog"}
  
  if(!(test-path function:Test-Transcribing ))
    {write-error "$((get-date).ToString("HH:mm:ss")):MISSING DEPENDENT FUNCTION:Test-Transcribing"; CleanUp ;} ; 
    
  # 10:48 AM 1/14/2015 adde lab support for archpath
  # 10:56 AM 1/14/2015 adde SITEc FS support
  

  if($Host.Name -ne "Windows PowerShell ISE Host"){
        Try {
            if ($bDebug) {write-host -foregroundcolor green "$(Get-Date -Format "HH:mm:ss"):`n`$outtransfile:$outtransfile" ;};
                if (Test-Transcribing) {
                    # can't move it if it's locked
                    Stop-Transcript
                    if ($bDebug) {write-host -foregroundcolor green "`$Transcript:$Transcript"} ;
                }  # if-E
        } Catch {
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Failed to move `n$Transcript to `n$Archpath"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Error in $($_.InvocationInfo.ScriptName)."
                Write-Error "$(Get-Date -Format "HH:mm:ss"): -- Error information"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Line Number: $($_.InvocationInfo.ScriptLineNumber)"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Offset: $($_.InvocationInfo.OffsetInLine)"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Command: $($_.InvocationInfo.MyCommand)"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Line: $($_.InvocationInfo.Line)"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Error Details: $($_)"
        }  # try-E;
  
        if (!(Test-Transcribing)) {  return $true } else {return $false};
    } else {
        write-host "Stop-Transcribing:SKIP PS ISE does not support transcription commands";
        return $true
    }
  

}#*------------^ END Function Stop-TranscriptLog ^------------

#*------------v Function Start-TranscriptLog v------------
function start-TranscriptLog {
  <#.SYNOPSIS
  Configures and launches a transcript
  .NOTES
  #Written By: Todd Kadrie
  #Website:	http://tinstoys.blogspot.com
  #Twitter:	http://twitter.com/tostka
  Requires test-transcribing() function
  
  Change Log
  # 10:19 AM 12/16/2015 repl: get-timestamp() with explicit $(Get-Date -Format "HH:mm:ss")
  # 10:19 AM 12/10/2014 cleanup
  # 8:48 AM 12/10/2014 cleanup
  12:36 PM 12/9/2014

  .INPUTS
  None

  .OUTPUTS
  Outputs $TRUE/FALSE reflecting transcribe status

  .EXAMPLE
  start-TranscriptLog $Transcript
  #>

  # $outransfile=$LogDir + $ScriptNameNoExt + "-" + $user.ToUpper()  + "-(" + $ticket + ")-" + $TimeStampNow + "-trans.txt" ;
  
  param(
    [parameter(Mandatory=$true,Helpmessage="Transcript location")]
    [ValidateNotNullOrEmpty()]
    [alias("tfile")]
    [string]$Transcript
  )

  # Have to set relative $ScriptDir etc OUTSIDE THE FUNC, build full path to generic core $Transcript vari, and then 
  # start-transcript will auto use it (or can manual spec it with -path)

    if(!(test-path function:Test-Transcribing ))
    {write-error "$((get-date).ToString("HH:mm:ss")):MISSING DEPENDENT FUNCTION:Test-Transcribing" ; CleanUp ;} ; 
    
    if($Host.Name -NE "Windows PowerShell ISE Host"){
        Try {
                if (Test-Transcribing) {Stop-Transcript} 
  
                if($bdebug) {$Transcript}
                # prevaidate specified logging dir is present
                $TransPath=(Split-Path $Transcript).tostring();
                if($bdebug) {$TransPath;}
                if (Test-Path $TransPath ) { } else {mkdir $TransPath};
                #invoke-pause2
                Start-Transcript -path $Transcript
                if (Test-Transcribing) {  return $true } else {return $false};
            } Catch {
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Failed to create $TransPath"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Error in $($_.InvocationInfo.ScriptName)."
                Write-Error "$(Get-Date -Format "HH:mm:ss"): -- Error information"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Line Number: $($_.InvocationInfo.ScriptLineNumber)"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Offset: $($_.InvocationInfo.OffsetInLine)"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Command: $($_.InvocationInfo.MyCommand)"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Line: $($_.InvocationInfo.Line)"
                Write-Error "$(Get-Date -Format "HH:mm:ss"): Error Details: $($_)"
            }  # try-E;

    } else {
        write-host "Test-Transcribing:SKIP PS ISE does not support transcription commands [returning $true]";
        return $true ; 
    };  # if-E


}#*------------^ END Function Start-TranscriptLog ^------------

#*------------v Function Send-EmailNotif v------------
Function Send-EmailNotif {
  # 12:39 PM 12/16/2015 ren SmtpBodyHtml to the real param name BodyAsHtml ; 
  #      repl: get-timestamp() with explicit $(Get-Date -Format "HH:mm:ss")
  # 12:23 PM 12/15/2015 revised back to a single $email hash, and test & add vers-specific fields etc as needed
  # 10:35 AM 8/21/2014 always use a port; tested for $SMTPPort: if not spec'd defaulted to 25.

  <#
  PARAM(
    [parameter(Mandatory=$true)]
    [alias("from")]
    [string] $SMTPFrom,
    [parameter(Mandatory=$true)]
    [alias("to")]
    [string] $SmtpTo,
    [parameter(Mandatory=$true)]
    [parameter(Mandatory=$true)]
    [alias("subj")]
    [string] $SMTPSubj,
    [parameter(Mandatory=$true)]
    [alias("server")]
    [string] $SMTPServer,
    [parameter(Mandatory=$true)]
    [alias("port")]
    [string] $SMTPPort,
    [parameter(Mandatory=$true)]
    [alias("body")]
    [string] $SmtpBody,
    [parameter(Mandatory=$false)]
    [alias("SmtpBodyHtml")]
    [switch] $BodyAsHtml,
    [parameter(Mandatory=$false)]
    [alias("attach")]
    [string] $attachment 
  )
  #>
					
    # before you email conv to str & add CrLf:
    $SmtpBody = $SmtpBody | out-string
    # just default the port if missing, and always use it

    # 9:06 AM 12/14/2015 baseline hash, no port no attachment, and add the fields as needed/supported
    $email = @{
        From = $SMTPFrom ; 
        To = $SMTPTo ; 
        Subject = $SMTPSubj ; 
        SMTPServer = $SMTPServer ; 
        Body = $SmtpBody ; 
    } ; 
    # # 9:05 AM 12/14/2015: actually port param isn't supported in PSv2 send-mailmessage; 25 is assumed
    if ($($host.version.major -gt 2) -and ($SMTPPort -eq $null)) {
        $SMTPPort = 25;
        $email.add("Port",$SMTPPort );
    }	else {
        if($SMTPPort) {write-verbose -verbose:$true  "Skipping specified 'SMTPPort' parameter: Powershellv2 does not support a Port specification for send-mailmessage (defaults to 25 only)"} ; 
    } # if-block end ; 

    # define/update variables into $email splat for params
    if ($attachment -AND (test-path $attachment)) {
        $email.add("Attachments",$attachment);
    } ; 
    
    # 12:37 PM 12/16/2015 this isn't picking up $BodyAsHtml, because my Send-EmailNotification() had the parm labeled SmtpBodyHtml, changed it to match with an alias for orig in the param block
    if($BodyAsHtml){
        #$email.BodyAsHtml = $true
        $email.add("BodyAsHtml",$true);
    } # if-E
	
    <# 11:59 AM 8/28/2013 mailing debugging code
    write-host  -ForegroundColor Yellow "Emailing with following parameters:"
    $email
    write-host "body:"
    $SmtpBody
    write-host ("-" * 5)
    write-host "body.length: " $SmtpBody.length
    #>
    write-host "sending mail..."
    #$email
    # echo BIG multiline component, except it from standard loop:
    $BARS=("="*10);
    foreach($row in $email) {
        foreach($key in $row.keys) {
        if($key -eq "Body"){
            write-host "$($key): `n $BARS v BODY v $BARS `n$($row[$key])`n $BARS ^ BODY ^ $BARS " 
        } else {
            write-host "$($key): $($row[$key])" ;
        };
        } # loop-E; 
    } # loop-E ;

    $error.clear() 

    # 8:57 AM 1/29/2015 add a try/catch to it, to echo full errors
    TRY {
        #invoke-pause2
        send-mailmessage @email 
    } Catch {
        Write-Error "$(Get-Date -Format "HH:mm:ss"): Failed send-mailmessage attempt"
        Write-Error "$(Get-Date -Format "HH:mm:ss"): Error in $($_.InvocationInfo.ScriptName)."
        Write-Error "$(Get-Date -Format "HH:mm:ss"): -- Error information"
        Write-Error "$(Get-Date -Format "HH:mm:ss"): Line Number: $($_.InvocationInfo.ScriptLineNumber)"
        Write-Error "$(Get-Date -Format "HH:mm:ss"): Offset: $($_.InvocationInfo.OffsetInLine)"
        Write-Error "$(Get-Date -Format "HH:mm:ss"): Command: $($_.InvocationInfo.MyCommand)"
        Write-Error "$(Get-Date -Format "HH:mm:ss"): Line: $($_.InvocationInfo.Line)"
        Write-Error "$(Get-Date -Format "HH:mm:ss"): Error Details: $($_)"
    } ; # try/catch-E

    # then pipe just the errors out to console
    #if($error.count -gt 0){write-host $error }

}#*------------^ END Function Send-EmailNotif ^------------ ;  

#-------------------v Function Cleanup v-------------------
function Cleanup {
    # clear all objects and exit
    # Clear-item doesn't seem to work as a variable release 
   
    if(!(test-path function:Start-iseTranscript ) -OR !(test-path function:Stop-TranscriptLog ) -OR !(test-path function:Archive-Log) )
    {write-error "$((get-date).ToString("HH:mm:ss")):MISSING DEPENDENT FUNCTION:Test-Transcribing;Stop-TranscriptLog;Archive-Log" ; CleanUp ;} ; 
    
    # # 10:19 AM 12/16/2015 repl: get-timestamp() with explicit $(Get-Date -Format "HH:mm:ss")
    # 8:39 AM 12/10/2014 shifted to stop-transcriptLog function
    # 7:43 AM 1/24/2014 always stop the running transcript before exiting
    if ($bDebug) {"CLEANUP"}
    
    # 12:29 PM 12/16/2015 remove the report file
    if(test-path ($RptFile)){
        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Removing $($RptFile)..." ;
        remove-item -path ($RptFile)
    }
    if(test-path ($RptFile.replace(".html",".csv"))){
        write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Removing $($RptFile.replace(".html",".csv"))..." ;
        remove-item -path ($RptFile.replace(".html",".csv"))
    }
    
    
    #stop-transcript
    # 11:16 AM 1/14/2015 aha! does this return a value!??
    if($Host.Name -eq "Windows PowerShell ISE Host"){
        # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
        $logname= (join-path -path (join-path -path $ScriptDir -childpath "logs") -childpath ($ScriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
        write-host "`$logname: $logname";
        #Start-iseTranscript
        Start-iseTranscript -logname $logname ;
        #Archive-Log $logname ;
    } else {
        if($bDebug){ write-host -ForegroundColor Yellow "$(Get-Date -Format "HH:mm:ss"):Stop Transcript" };
        Stop-TranscriptLog ; 
        if($bDebug){ write-host -ForegroundColor Yellow "$(Get-Date -Format "HH:mm:ss"):Archive Transcript" };
        #Archive-Log $Transcript ; 
    } # if-E
    exit
} #*------------^ END Function Cleanup ^------------

# *** ENDREGION SETUP MARKER
#endregion SETUP

#*======^ END FUNCTIONS  ^======

#*------------v SUB MAIN v------------

#region LOAD

$sName="Microsoft.Exchange.Management.PowerShell.E2010"; if (!(Get-PSSnapin | where {$_.Name -eq $sName})) {Add-PSSnapin $sName -ea Stop};

# stock the Ex obj
$ExServer= get-exchangeserver -identity $env:computername -ea stop ; 

#====== V TRANSCRIPT-SUB MAIN- BOILERPLATE V======
# build outtransfile name: 
# leverage the global variable, that start-Transcript should leverage
$TimeStampNow = Get-Date -uformat "%Y%m%d-%H%M" ;
# name for the script & pass time
$Transcript = ((Split-Path -parent $MyInvocation.MyCommand.Definition) + "\logs\" + ([system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ) + "-" + $TimeStampNow + "-trans.log");
if ($bDebug) {write-host -foregroundcolor green "$(Get-Date -Format "HH:mm:ss"):`n`$Transcript:$Transcript";};
start-TranscriptLog $Transcript
#====== ^ OUTPUT FILE HANDLING BOILERPLATE (USE IN SUB MAIN) ^======

write-host ("$(Get-Date -Format "HH:mm:ss"):`nPASS STARTED" + ("="*5)) ;

# 1:57 PM 12/17/2015 add a manual delay for sub Win2012R2 OS's, SchtedTask startdelay of 3mins doesn't appear effective
if( (Get-WMIObject -class Win32_OperatingSystem).version -lt '6.3.9600'  ){
  if(!($bdebug)){
      # sub Win2012R2, delay here
      $sPollDelay=300 ; # # seconds to wait to check services
      write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Pre-Win2012R2 OS detected, waiting $($sPollDelay) seconds to check service status... " ;
      start-sleep -s 300  ; 
  } else {
      write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Debug detected, skipping startup check delay... " ;
  }; 
} else {
  # win2012R2 OS, the task startdelay should work...
};

#====== V HTML-EMAIL-SUB MAIN-BOILERPLATE V======
# validate dependency variables
if(!( ($ExServer) -AND ($ScriptNameNoExt) -AND ($ScriptBaseName)) ) {
    write-error "$((get-date).ToString("HH:mm:ss")):UNDEFINED `$ExServer($($ExServer)) or `$ScriptNameNoExt($($ScriptNameNoExt)) or `$ScriptBaseName($($ScriptBaseName)) variable. EXITING ";
    Cleanup
} # if-E
# leverage computername or mailboxrole to deter smtpserver
<# orig code
if ($env:COMPUTERNAME -eq 'MyComputer') {
    $SMTPServer = "Server0";
    $SMTPPort = 8111 ;
} elseif ($env:USERDOMAIN -eq "DOMAIN"){
    # lab
    write-host ("$(Get-Date -Format "HH:mm:ss"):DOMAIN: Using `$SMTPServer:$SMTPLab")
    $SMTPServer = $SMTPLab ;
} elseif ($ExServer.ismailboxserver){
    #$SMTPServer = "Server0";
    #1:08 PM 10/2/2017 grabbed from out of band copy:
    $SMTPServer = (get-transportserver "$(($env:computername).substring(0,3))*" | select -expand name | get-random );
      # 12:12 PM 10/2/2017 fixed typo: missing $
      write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Using SMTPServer:$($SMTPServer)" ;
}else{
  # hubs have access to vscan
  $SMTPServer = "vscan.DOMAIN.com" ;
} # if-E
#>
# 1:31 PM 10/2/2017 try dump in dyn code
if((get-exchangeserver $env:computername).IsHubTransportServer) { 
    $SMTPServer = "$($env:computername)" ; 
} else {
    $SMTPServer = (get-transportserver "$(($env:computername).substring(0,3))*" | select -expand name | get-random );
}; 

$TimeStampNow = get-date -uformat "%Y%m%d-%H%M" ;
#========
$SMTPFrom = (($ScriptBaseName.replace(".","-")) + "@DOMAIN.com") ; 
# 2:08 PM 3/20/2015
$SMTPSubj= ("POSSIBLE REBOOT?:$($env:COMPUTERNAME)($TimeStampNow)" ); 
Write-Host ((get-date).ToString("HH:mm:ss") + ":`$SMTPServer:" + $SMTPServer ) ;
$SMTPTo="todd.kadrie@DOMAIN.com" ;
$SmtpBody = @() ;
# (`n = CrLf in body)
#====== ^ HTML-EMAIL-SUB MAIN-BOILERPLATE ^======

# mailed-report version of the svcstatus script
$msg=("=" * 15) ;
write-host -foregroundcolor green $msg;

$msg="$(Get-Date -Format "HH:mm:ss"):`n`Checking Exchange Service Health on $ComputerName `n" ;  
write-host -foregroundcolor green $msg ; 

# 11:25 AM 12/14/2015 Lync svcs
#$svcs = (Get-CsWindowsService -computername $env:COMPUTERNAME | select name,status );
# Ex svcs
$svcs = (get-service -include *exchange* -computername $env:COMPUTERNAME | ?{$_.name -notmatch '^(wsbexchange)$'} | select name,DisplayName,status );
#$svcs | Format-Table -AutoSize | tee-object -FilePath $Rptfile -append ;  ;
# nope, tee -append is a PSv3 thing, use add-content
$svcs | Format-Table -AutoSize 
#*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
# collect data object into stock vari & assign a caption for this chunk of report
# use | out-string on the inbound object, if getting object references rather than text
$RptFragData = $svcs #| out-string; 
$RptFragCaption = "Exchange Service Health on $ComputerName" ; 
# note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
$RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
# append the fragment to the $RptContent aggregator 
if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
#*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^======

# Lync svcs code
#$DeadSvc = (Get-CsWindowsService -computername $FE | ?{$_.Status -ne 'Running'}  );
# Ex2010 svcs code
if($ExServer.ismailboxserver){
    #$DeadSvc = ($svcs | ?{$_.Status -ne 'Running'}  );
    $DeadSvc = ($svcs | ?{($_.Status -ne 'Running') -AND ($_.NAME -notmatch $rgxExMbxNonRunSvc)}  );
} else {
    # search is only supposed to be running on mbx
    # get-service -include *exchange* -computername $env:COMPUTERNAME | ?{$_.NAME -ne 'msftesql-Exchange'}
    $DeadSvc = ($svcs | ?{($_.Status -ne 'Running') -AND ($_.NAME -notmatch $rgxExHubNonRunSvc)}  );
}

if ($DeadSvc){    
    #write-host -foregroundcolor red "STOPPED SVCS";
    $msg="STOPPED SVCS" ;
    write-host -foregroundcolor green $msg;
    $DeadSvc | ft Name,Status
    #*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
    # collect data object into stock vari & assign a caption for this chunk of report
    # use | out-string on the inbound object, if getting object references rather than text
   $RptFragData = $DeadSvc | select Name,Status #| out-string ;
    $RptFragCaption = "STOPPED SVCS"  ; 
    # note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
    $RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
    # append the fragment to the $RptContent aggregator 
    if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
    #*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^======
    
    # 1:35 PM 10/2/2017 add remedial svc start
    $maxRepeat = 20 ; 
    foreach($svc in $DeadSvc){
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Starting Svc:$($svc.Name) and waiting for Running|Stopped Status" ;
        start-service -name $svc.Name -ea 0 ; 
        start-sleep -s 5 ; 
        Do { 
            write-host "." -NoNewLine;
            $maxRepeat-- ; 
            Start-Sleep -s 5
        } 
        Until ((get-service -name $svc.Name |?{$_.status -match '^(Running|Stopped)$'}) -OR ($maxRepeat -eq 0)) ; 
    } ; 
    # 2:06 PM 10/2/2017 re-pull curr status
    $svcs = (get-service -include *exchange* -computername $env:COMPUTERNAME | ?{$_.name -notmatch '^(Mimosa.*|wsbexchange)$'} | select name,DisplayName,status );
    if($ExServer.ismailboxserver){
        $DeadSvc = ($svcs | ?{($_.Status -ne 'Running') -AND ($_.NAME -notmatch $rgxExMbxNonRunSvc)}  );
    } else {
        # search is only supposed to be running on mbx
        # get-service -include *exchange* -computername $env:COMPUTERNAME | ?{$_.NAME -ne 'msftesql-Exchange'}
        $DeadSvc = ($svcs | ?{($_.Status -ne 'Running') -AND ($_.NAME -notmatch $rgxExHubNonRunSvc)}  );
    } ; 
    
    $bRet=$DeadSvc ; 
    #*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
    # collect data object into stock vari & assign a caption for this chunk of report
    # use | out-string on the inbound object, if getting object references rather than text
    if ($bRet) { 
        $RptFragData = $bRet #| out-string ; 
    } else { $RptFragData = "(no un-remediated/stopped services present) " }; 
    $RptFragCaption = "Post-Remediated Start Status:" ; 
    # note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
    $RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
    # append the fragment to the $RptContent aggregator 
    if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
    #*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^======
    
} else {
    #write-host "[no non-runnicng svcs (that should be running)]";
    # 11:33 AM 12/15/2015 add `n
    $msg= "[no non-runnicng svcs (that should be running)]`n";
    write-host -foregroundcolor green $msg;
    #*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
    # collect data object into stock vari & assign a caption for this chunk of report
    # use | out-string on the inbound object, if getting object references rather than text
    $RptFragData = " " #| out-string ; 
    $RptFragCaption = "[no non-runnicng svcs (that should be running)]" ;
    # note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
    $RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
    # append the fragment to the $RptContent aggregator 
    if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
    #*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^====== 
} ;


#write-host ("-" * 3);
#write-host ("=" * 15);
$msg=("=" * 15) ;
write-host -foregroundcolor green $msg;

#write-host -foregroundcolor green "$(Get-Date -Format "HH:mm:ss"):`n`Checking Test-ServiceHealth  on $ComputerName `n" ;
$msg= "$(Get-Date -Format "HH:mm:ss"):`n`Checking Test-ServiceHealth  on $ComputerName `n";
write-host -foregroundcolor green $msg;

#Test-ServiceHealth | tee-object -FilePath $Rptfile -append ; 
$bRet=Test-ServiceHealth 
# nope, tee -append is a PSv3 thing, use add-content
$bRet ; 
#*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
# collect data object into stock vari & assign a caption for this chunk of report
# use | out-string on the inbound object, if getting object references rather than text
#$RptFragData = $bRet #| out-string ; 
#12:00 PM 12/16/2015 issue with Test-ServiceHealth is that the ServicesRunning & ServicesNotRunning are multivalue
# they need to be joined manually
$RptFragData = $bRet | select Role,RequiredServicesRunning,@{n="ServicesRunning";e={[string]::join(", ",$_.ServicesRunning)}},@{n="ServicesNotRunning";e={[string]::join(", ",$_.ServicesNotRunning)}}
$RptFragCaption = "Test-ServiceHealth  on $ComputerName" ; 
# note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
$RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
# append the fragment to the $RptContent aggregator 
if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
#*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^====== 


#write-host ("=" * 15);
$msg=("=" * 15) ;
write-host -foregroundcolor green $msg;
if($ExServer.ismailboxserver){
      #write-host -foregroundcolor green "$(Get-Date -Format "HH:mm:ss"):`n`Checking Test-ReplicationHealth on $ComputerName `n" ;
      $msg="$(Get-Date -Format "HH:mm:ss"):`n`Checking Test-ReplicationHealth on $ComputerName `n" ;
      write-host -foregroundcolor green $msg;
      $bRet = Test-ReplicationHealth #| out-string  ; 
      $bRet ; 
      #*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
      # collect data object into stock vari & assign a caption for this chunk of report
      # use | out-string on the inbound object, if getting object references rather than text
      $RptFragData = $bRet | SELECT Server,Check,Result,Error #| out-string ; 
      $RptFragCaption = "Test-ReplicationHealth on $ComputerName" ; 
      # note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
      #$RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
      # 1:06 PM 12/16/2015 chg to list
      $RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
      # append the fragment to the $RptContent aggregator 
      if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
      #*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^====== 

      $msg=("=" * 15) ;
      # Get-MailboxDatabaseCopyStatus | sort name
      $msg="$(Get-Date -Format "HH:mm:ss"):`n`Checking Get-MailboxDatabaseCopyStatus on $ComputerName `n" ;
      write-host -foregroundcolor green $msg;
      #$bRet = Get-MailboxDatabaseCopyStatus | sort name #| out-string ;  
      # 1:12 PM 12/16/2015 subset select
      $bRet = Get-MailboxDatabaseCopyStatus | sort status,name | select Name,Status,CopyQueueLength,ReplayQueueLength,LastInspectedLogTime,ContentIndexState #| out-string ;  
      $bRet ; 
      #*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
      # use | out-string on the inbound object, if getting object references rather than text
      $RptFragData = $bRet #| out-string ; 
      $RptFragCaption = "Get-MailboxDatabaseCopyStatus on $ComputerName" ; 
      # note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
      $RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
      # append the fragment to the $RptContent aggregator 
      if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
      #*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^====== 
      
      $msg=("=" * 15) ;
      #write-host -foregroundcolor green "$(Get-Date -Format "HH:mm:ss"):`nChecking Mailstore Delivery Queue Status on $ComputerName (non-'Ready' queues)`n" ;
      $msg="$(Get-Date -Format "HH:mm:ss"):`nChecking Mailstore Hub Delivery Queues (non-Ready)`n" ;      
      write-host -foregroundcolor green $msg; 
      #9:58 AM 12/21/2015 pull local mailstore-named queues
      # 10:36 AM 12/21/2015 have to spec server names too, won't query local without local hub role
      # 7:53 AM 10/3/2017 add -ea 0 (move past timeouts)
      switch -regex ($env:computername.tostring().substring(0,3)){
         "(SITE|YYY)" {
            # throwing error access queues, from lab hub box, so cycle them all locally
            $hubs = get-transportserver |?{$_.name -match "^(YYY|SITE).*$"} ; 
            foreach($hub in $hubs){
                $bRet = get-queue -server $hub -ea 0 |?{($_.NextHopDomain.tolower() -match "^(?i:Server2})$") -AND ($_.identity -match "^(SITE|YYY).*") } |?{$_.Status -ne 'Ready'} 
            } ; 
        } ; 
         "XXX" {
            $hubs = get-transportserver |?{$_.name -match "^XXX.*$"} ; 
            foreach($hub in $hubs){
                $bRet = get-queue -server $hub -ea 0 |?{ ($_.NextHopDomain.tolower() -match "^(?i:Server2})$") -AND ($_.identity -match "^XXX.*")  } |?{$_.Status -ne 'Ready'} ; 
            } ; 
        }
         "SITE" {
            $hubs = get-transportserver |?{$_.name -match "^SITE.*$"} ; 
            foreach($hub in $hubs){
                $bRet = get-queue -server $hub -ea 0 |?{ ($_.NextHopDomain.tolower() -match "^(?i:Server2})$") -AND ($_.identity -match "^SITE.*") } |?{$_.Status -ne 'Ready'} ; 
            } ; 
        }
      } ;

      $bRet ; 
      #*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
      # collect data object into stock vari & assign a caption for this chunk of report
      # use | out-string on the inbound object, if getting object references rather than text
      if ($bRet) { 
          $RptFragData = $bRet #| out-string ; 
      } else { $RptFragData = "(no non-Ready MailStore queues present) " }  ; 
      $RptFragCaption = "MailStore non-'Ready' queues" ; 
      # note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
      $RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
      # append the fragment to the $RptContent aggregator 
      if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
      #*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^======
      
      # 2:20 PM 10/2/2017 run test-mapi on a random mbx per db
        $DAGName=(Get-MailboxServer $env:computername).DatabaseAvailabilityGroup ;
        $dbs = (Get-MailboxDatabase | where {$_.MasterServerOrAvailabilityGroup -eq $DAGName}) ;
        $bRet=$null ; $Result=$null ; 
        # aggregator
        $Output = @()
        foreach ($db in $dbs) {
            write-host -fore yellow $db.name ;
            $mbx = get-mailbox -database $db | get-random ;
            write-host -fore yellow ("Testing mbx:" + $mbx) ;
            if($host.version.major -ge 3){
                $Hash=[ordered]@{Dummy = $null ; } ;
            } else {$Hash = New-Object Collections.Specialized.OrderedDictionary ; } ;
            If($Hash.Contains("Dummy")){$Hash.remove("Dummy")} ; 
            if($mbx) { 
                $Result=($mbx | Test-MAPIConnectivity | select Mailbox,Server,Database,Result,Latency,Error) ; 
                $Hash.Add("Mailbox",$($Result.Mailbox)) ; 
                $Hash.Add("MailboxServer",$($Result.Server)) ; 
                $Hash.Add("Database",$($Result.Database)) ; 
                $Hash.Add("Result",$($Result.Result)) ; 
                $Hash.Add("Latency",$($Result.Latency)) ; 
                $Hash.Add("Error",$($Result.Error)) ; 
                $Output += New-Object PSObject -Property $Hash ; 
            } else {
                $Hash.Add("Mailbox","(empty db)") ; 
                $Hash.Add("MailboxServer",$($db.Server)) ; 
                $Hash.Add("Database",$($db.name)) ; 
                $Hash.Add("Result","(empty db)") ; 
                $Hash.Add("Latency",$null) ; 
                $Hash.Add("Error",$null) ; 
                $Output += New-Object PSObject -Property $Hash ; 
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$(($Output|out-string).trim())" ; 
                #$($db.name):(empty db)" ;
            }  ;
            (get-date).ToString("mm/dd/yyyy HH:mm:ss") ;
        };
        #$bRet=$Output ; 
        $bret=$Output | select mailbox,mailboxserver,database,result,latency,error | sort mailboxserver,database # | format-table -auto
        
        #*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
        # collect data object into stock vari & assign a caption for this chunk of report
        # use | out-string on the inbound object, if getting object references rather than text
        if ($bRet) {
            $RptFragData = $bRet #| out-string ;
        } else { $RptFragData = "(no non-Ready MailStore queues present) " }  ;
        $RptFragCaption = "Access-Confirm: Test-MAPIConnectivity a random mailbox in each db" ;
        # note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
        $RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ;
        # append the fragment to the $RptContent aggregator
        if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ;
        #*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^======

} ; 
if($ExServer.IsHubTransportServer){
      $msg=("=" * 15) ;
      $msg="$(Get-Date -Format "HH:mm:ss"):`nChecking Queue Status on $ComputerName (non-Ready)`n" ;
      write-host -foregroundcolor green $msg; 
      $bRet = get-queue -sortorder:-MessageCount |?{$_.status -notmatch '(Ready|Active)'} #| out-string ; 
      $bRet ; 
      #*======v HTML-EMAIL-REPORT-ASSY-BOILERPLATE v======
      # collect data object into stock vari & assign a caption for this chunk of report
      # use | out-string on the inbound object, if getting object references rather than text
      if ($bRet) { 
          $RptFragData = $bRet #| out-string ; 
      } else { $RptFragData = "(no non-Ready queues present) " }  ; 
      $RptFragCaption = "Queue Status on $ComputerName (non-'Ready' queues)" ; 
      # note: -As LIST creates a vertical report; -As TABLE creates a horizontal report;
      $RptFragment = $RptFragData | ConvertTo-Html -AS TABLE  -Fragment -PreContent "<h3>$RptFragCaption</h3>" | Out-String ; 
      # append the fragment to the $RptContent aggregator 
      if($RptContent.length){ $RptContent+=$RptFragment } else {$RptContent=$RptFragment} ; 
      #*======^ HTML-EMAIL-REPORT-ASSY-BOILERPLATE ^======
} ; 
$msg=("=" * 15) ;
write-host -foregroundcolor green $msg;


#endregion DataGather ; # -------

#*======v HTML-CSS-ASSY-BOILERPLATE v======
#region ReportAssemble ; # -------
<# specs of the fields supported by ConvertTo-Html (as a hashtable)
$HtmlFmt=@{ 
  As=[TABLE|LIST] ;
  Body=$PageBody # text to add after the opening <BODY>
  Title=$PageTitle ; 
  Head=$sHTMLhead # content for the <HEAD> tag; if Head usedthe -Title is ignored; 
  PreContent=$sHtmlPreLogo # text added before opening <TABLE>; 
  PostContent=$sHtmlPost} # text added after closing </TABLE>;
}; 
#>
<# 1:54 PM 3/24/2015 
Works: But you have to put your '$RptContent' of pre-Html-rendered Fragments in the PostContent. 
And if you want a footer, you have to concat $sHtmlFooter onto the tail of the $RptContent
#>
$HtmlFmt=@{ Title=$PageTitle ; 
  Head=$sHTMLhead ; 
  PreContent=$sHtmlPre ; 
  PostContent=$($RptContent + $sHtmlFooter) };
#ConvertTo-HTML @HtmlFmt | Set-Content ($FileBaseName + ".html") ;

# *** BREAKPOINT ;
ConvertTo-HTML @HtmlFmt | Set-Content ($Rptfile);

#endregion ReportAssemble ; # -------
#*======^ HTML-CSS-ASSY-BOILERPLATE ^======

#*======v HTML-EMAIL-ASSY-BOILERPLATE v======
# 12:13 PM 1/26/2015 mail a report
#Load as an attachment into the body text:
#$body = (Get-Content "path-to-file\file.html" ) | converto-html ;

if(!($BodyAsHtml)){
    # 2:19 PM 3/25/2015 original mailing body code
    $SmtpBody += ("Pass Completed "+ [System.DateTime]::Now + "`nResults Attached: " + $attachment) ;
    $SmtpBody += ('-'*50) ;
    #$SmtpBody += (gc $outtransfile | ConvertTo-Html) ;
    #
} else {
    # 2:19 PM 3/25/2015 shift the report into the msg body
    #ConvertTo-HTML @HtmlFmt | Set-Content ($Rptfile);
    $SmtpBody = ConvertTo-HTML @HtmlFmt |Out-String  ;
} # if-E
#*======^ HTML-EMAIL-ASSY-BOILERPLATE ^======

Send-EmailNotif

Cleanup; 
# *** REGION REPORT END MARKER
#endregion REPORT 
# ============



