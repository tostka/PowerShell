# AdGroupMembersRecurse.ps1
# prev named: get-AdGroupMembersRecurse.ps1
# debug command: Clear-Host ; AdGroupMembersRecurse.ps1 -Groups "Group1","Group2" ; 
  <# 
  .SYNOPSIS
  AdGroupMembersRecurse.ps1 - Recursive member lookup, redirects by domain
  .NOTES
  Written By: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka
  Change Log
  [VERSIONS]
  * 1:34 PM 6/20/2016 ren: AdGroupMembersRecurse.ps1 => AdGroupMembersRecurse.ps1 (was constantly conflicting in autotype of get-adgroup etc).
  * 12:43 PM 2/16/2016 added CreateAcronymFromCaps()
  * 12:41 PM 2/16/2016 pulled out the manual get-aduser - the data is already in the get-adgroupmember output (incl's name,DN,samaccountname etc).
  * 12:23 PM 12/11/2015 rewrote as a pipelined function, seems to be working for small groups.
  * 11:29 AM 12/11/2015 - never completed, nother req, lets do it now
  * 7:52 AM 10/26/2015 - initial build
  .DESCRIPTION
  AdGroupMembersRecurse.ps1 - Recursive member lookup, redirects by domain
  Uses get-adgroup | get-adgroupmember -recursive to lookup full group membership. Includes detect and redir of CN user qrys to CN dc's.
  .PARAMETER  Groups
  Specify AD Group(s) to have membership recursively reported to csv (multiple should be comma-delimted)
  .PARAMETER  NoCSV
  Suppress CSV Export (console-only)[-NoCSV]
  .PARAMETER  ShowDebug
  Switch to output Debugging messages[-ShowDebug]
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  Outputs reports to CSV file (one per group specified in groups)
  .EXAMPLE
  .\AdGroupMembersRecurse.ps1 -groups "Lyn-App-DataScan-G","Lyn-App-DataScan Admin-G"
  Process two groups
  .EXAMPLE
  .\AdGroupMembersRecurse.ps1 -groups "Group1","Group2"
  Process two very large groups
  .LINK
  *---^ END Comment-based Help  ^--- #>

# 11:31 AM 12/11/2015 add requires for ADMS & PSv3
<# #Requires –Version 3  disabled, server etc are on PSv2
#>

[CmdletBinding()]
#Requires –Modules ActiveDirectory
Param(
  [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelinebyPropertyName=$True,
    HelpMessage='Specify AD Group(s) to have membership recursively reported to csv (multiple should be comma-quote-delimted)')]
  $Groups,
  [Parameter(HelpMessage='Suppress CSV Export (console-only)[-NoCSV]')]
  [switch] $NoCSV,
  [Parameter(HelpMessage='Switch to output Debugging messages[-ShowDebug]')]
  [switch] $ShowDebug
);

# pick up the bDebug from the $ShowDebug switch parameter
if ($ShowDebug) {$bDebug=$true};

if($host.version.major -lt 3){
    $mName="ActiveDirectory"; if (!(Get-Module | where {$_.Name -eq $mName})) {Import-Module $mName -ErrorAction Stop ;} ;
} ; 
# 2:10 PM 2/4/2015 shifted to here to accommodate include locations
$scriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
# 2:12 PM 3/25/2015 moved below to script head to make avail for html footer boilerplate
# fr 2266
$scriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName)) ;
# fr 2248
$scriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
$timeStampNow = get-date -uformat "%Y%m%d-%H%M" ;

# 2:01 PM 2/5/2015
$sQot = [char]34
$sQotS = [char]39

#*======^ END BOILERPLATE SCRIPT-REFERENCES ^======

#*================v FUNCTIONS  v================

#*------v Function AdGroupMembersRecurse v------
Function AdGroupMembersRecurse {
    <# 
      .SYNOPSIS
      AdGroupMembersRecurse() - Recursive member lookup, redirects by domain
      .NOTES
      Written By: Todd Kadrie
      Website:	http://tinstoys.blogspot.com
      Twitter:	http://twitter.com/tostka
      Change Log
      [VERSIONS]
      * 1:34 PM 6/20/2016 ren: Get-AdGroupMembersRecurse() => AdGroupMembersRecurse.ps1 (was constantly conflicting in autotype of get-adgroup etc).
      * 11:29 AM 12/11/2015 - never completed, nother req, lets do it now
      * 7:52 AM 10/26/2015 - initial build
      .DESCRIPTION
      AdGroupMembersRecurse() - Recursive member lookup, redirects by domain
      Uses get-adgroup | get-adgroupmember -recursive to lookup full group membership. Includes detect and redir of CN user qrys to CN dc's.
      .PARAMETER  Groups
      Specify AD Group(s) to have membership recursively reported to csv (multiple should be comma-delimted)
      .PARAMETER  NoCSV
      Suppress CSV Export (console-only)[-NoCSV]
      .PARAMETER  ShowDebug
      Switch to output Debugging messages[-ShowDebug]
      .INPUTS
      None. Does not accepted piped input.
      .OUTPUTS
      Outputs reports to CSV file (one per group specified in groups)
      .EXAMPLE
      .\AdGroupMembersRecurse.ps1 -groups "Lyn-App-DataScan-G","Lyn-App-DataScan Admin-G" -ShowDebug
      .LINK
      *---^ END Comment-based Help  ^--- #>

    # 11:31 AM 12/11/2015 add requires for ADMS & PSv3

    Param(
      [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelinebyPropertyName=$True,
        HelpMessage='Specify AD Group(s) to have membership recursively reported to csv (multiple should be comma-quote-delimted)')]
      $Groups,
      [Parameter(HelpMessage='Suppress CSV Export (console-only)[-NoCSV]')]
      [switch] $NoCSV,
      [Parameter(HelpMessage='Switch to output Debugging messages[-ShowDebug]')]
      [switch] $ShowDebug
    );

    BEGIN {
        # 2:04 PM 12/11/2015 
        $attachment = @()  ; 
    }  # BEG-E
    # $continue = $true
    PROCESS {
  
        foreach ($grp in $Groups) {
            try {
                $sw = [Diagnostics.Stopwatch]::StartNew();
                write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):===running: $grp===" ; 
                #$fn=".\$($grp.tostring().replace(" ","-"))-RECURSIVEMEMBERS-$(get-date -uformat '%Y%m%d-%H%M').csv" ; 
                # 7:43 AM 2/16/2016 make it descriptive: $scriptNameNoExt, build an Acronym on the capital letters of the scriptname
                #$fn=".\$($grp.tostring().replace(" ","-"))-RECURSIVEMEMBERS-$($scriptNameNoExt -split "" -cmatch '([A-Z])' -join "")-$(get-date -uformat '%Y%m%d-%H%M').csv" ; 
                # create-AcronymFromCaps
                $fn=".\$($grp.tostring().replace(" ","-"))-RECURSIVEMEMBERS-$(create-AcronymFromCaps $scriptNameNoExt)-$(get-date -uformat '%Y%m%d-%H%M').csv" ; 

                # 8:54 AM 2/16/2016 above gadu isn't necessary, the raw groupmember output includes name,distname,& samacctname already.
                Get-ADGroup -identity $grp | Get-ADGroupMember -Recursive | select SamAccountName,name,distinguishedname | export-csv $fn -notype ; 
                $sw.Stop() ; 
                #write-host -foregroundcolor green "Elapsed Time: (HH:MM:SS.ms)" $sw.Elapsed.ToString() ;
                write-host -foregroundcolor green "Elapsed Time: "  ($sw.Elapsed.Days.ToString() + "d " + $sw.Elapsed.Hours.ToString() + "h " + $sw.Elapsed.Minutes.ToString() + "m " + $sw.Elapsed.Seconds.ToString() + "s " + $sw.Elapsed.Milliseconds.ToString() + "ms")
                if(test-path $fn){
                    # accumulate attachments for send-mailmessage
                    $attachment += $fn ; 
                    $iCount = (import-csv $fn | measure).count ; 
                    write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")): $($iCount) matching users exported to:$($fn)" ;
                    # 2:03 PM 12/11/2015 looks like the $attachments in here isn't updatting outsid, so lets wrap it and export it to pl
                    

                } else {
                    write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):No Results output" ;
                } ; 
            } catch {
                Write-Error "$(Get-TimeStamp): FAILURE!" ;
                # opt extended error info...
                Write-Error "$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName)." ; 
                Write-Error "$(Get-TimeStamp): -- Error information" ;
                Write-Error "$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)" ;
                Write-Error "$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)" ;
                Write-Error "$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)" ;

                Write-Error "$(Get-TimeStamp): Line: $($_.InvocationInfo.Line)" ;
                Write-Error "$(Get-TimeStamp): Error Details: $($_)"
            } # try/cat-E
        } # loop-E
    } # Proc-E
    END {
        if( ($attachment | measure).count -gt 0){
            $RetHash= @{
                Attachments=$attachment;
                Results=$true ; 
            } ; 
        } else {
            $RetHash= @{
               Attachments=$null;
               Results=$false
            } ; 
        } # if-E 
        # dumps $RetHash into the pipeline, returns at call point
        Write-Output $RetHash ;
    } 
}  ; #*------^ END Function AdGroupMembersRecurse ^------

#*------v Function create-AcronymFromCaps()  v------
Function create-AcronymFromCaps {
    <# 
    .SYNOPSIS
    create-AcronymFromCaps - Creates an Acroynm From string specified, by extracting only the Capital letters from the string
    .NOTES
    Written By: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka
    Change Log
    [VERSIONS]
    12:14 PM 2/16/2016 - working
    8:58 AM 2/16/2016 - initial version
    .DESCRIPTION
    create-AcronymFromCaps - Creates an Acroynm From string specified, by extracting only the Capital letters from the string
    .PARAMETER  String
    String to be convered to a 'Capital Acrynym'
    .INPUTS
    Accepts piped input.
    .OUTPUTS
    Returns a string of the generated Acronym into the pipeline
    .EXAMPLE
    create-AcronymFromCaps "AdGroupMembersRecurse" ; 
    Create an Capital-letter Acroynm for the specified string
    .EXAMPLE
    $fn=".\$(create-AcronymFromCaps $scriptNameNoExt)-$(get-date -uformat '%Y%m%d-%H%M').csv" ; 
    Create a filename based off of an Acronym from the capital letters in the ScriptNameNoExt.
    .LINK
    *---^ END Comment-based Help  ^--- #>

    Param(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="String to be convered to a 'Capital Acrynym'[string]")][ValidateNotNullOrEmpty()]
        [string]$String        
    ) # PARAM BLOCK END

    #"AdGroupMembersRecurse" -split "" -cmatch '([A-Z])' -join "" ;
    $AcroCap = $String  -split "" -cmatch '([A-Z])' -join ""  ; 
    # drop it back into the pipeline
    write-output $AcroCap ; 
} #*------^ END Function create-AcronymFromCaps() ^------ ; 

#*================^ END FUNCTIONS  ^================

#*----------------v SUB MAIN v----------------
# 11:14 AM 4/2/2015 start banner
Write-Verbose "---------------------" -Verbose:$verbose
Write-Verbose "START ==== $($scriptBaseName) ====" -Verbose:$verbose

#region SETUP
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
$TimeStampNow = get-date -uformat "%Y%m%d-%H%M"
$Transcript=$ScriptDir + "logs"
if (!(test-path $Transcript)) {Write-Host "Creating dir: $Transcript" ;mkdir $Transcript ;} ;
$Transcript+="\" + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.log" ;
Trap {Continue} Stop-Transcript ;
start-transcript -path $Transcript ;

# attachments hash
#$attachment = @()  ; 

if($NoCSV){
    $bRsults = AdGroupMembersRecurse -groups $Groups -NoCSV -ShowDEbug ; 
} else {
    $bRsults = AdGroupMembersRecurse -groups $Groups -ShowDEbug ; 
} ; 

stop-transcript
<# Attachments=$null;
  Results#>

# Load an attachment into the body text:
if(test-path $Transcript) {
  write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Adding Transcript to Body $($Transcript)" ;
  $body = (Get-Content $Transcript ) | converto-html ;
} else {
  $body = "Script output attached" ;
}  ; 


$email=@{
   To="Todd.kadrie@domain.com" ; 
   From="$($ScriptNameNoExt)@domain.com" ; 
   Subject="Results:$($ScriptNameNoExt)-$($TimeStampNow)" ; 
   BodyAsHTML=$true ; 
   Body=($Body | out-string); 
   SMTPServer = "server" ; 
   Verbose = $Verbose;
} ; 
# 2:17 PM 12/11/2015: port isn't supported on psv2
if($host.version.major -gt 2){
    $email.add("Port","25");
}


<# 1:48 PM 12/11/2015 simplied below works
$email=@{
   To="Todd.kadrie@domain.com" ; 
   From="$($ScriptNameNoExt)@domain.com" ; 
   Subject="Results:$($ScriptNameNoExt)-$($TimeStampNow)" ; 
   SMTPServer = "server" ; 
   Verbose = $Verbose;
} ; 
#>
if($bRsults.attachments){
    write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Adding Attachement" ;
    $email.add("Attachments",$bRsults.attachments);
} ; 
if(($env:COMPUTERNAME) -EQ "LYN-3V6KSY1"){$email.Port = 8111 }

# 12:48 PM 12/11/2015 ps2 permits pipeling collection of files into send-mailmsg and they are included as attachments
write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Sending Report Email..." ;
#"$($email | out-string)" ; 
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


#$attachment | send-mailmessage @email ; 
try { send-mailmessage @email -ErrorAction Stop } 
Catch {
    # stick a pause in here, if it's ConnectionTimeout'ing on Exch, you need to give it time to reset
    #Start-Sleep -Seconds 10
    $Exit ++ ; 
    Write-Error -verbose:$true "Failed to send message because: $($Error[0])" ; 
    Write-Error -verbose:$true "Try #: $Exit" ; 
    If ($Exit -eq $Retries) {Write-Warning "Unable to send message!"} ; 
} # try-E
# dump a notice anyway
#Send-MailMessage -SmtpServer server -From AdGroupMembersRecurse.ps1@domain.com  -To todd.kadrie@domain.com -Subject "Script Results"

# *** 2:19 PM 12/11/2015 it's all working now except tucking the $transcript into the msgbody *** 


Write-Verbose "END ==== $($scriptBaseName) ====" -Verbose:$verbose
#*----------------^ END Function SUB MAIN ^----------------

<# #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$Msgs = (Get-ExchangeServer | where { $_.isHubTransportServer -eq $true -and $_.Site -like '*SiteName*' -and $_.admindisplayversion.major -eq 14} | get-messagetrackinglog -resultsize unlimited -Sender "AdGroupMembersRecurse@domain.com" -Start "$(([datetime]::Now.AddMinutes(-5)))"  ) ; $Msgs | ?{$_.connectorid –like '*\SMTP/VSCAN Relay to Ex2010'} | select Timestamp,Source,EventId,RelatedRecipientAddress,Sender,{$_.Recipients},RecipientCount,{$_.RecipientStatus},MessageSubject,TotalBytes,{$_.Reference},MessageLatency,MessageLatencyType,InternalMessageId,MessageId,ReturnPath,ClientIp,ClientHostname,ServerIp,ServerHostname,ConnectorId,SourceContext,MessageInfo,{$_.EventData} | export-csv -notype .\logs\TICKET-UID-XXX-EvtRcv-TIMESTMP.csv  ;
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

#>
