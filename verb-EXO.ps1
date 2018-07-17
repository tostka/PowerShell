# verb-EXO.ps1
#           ===========
#*======v VERB-EXO SET v======
<# Versions: # 9:24 PM 7/16/2018 broad cleanup & tightening
# 9:04 PM 7/11/2018 synced to MyAccountsid-incl-ServerApp.ps1
#>
<##-=-=-distribute to all Exch servers=-=-=-=-=-=
[array]$files = (gci -path "\\$env:COMPUTERNAME\c$\path-to\scripts\verb-EXO.ps1" | ?{$_.Name -match "(?i:(.*\.(PS1|CMD)))$" })  ;[array]$srvrs = get-exchangeserver | ?{(($_.IsMailboxServer) -OR ($_.IsHubTransportServer))} | select  @{Name='COMPUTER';Expression={$_.Name }} ;$srvrs = $srvrs|?{$_.computer -ne $($env:COMPUTERNAME) } ; $srvrs | % { write-host "$($_.computer)" ; copy $files -Destination \\$($_.computer)\c$\path-to\scripts\ -whatif ; } ; get-date ;
#-=-=-=-=-=-=-=-=
#>

# values from central cfg 
if(!$DoRetries){$DoRetries = 4 ; } ;          # attempt retries
if(!$RetrySleep){$RetrySleep = 5 ; }          # mid-retry sleep in secs
if(!$retryLimit){[int]$retryLimit=1; }        # just one retry to patch lineuri duped users and retry 1x
if(!$retryDelay){[int]$retryDelay=20; }       # secs wait time after failure
if(!$abortPassLimit){$abortPassLimit = 4;}    # maximum failed users to abort entire pass

#*------v Function Reconnect-EXO v------
Function Reconnect-EXO {
    <# 
    .SYNOPSIS
    Reconnect-EXO - Test and reestablish PSS to https://ps.outlook.com/powershell/
    .NOTES
    Updated By: : Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	@tostka http://twitter.com/tostka
    Based on original function written by: ExactMike Perficient, Global Knowl... (Partner)  
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Change Log
    * 8:04 AM 11/20/2017 code in a loop in the reconnect-exo, until it hits or 100% times out
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 7:58 AM 3/15/2017 ren Disconnect/Connect/Reconnect-EXO => Disconnect/Connect/Reconnect-EXO, added pss, doing tweaks to put into prod use
    * 2/10/14 posted version 
    .DESCRIPTION
    I use this for routine test/reconnect of EXO. His orig use was within batches, to break up and requeue chunks of commands.
    Mike's original comment: Below is one 
    example of how I batch items for processing and use the 
    Reconnect-EXO function.  I'm still experimenting with how to best 
    batch items and you can see here I'm using a combination of larger batches for 
    Write-Progress and actually handling each individual item within the 
    foreach-object script block.  I was driven to this because disconnections 
    happen so often/so unpredictably in my current customer's environment: 
    
    #-=-=Batch sample-=-=-=-=-=-=
    $batchsize = 10 ;
    $RecordCount=$mr.count #this is the array of whatever you are processing ;
    $b=0 #this is the initialization of a variable used in the do until loop below ;
    $mrs = @() ;
    do {
        Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
        $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-EXO; $_ | Get-OLMoveRequestStatistics} ;
        $b=$b+$batchsize ;
        } ;
    until ($b -gt $RecordCount) ; 
    #-=-=-=-=-=-=-=-=
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Reconnect-EXO; 
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    
    # 8:09 AM 11/20/2017 fault tolerant looping exo connect, don't let it exit until a connection is present, and stable, or return error for hard time out
    $tryNo=0 ; 
    Do {
        $tryNo++ ; 
        write-host "." -NoNewLine;Start-Sleep -m (1000 * 5)
        if ($EOLSession.state -eq 'Broken' -or !$EOLSession) {Disconnect-EXO; Start-Sleep -Seconds 3; Connect-EXO} ;     
        if( !(Get-PSSession|?{($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}) ){
            Reconnect-EXO ; 
        }  ; 
        if($tryNo -gt $DoRetries ){throw "RETRIED EXO CONNECT $($tryNo) TIMES, ABORTING!" } ;
    } Until ((Get-PSSession |?{$_.ComputerName -match $rgxExoPsHostName -AND $_.State -eq "Opened" -AND $_.Availability -eq "Available"}))

}#*------^ END Function Reconnect-EXO ^------
if(!(get-alias | ?{$_.name -like "rxo"})) {Set-Alias 'rxo' -Value 'Reconnect-EXO' ; } ;
#*------v Function Connect-EXO v------
Function Connect-EXO {
    <# 
    .SYNOPSIS
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .NOTES
    Written By: : Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	@tostka http://twitter.com/tostka
    Based on 'overlapping functions' concept by: ExactMike Perficient, Global Knowl... (Partner)  
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Change Log
    # 8:22 AM 11/20/2017 tried splicing in retry loop into reconnect-exo as well, to avoid using any state testing in scripts, localize it 1x here.
    # 1:49 PM 11/2/2017 coded around non-profile gaps from not having get-admincred() - added the prompt code in to fake it 
    # 12:26 PM 9/11/2017 debugged retry - catch doesn't fire properly on new-Pssession, have to test the $error state, to detect auth fails (assuming the bad pw error# is specific). $error test is currently targeting specific error returned on a bad password. Added retry, for when connection won't hold and fails breaks - need to watch out that bad pw doesn't lock out the acct!
    # 12:50 PM 6/2/2017 expanded pshelp, added Examples, cleaned up deadwood
    * # 12:10 PM 3/15/2017 Connect-EXO typo, disable prefix auto spec, unless actually blanked. switch ProxyEnabled to non-Mandatory.
    .DESCRIPTION
    .PARAMETER  ProxyEnabled
    Use Proxy-Aware SessionOption settings [-ProxyEnabled]
    .PARAMETER  CommandPrefix
    [noun]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix exolab ]
    .PARAMETER  Credential
    Credential to use for this connection [-credential 's-todd.kadrie@DOMAIN.com']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    connect-exo
    Connect using defaults, and leverage any pre-set $global:o365cred variable
    .EXAMPLE
    connect-exo -CommandPrefix exo -credential (Get-Credential -credential s-todd.kadrie@DOMAIN.com)  ; 
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    *---^ END Comment-based Help  ^--- #>

    Param(
        [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")][boolean]$ProxyEnabled = $False,  
        [Parameter(HelpMessage="[noun]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix exolab]")][string]$CommandPrefix = 'exo',
        [Parameter(HelpMessage="Credential to use for this connection [-credential 's-todd.kadrie@DOMAIN.com']")]$Credential = $global:o365cred
    ) ; 
    # 12:10 PM 3/15/2017 disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
    if(!$CommandPrefix){ $CommandPrefix='exo' ; } ; 
    
    
    # toggle whether to use federated or non-federated (CO) accts for prompts
    if(!$global:bUseo365COAdminUID){$bUseo365COAdminUID = $false} ; 
    # federated admin acct UPN
    if(!$global:o365AdmUid){$o365AdmUid = "s-todd.kadrie@DOMAIN.com" };
    # non-federated admin acct UPN
    if(!$global:o365COAdmUid){$o365COAdmUid="c-todd.kadrie@DOMAINco.onmicrosoft.com"} ; 
    # federated admin UPN tenant2
    if(!$global:o365LabAdmUid){$o365LabAdmUid = "Admino365cloud@DOMAIN.onmicrosoft.com"} ; 
    # non-federated admin acct UPN tenant2
    if(!$global:o365LabCOAdmUid){$o365LabCOAdmUid="Admino365cloud@DOMAIN.onmicrosoft.com"} ; # non-federated admin acct UPN
    
    $EXOsplat=@{
        ConfigurationName="Microsoft.Exchange" ;
        ConnectionUri="https://ps.outlook.com/powershell/" ;
        Authentication="Basic" ;
        AllowRedirection=$true;
    } ; 

    # obtain o365cred if not already loaded
    if(!$global:o365cred){
        # 1:39 PM 11/2/2017 use Get-AdminCred, if it's present, if not do it manual
        if(test-path function:\get-admincred) { 
            Get-AdminCred ; $EXOsplat.Add("Credential",$global:o365cred);
        } else {
            switch($env:USERDOMAIN){
               "DOMAIN1" { 
                  write-host -foregroundcolor yellow "PROMPTING FOR O365 CRED ($($o365AdmUid ))" ; 
                  if(!$bUseo365COAdminUID){
                      if($o365AdmUid ){$global:o365cred = Get-Credential -Credential $o365AdmUid } else { $global:o365cred = Get-Credential } ; 
                  } else {
                      if($o365COAdmUid){global:o365cred = Get-Credential -Credential $o365COAdmUid} else { $global:o365cred = Get-Credential } ; 
                  } ; 
                }
                "DOMAIN2" { 
                    write-host -foregroundcolor yellow "PROMPTING FOR O365 CRED ($($o365LabAdmUid ))" ; 
                    if(!$bUseo365COAdminUID){
                        if($o365LabAdmUid){$global:o365cred = Get-Credential -Credential $o365LabAdmUid} else { $global:o365cred = Get-Credential } ; 
                    } else {
                        if($o365LabCOAdmUid){$global:o365cred = Get-Credential -Credential $o365LabCOAdmUid} else { $global:o365cred = Get-Credential } ; 
                    } ; 
                }
                default {
                    write-host -foregroundcolor yellow "$($env:USERDOMAIN) IS AN UNKNOWN DOMAIN`nPROMPTING FOR O365 CRED:" ; 
                    $global:o365cred = Get-Credential 
                } ; 
            } ; 
            # add cred to the splat
            $EXOsplat.Add("Credential",$global:o365cred);
        }  ; 
    } else {
        $EXOsplat.Add("Credential",$global:o365cred);
    } ; 
    # Show where connected...
    Add-PSTitleBar 'EXO' ;
    # set color scheme to White text on Black
    #$HOST.UI.RawUI.BackgroundColor = "Black" ; $HOST.UI.RawUI.ForegroundColor = "White" ;
    If ($ProxyEnabled) {
        $EXOsplat.Add("sessionOption",$(New-PsSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic));
    	Write-Host "Connecting to Exchange Online via Proxy"  ; 
    } Else {
    	Write-Host "Connecting to Exchange Online"  ; 
    } ; 
    # 11:05 AM 9/11/2017 monitor-EXOMigr.ps1 keeps crapping out on connects, put in a retry process...
    $Exit = 0 ; # zero out $exit each new cmd try/retried
    # do loop until up to 4 retries...
    Do {
        $error.clear() ;
        Try {
            $Global:EOLSession = New-PSSession @EXOsplat ; 
            # 11:41 AM 9/11/2017 fails above don't actually trigger catch, you have to test the $error for $error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed'
            if($error.count -ne 0) {
                if($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed'){
                    write-warning "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    EXIT ; 
                } ; 
            } ; 
            $Global:EOLModule = Import-Module (Import-PSSession $Global:EOLSession -Prefix $CommandPrefix -DisableNameChecking -AllowClobber) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ; 
            <# on a bad pw in the new-pssession, you get back:
            #-=-=-=-=-=-=-=-=
            New-PSSession : [outlook.office365.com] Connecting to remote server outlook.office365.com failed with the following
            error message : [ClientAccessServer=SN4PR0501CA0130,BackEndServer=cy1pr04mb1963.namprd04.prod.outlook.com,RequestId=bf1
            f7979-b685-4179-8a09-f2dfedfb9a91,TimeStamp=9/11/2017 4:22:11 PM] Access Denied For more information, see the
            about_Remote_Troubleshooting Help topic.
            At line:1 char:1
            + New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outl ...
            + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                + CategoryInfo          : OpenError: (System.Manageme....RemoteRunspace:RemoteRunspace) [New-PSSession], PSRemotin
               gTransportException
                + FullyQualifiedErrorId : -2144108477,PSSessionOpenFailed
            #-=-=-=-=-=-=-=-=
            #-=-=-=-=-=-=-=-=
            +[MyAccount]::[PS]:C:\u\w\e\scripts$ $error[0] | select *
            writeErrorStream      : True
            PSMessageDetails      :
            Exception             : System.Management.Automation.Remoting.PSRemotingTransportException: Connecting to remote
                                    server outlook.office365.com failed with the following error message : [ClientAccessServer=SN4P
                                    R0501CA0130,BackEndServer=cy1pr04mb1963.namprd04.prod.outlook.com,RequestId=bf1f7979-b685-4179-
                                    8a09-f2dfedfb9a91,TimeStamp=9/11/2017 4:22:11 PM] Access Denied For more information, see the
                                    about_Remote_Troubleshooting Help topic.
            TargetObject          : System.Management.Automation.RemoteRunspace
            CategoryInfo          : OpenError: (System.Manageme....RemoteRunspace:RemoteRunspace) [New-PSSession],
                                    PSRemotingTransportException
            FullyQualifiedErrorId : -2144108477,PSSessionOpenFailed
            ErrorDetails          : [outlook.office365.com] Connecting to remote server outlook.office365.com failed with the
                                    following error message : [ClientAccessServer=SN4PR0501CA0130,BackEndServer=cy1pr04mb1963.nampr
                                    d04.prod.outlook.com,RequestId=bf1f7979-b685-4179-8a09-f2dfedfb9a91,TimeStamp=9/11/2017
                                    4:22:11 PM] Access Denied For more information, see the about_Remote_Troubleshooting Help
                                    topic.
            InvocationInfo        : System.Management.Automation.InvocationInfo
            ScriptStackTrace      : at <ScriptBlock>, <No file>: line 1
            PipelineIterationInfo : {0, 1}
            #-=-=-=-=-=-=-=-=
            #>
             # break-exit here, completes the Until block
            $Exit = $Retries ; 
        } Catch {            
            # pause to give it time to reset
            Start-Sleep -Seconds $RetrySleep ; 
            $Exit ++ ; 
            Write-Verbose "Failed to exec cmd because: $($Error[0])" ; 
            Write-Verbose "Try #: $Exit" ; 
            If ($Exit -eq $Retries) {Write-Warning "Unable to exec cmd!"} ; 
        } # try-E
    } Until ((Get-PSSession |?{$_.ComputerName -match $rgxExoPsHostName -AND $_.State -eq "Opened" -AND $_.Availability -eq "Available"}) -or ($Exit -eq $Retries) ) # loop-E    
} ; #*------^ END Function Connect-EXO ^------
if(!(get-alias | ?{$_.name -like "cxo"})) {Set-Alias 'cxo' -Value 'Connect-EXO' ; } ;
#*------v Function Disconnect-EXO v------
Function Disconnect-EXO {
    <# 
    .SYNOPSIS
    Disconnect-EXO - Disconnects any PSS to https://ps.outlook.com/powershell/ (cleans up session after a batch or other temp work is done)
    .NOTES
    Updated By: Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	@tostka http://twitter.com/tostka
    Based on original function written by:  ExactMike Perficient, Global Knowl... (Partner)  
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Change Log
    # 11:23 AM 7/10/2018: made exo-only (was overlapping with CCMS)
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 8:49 AM 3/15/2017 Disconnect-EXO: add Remove-PSTitleBar 'EXO' to clean up on disconnect
    * 2/10/14 posted version 
    .DESCRIPTION
    I use this to smoothly cleanup connections. 
    Mike's original notes: 
    The function Disconnect-EXO gets used within batches.  Below is one 
    example of how I batch items for processing and use the 
    Disconnect-EXO function.  
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Disconnect-EXO; 
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    *---^ END Comment-based Help  ^--- #>
    # 9:25 AM 3/21/2017 getting undefined on the below, pretest them
    if($Global:EOLModule){$Global:EOLModule | Remove-Module -Force ; } ; 
    if($Global:EOLSession){$Global:EOLSession | Remove-PSSession ; } ; 
    Get-PSSession | Where-Object {$_.ComputerName -match '^ps\.outlook\.com$'} | Remove-PSSession ; 
    Remove-PSTitlebar 'EXO' ; 
} ; #*------^ END Function Disconnect-EXO ^------
if(!(get-alias | ?{$_.name -like "dxo"})) {Set-Alias 'dxo' -Value 'Disconnect-EXO' ; } ;
#*======^ END END VERB-EXO SET ^======
#           ===========

