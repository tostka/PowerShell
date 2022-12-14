# verb-o365-gist-pub.ps1
## ======== V O365 EXO ETC FUNCS() V ===
# stuff your Office 365 admin account credentials into variables for unprompted reuse
<# the following are Admin (UID) and non-admin (LUA) credential objects for different tenants and logons
I keep their definitions in another profile-level pre-loaded module.
$o365AdmUid = "logon@domain.com" ; # Tenant 1 primary admin logon UPN
$o365LabAdmUid = "logon@domainlab.com" ;  # # Tenant 2 primary admin logon UPN
$o365COAdmUid="logon@tenant1.onmicrosoft.com" ; # Tenant 1 optional cloud-only admin acct (for backup in case of loss of federated access)
$o365LabCOAdmUid="logon@tenant2.onmicrosoft.com" ; Tenant 2 cloud-only admin acct

# flag that switches from federated/broken SID acct to cloud only $o365COAdmUid
$bUseo365COAdminUID = $false ; 

# You can stock the above with credential objects via:
$variable=get-credential
# or even pull from xml:
$variable= Import-CLIXML c:\path-to\credential.xml
# after storing the credential in xml:
Get-Credential |  Export-Clixml c:\path-to\credential.xml 
# if using the xml process above, they're authenticated via Windows data protection API, and the key used to encrypt the password is specific to both the user and the machine that the code is running under. Completely user logon-specific and non-portable across machines (e.g. you have to recreate on each new machine for each logon to use the settings).
#>
Function Get-O365AdminCred { 
    If (-not($o365cred)) { 
        #$global:o365cred = Get-Credential -Credential $o365AdmUid 
        # leverage external function
        Get-AdminCred ; 
    } 
}#*------^ END Function Get-O365AdminCred ^------

#           ===========
#*======v VERB-EXO SET v======
<# set of functions that connect/reconnect/disconnect with Exchange Online remote powershell, 
Reconnect-Exo checks for existing, clears broken, and issues a clean connect-exo. 
Connect-exo does a clean initial connection. 
Disconnect-Exo cleanly unloads any EXO connection. 
#>
#*------v Function Reconnect-EXO v------
Function Reconnect-EXO {
    <# 
    .SYNOPSIS
    Reconnect-EXO - Test and reestablish PSS to https://ps.outlook.com/powershell/
    .NOTES
    Written By: Todd Kadrie
    Website:	https://www.toddomation.com
    Twitter:	https://twitter.com/tostka
    Inspired by function written by: ExactMike Perficient, Global Knowl... (Partner)  
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Change Log
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 2/10/14 posted version 
    .DESCRIPTION
    I use this for routine test/reconnect of EXO. His orig use was within batches, to break up and requeue chunks of commands. Great concept for either use. :D
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
    
    if ($EOLSession.state -eq 'Broken' -or !$EOLSession) {Disconnect-EXO; Start-Sleep -Seconds 3; Connect-EXO} ;     
}#*------^ END Function Reconnect-EXO ^------
if(!(get-alias | ?{$_.name -like "rxo"})) {Set-Alias 'rxo' -Value 'Reconnect-EXO' ; } ;
#*------v Function Connect-EXO v------
Function Connect-EXO {
    <# 
    .SYNOPSIS
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .NOTES
    Written By: : Todd Kadrie
    Website:	https://www.toddomation.com
    Twitter:	https://twitter.com/tostka
    Based on 'overlapping functions' concept by: ExactMike Perficient, Global Knowl... (Partner)  
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Change Log
    # 1:49 PM 11/2/2017 Connect-EXO:coded around non-profile gaps from not having get-admincred() - added the prompt code in to fake it 
    # 12:26 PM 9/11/2017 debugged retry - catch doesn't fire properly on new-Pssession, have to test the $error state, to detect auth fails (assuming the bad pw error# is specific). $error test is currently targeting specific error returned on a bad password.
    # 11:13 AM 9/11/2017 added retry, for when connection won't hold and fails breaks - need to watch out that bad pw doesn't lock out the acct!
    * 7:48 AM 3/15/2017 added pss, doing tweaks to put into prod use
    .DESCRIPTION
    .PARAMETER  ProxyEnabled
    Use Proxy-Aware SessionOption settings [-ProxyEnabled]
    .PARAMETER  CommandPrefix
    [noun]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix exolab ]
    .PARAMETER  Credential
    Credential to use for this connection [-credential 'logon@DOMAIN.com']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    connect-exo
    Connect using defaults, and leverage any pre-set $global:o365cred variable
    .EXAMPLE
    connect-exo -CommandPrefix exo -credential (Get-Credential -credential logon@DOMAIN.com)  ; 
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    *---^ END Comment-based Help  ^--- #>

    Param(
        [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")][boolean]$ProxyEnabled = $False,  
        [Parameter(HelpMessage="[noun]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix exolab]")][string]$CommandPrefix = 'exo',
        [Parameter(HelpMessage="Credential to use for this connection [-credential 'logon@DOMAIN.com']")]$Credential = $global:o365cred
    ) ; 
    if(!$CommandPrefix){ $CommandPrefix='exo' ; } ; 
    
    <# below are profile only, for nonprofile, instead I copy in the following populated constants:$bUseo365COAdminUID, $o365AdmUid, $o365COAdmUid, $o365LabAdmUid, $o365LabCOAdmUid
    #>
    
    $EXOsplat=@{
        ConfigurationName="Microsoft.Exchange" ;
        ConnectionUri="https://ps.outlook.com/powershell/" ;
        Authentication="Basic" ;
        AllowRedirection=$true;
    } ; 

    if(!$global:o365cred){
        write-host -foregroundcolor green "`$GLOBAL:O365CRED NOT FOUND: prompting..." ; 
        #$global:o365cred=get-credential "logon@DOMAIN.com" ; 
        # 1:39 PM 11/2/2017 use Get-AdminCred if it's present, but outside of profile use, it's not (lifted from non-profile verb-ExO.ps1)
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
        }  ; 
    } else {
        $EXOsplat.Add("Credential",$global:o365cred);
    }
    # I want to see where I connected...
    Add-PSTitleBar 'EXO' ;

    If ($ProxyEnabled) {
        $EXOsplat.Add("sessionOption",$(New-PsSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic));
    	Write-Host "Connecting to Exchange Online via Proxy"  ; 
    } Else {
    	Write-Host "Connecting to Exchange Online"  ; 
    } ; 
    # 11:05 AM 9/11/2017 monitor-EXOMigr.ps1 keeps crapping out on connects, put in a retry process...
    # 8:56 AM 6/6/2016 add retry support
    $Exit = 0 ; # zero out $exit each new cmd try/retried
    # do loop until up to 4 retries...
    Do {
        $error.clear() ;# 11:42 AM 9/11/2017 add preclear, we're going to post-test the $error
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
            #>
             # break-exit here, completes the Until block
            $Exit = $Retries ; 
        } Catch {
            # capture auth errors - nope, they never get here, if use throw, it doesn't pass in the auth $error, gens a new one. 
            # pause to give it time to reset
            Start-Sleep -Seconds $RetrySleep ; 
            $Exit ++ ; 
            Write-Verbose "Failed to exec cmd because: $($Error[0])" ; 
            Write-Verbose "Try #: $Exit" ; 
            If ($Exit -eq $Retries) {Write-Warning "Unable to exec cmd!"} ; 
        } # try-E
    } Until ($Exit -eq $Retries) # loop-E    
} ; #*------^ END Function Connect-EXO ^------
if(!(get-alias | ?{$_.name -like "cxo"})) {Set-Alias 'cxo' -Value 'Connect-EXO' ; } ;
#*------v Function Disconnect-EXO v------
Function Disconnect-EXO {
    <# 
    .SYNOPSIS
    Disconnect-EXO - Disconnects any PSS to https://ps.outlook.com/powershell/ (cleans up session after a batch or other temp work is done)
    .NOTES
    Updated By: Todd Kadrie
    Website:	https://www.toddomation.com
    Twitter:	https://twitter.com/tostka
    Inspired by function written by:  ExactMike Perficient, Global Knowl... (Partner)  
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Change Log
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 8:49 AM 3/15/2017 Disconnect-EXO: add Remove-PSTitleBar 'EXO' to clean up on disconnect, doing tweaks to put into prod use
    * 2/10/14 posted version 
    .DESCRIPTION
    I use this to smoothly cleanup connections. 
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
    Get-PSSession | Where-Object {$_.ComputerName -like '*.outlook.com'} | Remove-PSSession ; 
    Remove-PSTitlebar 'EXO' ; 
} ; #*------^ END Function Disconnect-EXO ^------
if(!(get-alias | ?{$_.name -like "dxo"})) {Set-Alias 'dxo' -Value 'Disconnect-EXO' ; } ;
#*======^ END END VERB-EXO SET ^======
#           ===========

#           ===========
##*======v VERB-CCMS SET v======
<# set of functions that connect/reconnect/disconnect with o365 Security & Compliance remote powershell, 
Reconnect-CCMS checks for existing, clears broken, and issues a clean connect-CCMS. 
Connect-CCMS does a clean initial connection. 
Disconnect-CCMS cleanly unloads any EXO connection. 
#>
#*------v Function Reconnect-CCMS v------
Function Reconnect-CCMS {
    <# 
    .SYNOPSIS
    Reconnect-CCMS - Test and reestablish PSS to https://ps.outlook.com/powershell/
    .NOTES
    Written By: Todd Kadrie
    Website:	https://www.toddomation.com
    Twitter:	https://twitter.com/tostka
    Port of my verb-EXO functs for o365 Sec & Compliance Ctr RemPS
    Change Log
    # 1:04 PM 6/20/2018 CCMS variant, works
    .DESCRIPTION
    I use this for routine test/reconnect of CCMS. 
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Reconnect-CCMS; 
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    if ($CCMSSession.state -eq 'Broken' -or !$CCMSSession) {Disconnect-CCMS; Start-Sleep -Seconds 3; Connect-CCMS} ;     
}#*------^ END Function Reconnect-CCMS ^------
if(!(get-alias | ?{$_.name -like "rccms"})) {Set-Alias 'rccms' -Value 'Reconnect-CCMS' ; } ;
#*------v Function Connect-CCMS v------
Function Connect-CCMS {
    <# 
    .SYNOPSIS
    Connect-CCMS - Establish PSS to https://ps.compliance.protection.outlook.com/powershell-liveid/
    .NOTES
    Written By: : Todd Kadrie
    Website:	https://www.toddomation.com
    Twitter:	https://twitter.com/tostka
    Change Log
    # 1:31 PM 7/9/2018 added suffix hint: if($CommandPrefix){ '(Connected to CCMS: Cmds prefixed [verb]-cc[Noun])' ; } ;
    # 12:25 PM 6/20/2018 port from cxo:     Primary diff from EXO connect is the "-ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/" all else is the same, repurpose connect-EXO to this
    .DESCRIPTION
    .PARAMETER  ProxyEnabled
    Use Proxy-Aware SessionOption settings [-ProxyEnabled]
    .PARAMETER  CommandPrefix
    [noun]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix exolab ]
    .PARAMETER  Credential
    Credential to use for this connection [-credential 'logon@DOMAIN.com']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Connect-CCMS
    Connect using defaults, and leverage any pre-set $global:o365cred variable
    .EXAMPLE
    Connect-CCMS -CommandPrefix exo -credential (Get-Credential -credential logon@DOMAIN.com)  ; 
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/connect-to-scc-powershell?view=exchange-ps
    *---^ END Comment-based Help  ^--- #>

   Param(
        [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")][boolean]$ProxyEnabled = $False,  
        [Parameter(HelpMessage="[noun]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix CCMSlab]")][string]$CommandPrefix = 'cc',
        [Parameter(HelpMessage="Credential to use for this connection [-credential 'logon@DOMAIN.com']")]$Credential = $global:o365cred
    ) ; 
    # 12:10 PM 3/15/2017 disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
    if(!$CommandPrefix){ $CommandPrefix='cc' ; } ; 
    
    # 9:17 AM 3/15/2017 Connect-CCMS: add quotes around all string/non-boolean/non-int values in the splat
    $CCMSsplat=@{
        ConfigurationName="Microsoft.Exchange" ;
        ConnectionUri="https://ps.compliance.protection.outlook.com/powershell-liveid/" ;
        Authentication="Basic" ;
        AllowRedirection=$true;
    } ; 

    if(!$global:o365cred){
        write-host -foregroundcolor green "`$GLOBAL:O365CRED NOT FOUND: prompting..." ; 
        #$global:o365cred=get-credential "logon@DOMAIN.com" ; 
        # 1:39 PM 11/2/2017 use Get-AdminCred if it's present, but outside of profile use, it's not (lifted from non-profile verb-ExO.ps1)
        if(test-path function:\get-admincred) { 
            Get-AdminCred ; $CCMSsplat.Add("Credential",$global:o365cred);
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
            $CCMSsplat.Add("Credential",$global:o365cred);
        }  ; 
    } else {
        $CCMSsplat.Add("Credential",$global:o365cred);
    }
    # I want to see where I connected...
    Add-PSTitleBar 'CCMS' ;
    If ($ProxyEnabled) {
        $CCMSsplat.Add("sessionOption",$(New-PsSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic));
        Write-Host "Connecting to Exchange Online via Proxy"  ; 
    } Else {
        Write-Host "Connecting to Exchange Online"  ; 
    } ; 
    $Exit = 0 ; # zero out $exit each new cmd try/retried
    # do loop until up to 4 retries...
    Do {
        $error.clear() ;# 11:42 AM 9/11/2017 add preclear, we're going to post-test the $error
        Try {
            $Global:CCMSSession = New-PSSession @CCMSsplat ; 
            if($error.count -ne 0) {
                if($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed'){
                    write-warning "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    EXIT ; 
                } ; 
            } ; 
            $Global:CCMSModule = Import-Module (Import-PSSession $Global:CCMSSession -Prefix $CommandPrefix -DisableNameChecking -AllowClobber) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ; 
             # break-exit here, completes the Until block
            $Exit = $Retries ; 
        } Catch {
            # capture auth errors - nope, they never get here, if use throw, it doesn't pass in the auth $error, gens a new one. 
            # pause to give it time to reset
            Start-Sleep -Seconds $RetrySleep ; 
            $Exit ++ ; 
            Write-Verbose "Failed to exec cmd because: $($Error[0])" ; 
            Write-Verbose "Try #: $Exit" ; 
            If ($Exit -eq $Retries) {Write-Warning "Unable to exec cmd!"} ; 
        } # try-E
    } Until ($Exit -eq $Retries) # loop-E   
    if($CommandPrefix){ '(Connected to CCMS: Cmds prefixed [verb]-cc[Noun])' ; } ;
} ; #*------^ END Function Connect-CCMS ^------
if(!(get-alias | ?{$_.name -like "cccms"})) {Set-Alias 'cccms' -Value 'Connect-CCMS' ; } ;
#*------v Function Disconnect-CCMS v------
Function Disconnect-CCMS {
    <# 
    .SYNOPSIS
    Disconnect-CCMS - Disconnects any PSS to https://ps.outlook.com/powershell/ (cleans up session after a batch or other temp work is done)
    .NOTES
    Updated By: Todd Kadrie
    Website:	https://www.toddomation.com
    Twitter:	https://twitter.com/tostka
    Change Log
    # 12:42 PM 6/20/2018 ported over from disconnect-exo
    .DESCRIPTION
    I use this to smoothly cleanup connections. 
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Disconnect-CCMS; 
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    *---^ END Comment-based Help  ^--- #>
    # 9:25 AM 3/21/2017 getting undefined on the below, pretest them
    if($Global:CCMSModule){$Global:CCMSModule | Remove-Module -Force ; } ; 
    if($Global:CCMSSession){$Global:CCMSSession | Remove-PSSession ; } ; 
    # "https://ps.compliance.protection.outlook.com/powershell-liveid/" ; should still work below
    Get-PSSession | Where-Object {$_.ComputerName -like '*.outlook.com'} | Remove-PSSession ; 
    Remove-PSTitlebar 'CCMS' ; 
} ; #*------^ END Function Disconnect-CCMS ^------
if(!(get-alias | ?{$_.name -like "dccms"})) {Set-Alias 'dccms' -Value 'Disconnect-CCMS' ; } ;
##*======^ END VERB-CCMS SET ^======
#           ===========

## ======== ^ O365 EXO ETC FUNCS() ^ ===

