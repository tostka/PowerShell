# run-LogParser-RemoteQrySeries.ps1
# run-LogParser-FanQry.ps1

<# 
.SYNOPSIS
run-LogParser-FanQry.ps1 - Run Fan-Remoted LogParser Queries on the specified mail servers
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
9:44 AM 4/14/2016 functioning now, as serial invoke-command series on pre-opened PSS (was failing when trying to do both conn & cmd on same line).
7:47 AM 3/30/2016 - initial version
.DESCRIPTION
run-LogParser-FanQry.ps1 - Run Fan-Remoted LogParser Queries on the specified mail servers
.PARAMETER  Server
Target Server(s) to be utilized for fan-remoting(use -Server or -Site, not both)
.PARAMETER Site
Alternate Target Site to be utilized for fan-remoting (use -Server or -Site, not both)
.PARAMETER tDate
Target Date for parsing logs
.PARAMETER  showDebug
Debugging Flag [$switch]
.PARAMETER whatIf
Whatif Flag  [$switch]
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Outputs csv reports for queries, to the central ArchPath configured (HubServer\d$\scripts\logs)
.EXAMPLE
.LINK
*---^ END Comment-based Help  ^--- #>


Param(
    [Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Target Server(s) to be utilized for fan-remoting(use -Server or -Site, not both)")][ValidateNotNullOrEmpty()]
    [string]$Server
    ,[Parameter(Position=1,Mandatory=$false,HelpMessage="Alternate Target Site to be utilized for fan-remoting (use -Server or -Site, not both)[Site:Site1|Site2|Site3, Site1=default]")][ValidateSet("Site1","Site2","Site3")]
    [string]$Site="Site1"
    ,[Parameter(Mandatory=$False,HelpMessage="Specfify Date to have EWS Traffic Summarized into an Hourly Historgram[[mm/dd/yyyy] ")]
    [string]$LogDate
    ,[Parameter(HelpMessage='Debugging Flag [$switch]')]
    [switch] $showDebug
    ,[Parameter(HelpMessage='Whatif Flag  [$switch]')]
    [switch] $whatIf
) # PARAM BLOCK END

# debug: cls ; .\run-LogParser-FanQry.ps1 -Site Site1 -LogDate 04/8/16 -showdebug ;
#*======v CONSTANTS & ENV CONFIG v======
# pick up the bDebug from the $ShowDebug switch parameter
if ($ShowDebug) {$bDebug=$true};
if ($whatIf) {$bWhatIf=$true};
if ($bDebug) {$DebugPreference = "Continue" ; write-debug "(`$bDebug:$bDebug ;`$DebugPreference:$DebugPreference)" ; };

$ServersHCNA="HubServer0;HubServer1;BCCMS650;BCCMS651" ; $ServersHCNA=$ServersHCNA.split(";") ;
$ServersHCAU="Site3ServerName0;Site3ServerName1" ; $ServersHCAU=$ServersHCAU.split(";") ; 
$ServersHCEU="Site2ServerName0;Site2ServerName1" ; $ServersHCEU=$ServersHCEU.split(";") ; 

$ArchPath="\\HubServer0\D$\scripts\rpts\" ; 

$rgxDateValidMDY="^([0]?[1-9]|[1][0-2])[./-]([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0-9]{4}|[0-9]{2})$" ; 
# regex for matching/confirming server is a CAS server according per firm's s naming scheme 
$rgxTCasServers="[construct a regex matching the firm's naming standard for CAS servers]" ; 


if( $LogDate -notmatch $rgxDateValidMDY ){
        throw "Invalid Date specified ($($LogDate)), aborting" ; 
} else {
    $LogDate=get-date -Date $LogDate ; 
}; 


#$tServers = $ServersHCNA ;
$LPQryDesc="Process specified date log, and produce EWS Histogram per Hour" ; 

# 2:18 PM 3/29/2016 shift the $sQot's into the function
#$LPQry="SELECT QUANTIZE(TO_LOCALTIME(time),3600) as Hour, count(*) AS Hits INTO 'XOUTFILEX' FROM 'XTLOGPATHX' WHERE cs-uri-stem LIKE '%/EWS/%' GROUP BY Hour ORDER BY Hour"; 
# 2:26 PM 3/29/2016 when this comes through, everything after the first ' is truncated, lose them and add them later
$LPQry="SELECT QUANTIZE(TO_LOCALTIME(time),3600) as Hour, count(*) AS Hits INTO XOUTFILEX FROM XTLOGPATHX WHERE cs-uri-stem LIKE XFILTERX GROUP BY Hour ORDER BY Hour"; 
#$LogDate="3/27/16" ; 
#*======^ END CONSTANTS & ENV CONFIG ^======

#*======v GENERAL FUNCTIONS v======
#*----------------v Function Get-ExchangeServerInSite v----------------
Function Get-ExchangeServerInSite {
    <# 
    .SYNOPSIS
    Get-ExchangeServerInSite - Returns the name of an Exchange server in the local AD site. 
    .NOTES
    Written By: Mike Pfeiffer
    Website:	http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/

    Change Log
    1:58 PM 9/3/2015 - added pshelp and some docs
    April 12, 2010 - web version
    .DESCRIPTION
    Get-ExchangeServerInSite - Returns the name of an Exchange server in the local AD site. 
    Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange 2010 servers.
    Returned object includes the post-filterable Role property which reflects the following 
    installed-roles on the discovered server
	    Mailbox Role – 2
        Client Access Role – 4
        Unified Messaging Role – 16
        Hub Transport Role – 32
        Edge Transport Role – 64 
        Add the above up to combine roles:
        HubCAS = 32 + 4 = 36
        HubCASMbx = 32+4+2 = 38
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns the name of an Exchange server in the local AD site. 
    .EXAMPLE
    .\Get-ExchangeServerInSite
    .EXAMPLE
    get-exchangeserverinsite |?{$_.roles -match "(4|32|36)"}
    Return Hub,CAS,or Hub+CAS servers
    .EXAMPLE
    If(!($ExchangeServer)){$ExchangeServer=(Get-ExchangeServerInSite |?{($_.roles -eq 36) -AND ($_.FQDN -match "ServerPrefix.*")} | Get-Random ).FQDN } 
    Return a random HubCas Role server with a name beginning ServerPrefix
    .LINK
    http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
    #>

    $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite] ; 
    $siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName ; 
    $configNC=([ADSI]"LDAP://RootDse").configurationNamingContext ; 
    $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ; 
    $objectClass = "objectClass=msExchExchangeServer" ; 
    $version = "versionNumber>=1937801568" ; 
    $site = "msExchServerSite=$siteDN" ; 
    $search.Filter = "(&($objectClass)($version)($site))" ; 
    $search.PageSize=1000 ; 
    [void] $search.PropertiesToLoad.Add("name") ; 
    [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles") ; 
    [void] $search.PropertiesToLoad.Add("networkaddress") ; 
    $search.FindAll() | %{
        New-Object PSObject -Property @{
            Name = $_.Properties.name[0] ; 
            FQDN = $_.Properties.networkaddress |
                %{if ($_ -match "ncacn_ip_tcp") {$_.split(":")[1]}} ; 
            Roles = $_.Properties.msexchcurrentserverroles[0] ; 
        } ; 
    } ; 
} #*----------------^ END Function Get-ExchangeServerInSite ^---------------- ; 
#$ExHubCas = (Get-ExchangeServerInSite |?{($_.roles -eq 36) -AND ($_.FQDN -match "ServerPrefix.*")} | Get-Random ).FQDN ; 

#*======^ END GENERAL FUNCTIONS ^======


switch ($Site.ToUpper()){
    "Site1" {$tServers=$ServersHCNA}
    "Site2" {$tServers=$ServersHCAU}
    "Site3" {$tServers=$ServersHCEU}
}

write-host -ForegroundColor Yellow "$("="*6)`nWARNING!: THIS SCRIPT WILL FAN-RUN THE FOLLOWING LOGPARSER QRY ON SERVERS:`n$($tServers)`nFOR LOGS DATED:$($LogDate)`nRUNNING QRY:`n$LPQryDesc`n$("="*6)"
$bRet=Read-Host "Enter YYY to continue. Anything else will exit" 
if ($bRet.ToUpper() -eq "YYY") {

    write-host "===PASS STARTING: RUNNING SPECIFIED LOGPARSER QRY ON ALL LOCAL HUBS:`n$LPQryDesc`n" ;

    $error.clear() ;
    TRY {
        # 11:59 AM 4/13/2016 add -username & -message
        if(!($global:SIDcred)){$script:cred=Get-Credential -UserName "DOMAIN\adminid" -Message "No `$global:SIDcred found, enter suitable domain credentials for connection"} 
        else {$script:cred=$global:SIDcred} ; 

        # 9:03 AM 3/29/2016 This is a stock function that will be concat'd into the $SB scriptblock along with it's pre-expanded params hash
        # 2:10 PM 4/13/2016 on ex servers, it's executing in psv2, no [ordered] hashtables etc
        #*------v Function RunRemLPQry v------
        function RunRemLPQry ($LogDate,$LPQry,$ArchPath,$bDebug) {
            if ($bDebug) {$DebugPreference = "Continue" ; write-debug "(`$bDebug:$bDebug ;`$DebugPreference:$DebugPreference)" ; };
            # uses positional params
            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$LogDate: $LogDate"; } ;
            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$LPQry: $LPQry"; } ;
            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$ArchPath: $ArchPath"; } ;
            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$bDebug: $bDebug"; } ;
            #if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):MSG"; } ;
            $sQot = [char]34 ;
            $sQotS = [char]39
            $LPPath="C:\Program Files (x86)\Log Parser 2.2\LogParser.exe" ;
            # 12:29 PM 4/13/2016, invalid datestring correct yMMd => yyMMdd
            $LogFName="u_ex$(get-date $LogDate -Format 'yyMMdd')*.log" ; 
            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$LogFName:$LogFName"; } ;
            # convert to local shortpath
            $fso = New-Object -ComObject Scripting.FileSystemObject
            $LPPath= $fso.GetFile((get-childitem $LPPath).Fullname).Shortpath ;    
            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$LPPath:$LPPath"; } ;
            $fso = $null ; 
            $tLogPath="\\$($env:COMPUTERNAME)\C$\inetpub\logs\LogFiles\W3SVC1\$($LogFName)";
            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$tLogPath:$tLogPath"; } ;
            # 1:47 PM 4/13/2016 create a mult-value object to return out of the func
            $ProcProps = @{
                ErrorState = $($_);
                Completed =$false;
                OutFile=$null;
                OutReport=$null
                StdErr=$null;
                StdOut=$null;
            }  ; # props hash end
            # create an obj from the hash
            $ProcObj = New-Object PSObject -Property $ProcProps ; 
            if((test-path -path $LPPath)) {
                # 12:19 PM 4/13/2016 split the tests, simpler to isolate missing pieces
                if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):LPPath found"; } ;
                if (get-childitem -path $tLogPath) {
                    if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$LPPath:$LPPath`n`$LogFName:$LogFName"; } ;
                    #$oFile="C:\scripts\logs\$($env:COMPUTERNAME)-$($LogFName.replace('*','').replace('.',''))-EWSPerHr.csv" ;  
                    # 2:18 PM 4/13/2016back to UNC (old issue was the sQotS's)
                    $oFile="\\$($env:COMPUTERNAME)\C$\scripts\logs\$($env:COMPUTERNAME)-$($LogFName.replace('*','').replace('.',''))-EWSPerHr.csv" ;  
                    if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$oFile:$oFile"; } ;
                    if(test-path -path $oFile){write-host "Removing existing `$oFile:$oFile" ; Remove-Item $oFile -Force} ; 
                    write-host "`$oFile:$oFile" ; 
                    $SPArgs=$LPQry ;                     
                    $SPArgs=$SPArgs.replace("XTLOGPATHX",$($sQotS + $tLogPath + $sQotS)).replace("XOUTFILEX",$($sQotS + $oFile + $sQotS)).replace("XFILTERX","'%/EWS/%'") ; 
                    # 2:45 PM 3/29/2016 switch to dblquotes
                    $SPArgs=$sQot + $($SPArgs) + $sQot ;
                    $SPArgs+=" -o:csv" ; 
                    if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):`$SPArgs:$SPArgs"; } ;
                    # 11:56 AM 3/29/2016 try to capture the output
                    $soutf= [System.IO.Path]::GetTempFileName() ; 
                    $serrf= [System.IO.Path]::GetTempFileName() ; 
                    if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):tempfiles:`$soutf:$soutf`n`$serrf:$serrf"; } ;
                    write-host "=starting-process:$($LPPath)..."
                    $process=Start-Process -FilePath $($LPPath) -argumentlist $SPArgs -wait -NoNewWindow -PassThru -RedirectStandardOutput $soutf -RedirectStandardError $serrf ; 
                    switch($process.ExitCode){
                        0 {
                            write-host "Logparser exited without errors" ; 
                            if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):ExitCode 0 returned (No Error)"; } ;
                            # 2:50 PM 3/29/2016 add a 5sec wait to clear lock
                            sleep 5 ;
                            # $ProcObj: ErrorState=$($_);Action=$null;ProcdStat=$null;Completed=$false;Outfile=$null;StdErr=$null;StdOut=$null;
                            $ProcObj.ErrorState=$process.ExitCode ;
                            $ProcObj.Completed=$true ;
                            if(test-path $oFile){
                                if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):Populated output file found $(split-path -leaf $oFile)"; } ;
                                #if($($env:COMPUTERNAME) -ne "HubServer0"){ move-item -path $oFile -dest $ArchPath };                         
                                # 1:56 PM 4/13/2016 looks like you can't move files between hosts in a WinRM session, so send it back to call as part of the obj
                                #move-item -path $oFile -dest $ArchPath ;
                                $ProcObj.Outfile = $oFile ;
                                $ProcObj.OutReport = (import-csv $oFile | format-table -auto | out-string) ; 
                                #| out-default| out-string) ; 
                            } else { 
                                write-error "$((get-date).ToString("HH:mm:ss")):Missing Output file! $($oFile)";
                            } ; 
                        } ; 
                        1 {
                            write-host "Logparser exited with errors" ; if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):ExitCode 1 returned"; } ;
                            $ProcObj.Completed=$false ;
                        }
                        default {if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):other ExitCode returned $($process.ExitCode)"; } ;}
                    } ; 
                    # these dump into the pipeline (how the LogParser output gets output back into the calling script)
                    #if((get-childitem $soutf).length){ (get-content $soutf) | out-string } ; remove-item $soutf ; 
                    if((get-childitem $soutf).length){ $ProcObj.StdOut=(get-content $soutf) | out-string } ; remove-item $soutf ; 
                    #if((get-childitem $serrf).length){(get-content $serrf) | out-string} ; remove-item $serrf  ; 
                    if((get-childitem $serrf).length){$ProcObj.StdErr=(get-content $serrf) | out-string} ; remove-item $serrf  ; 
                } else { 
                    if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):LPPath NOT found! ABORTING"; } ;
                }  ;
            
            } else {
                if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):LPPath NOT found! ABORTING"; } ;
            }  ; 

            # dump $ProcObj into the pipeline, returns at call point
            Write-Output $ProcObj ; 

            if ($bDebug -OR ($DebugPreference = "Continue")) {
                Write-Debug -Verbose:$true "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
                $bDebug=$false ; $DebugPreference = "SilentlyContinue" ;
            } ; 
        } #*------^ END Function RunRemLPQry ^------ ;   

        #Nested quotes required due if any embedded spaces in value.
        # 12:08 PM 4/13/2016 need to ren tLogDate => LogDate to match the param in use
        $QryParams = @{
            LogDate = "$($LogDate)" ; 
            LPQry = "'$($LPQry)'"  ; 
            ArchPath = "'$($ArchPath)'"  ; 
            bDebug = "$(if($bDebug){"`$true"} else {"`$false"})"
        } ; 
        # archpath evals as: 'SELECT QUANTIZE(TO_LOCALTIME(time),3600) as Hour, count(*) AS Hits INTO  FROM '' WHERE cs-uri-stem LIKE '%/EWS/%' GROUP BY Hour ORDER BY Hour'   
        # create a scriptblock variable for the combo above (concat'ing funct + params hash)
        $SB = [scriptblock]::Create(".{${Function:RunRemLPQry}} $(&{$args}@QryParams)") ; 
    
        #$tServers | foreach { 
        #11:54 AM 4/13/2016 make it a stock foreach
        foreach ($srv in $tServers) {
            write-host ("`n=== v " + $($srv)) ;
            #invoke-command -cn $_ -cred $script:cred -scriptblock $SB ; 
            # 1:27 PM 3/29/2016 capture the lastexitcode
            #$oRet=invoke-command -cn $srv -cred $script:cred -scriptblock $SB ; 
            # 11:53 AM 4/13/2016 above failing with he logon failure
            # tests fine from EMS console, build the PSS separately
            $sess = New-PSSession -ComputerName $($srv) -Credential $script:cred ;
            # demo cmd to test pss: Invoke-Command -Session $sess -ScriptBlock{Get-Service} ;
            if($sess){ 
                $oRet=invoke-command -Session $sess -scriptblock $SB  ; 
                if($bDebug){
                    Write-Host -ForegroundColor Yellow "`$oRet contents:..."
                    #1:32 PM 2/25/2015 this is an object, returned from a command, can't loop it
                    # dump out the properties and values out of the ret'd obj
                    #$oRet | Select-Object -Property * 
                    write-host ($oRet | format-list | out-string )
                } # if-E
                <# $oRet componetns that come back
                    ErrorState = $($_);
                    Completed =$false;
                    Outfile=$null;
                    OutReport=$null;
                    StdErr=$null;
                    StdOut=$null;
                #>
                if(test-path $oRet.Outfile){
                    if($bDebug){Write-Debug "$((get-date).ToString("HH:mm:ss")):Populated output file found $(split-path -leaf $oRet.Outfile), moving to $($ArchPath) "; } ;
                    if( $cfile=(get-childitem -path (join-path -path $archpath -ChildPath $(split-path -leaf $oRet.Outfile)) -ErrorAction SilentlyContinue ) ){ 
                        $newname="$($cfile.BaseName)-$(get-date -format "yyyyMMdd-HHmmtt")$($cfile.Extension)" ; 
                        write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):renaming existing file $($cfile.Name) to $($newname)" ; 
                        Rename-Item -path $cfile -NewName $newname ; 
                    } ; 
                    move-item -path $($oRet.Outfile) -dest $($ArchPath) ; 
                }  # if-E outfile; 
                get-pssession -id $sess.id | Remove-PSSession ; 
            } else { 
                write-error "$((get-date).ToString("HH:mm:ss")):Failed to establish PSSession with server:$($srv)";
            } # if-E $sess
            write-host ("`n=== ^ " + $srv) ; 
        }  # loop-E $tservers ; 

        write-host "===PASS COMPLETED: ALL OUTPUT REPORTS CAN BE FOUND IN $($QryParams.ArchPath)" ;
    
    } CATCH {
        $msg=": Error Details: $($_)";
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
        # 1:00 PM 1/23/2015 autorecover from fail, STOP (debug), EXIT (close), or use Continue to move on in loop cycle
        Continue ; 
        #Exit ;
    };   # try/catch-E
} else {
     Write-Host "Invalid response. Exiting"
     # exit <asserted exit error #>
     #exit 1
} # if-block end
if ($bDebug -OR ($DebugPreference = "Continue")) {
    Write-Debug -Verbose:$true "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
    $bDebug=$false ; $DebugPreference = "SilentlyContinue" ;
} ; 

