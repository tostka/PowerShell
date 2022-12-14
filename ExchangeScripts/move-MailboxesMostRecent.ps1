# move-MailboxesMostRecent.ps1

<#
.SYNOPSIS
move-MailboxesMostRecent.ps1 - Simple, 'Toss the most recent NN mbxs into the air & let Automatic Mailbox Provisioning i redistribute them onto new dbs' script
.NOTES
Written By: Todd Kadrie
Website:	http://www.toddomation.com
Twitter:	http://twitter.com/tostka
Additional Credits: REFERENCE
Change Log
# 8:54 AM 3/30/2018 minor doc update
# 3:12 PM 11/1/2017 initial version
.DESCRIPTION
move-MailboxesMostRecent.ps1 - Simple, 'Toss the most recent NN mbxs into the air and and lonto new dbs' script
As part of using this script, you will want to pre-exclude you're 'congested' dbs from Automatic Mailbox Provisioning (AMP) by using this command mark them excluded from AMP: 
Set-MailboxDatabase -Identity dbID -IsExcludedFromProvisioning $true ;  

AMP is covered here https://technet.microsoft.com/en-us/library/ff477621%28v=exchg.150%29.aspx
The short notes on the process Ex AMP uses to determine a dest DB is:
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  1.    Exchange retrieves a list of all mailbox databases in the Exchange 201X 
  organization

  2.    Any mailbox database that's marked for exclusion from the distribution 
  process is removed from the available list of databases. You can control which 
  databases are excluded...

  3.    Any mailbox database that's outside of the database management scopes 
  applied to the administrator performing the operation is removed from the list 
  of available databases. For more information, see Database Scopes later in this 
  topic

  4.    Any mailbox database that's outside of the local Active Directory site 
  where the operation is being performed is removed from the list of available 
  databases

  5.    From the remaining list of mailbox databases, Exchange chooses a database randomly.
  If the database is online and healthy, the database is used by Exchange.
  If it's offline or not healthy, another database is chosen at random. 
  If no online or healthy databases are found, the operation fails with an error
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
.PARAMETER  Database
Source Mailbox Database to have mbxs moved
.PARAMETER  Number
Number of most-recently-created mailboxes to move [-Number nnnnn]
.PARAMETER ShowDebug
Parameter to display Debugging messages [-ShowDebug switch]
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Outputs move information to console. 
.EXAMPLE
.\move-MailboxesMostRecent.ps1 -Database db1 -Number 100 -whatif -showDebug ; 
Move the 100 most-recently created mailboxes on db1 database to other dbs (dest determined by AMP), whatif pass, show debugging messages.
.EXAMPLE
.\move-MailboxesMostRecent.ps1 -Database db1 -Number 100 -whatif -showDebug ; 
Move the 100 most-recently created mailboxes on db1 database to other dbs (dest determined by AMP), live pass.
.LINK
#>

Param(
    [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Mailbox Database to have mbxs moved")]
    [ValidateNotNullOrEmpty()][string]$Database,
    [Parameter(Mandatory=$True,HelpMessage="Number of most-recently-created mailboxes to move [-Number nnnnn]")]
    [int]$Number,
    [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
    [switch] $showDebug,
    [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
    [switch] $whatIf
) # PARAM BLOCK END


#*======v SUB MAIN v======

$domaincontroller="DOMAINCONTROLLER" ;
write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Using BatchName:$($BatchName)" ;
write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Moving most recent $($Number) mailboxes" ;

if($db=get-mailboxdatabase $database -domaincontroller $domaincontroller ){
    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Moving most recent $($Number) mailboxes" ;
    if($mbxs = (get-mailbox -Database $db.identity -DomainController $domaincontroller)) {
        if($showdebug){write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):($($mbxs.count) raw mailboxes returned in db)" };
        $mbxs = ($mbxs |sort whencreated)[(-1*$Number)..(-1)] ; 
        if($showdebug){write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):($($mbxs.count) net mailboxes subject to move)" };
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Filtered most recent mailboxes $($mbxs.count) returned" ;
        $BatchName="ExMoves-db-$($db.name)-REBAL-$(get-date -format 'yyyyMMdd-HHmmtt')";
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):BATCHNAME:$($BatchName)" ;
        foreach($mbx in $mbxs){
            "==$($mbx):" ;
            write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):WHATIF:$($whatif):Moving Mbx:$($mbx.alias):$($mbx.displayname):$($mbx.whencreated)" ;
            $spltMove=@{
                identity=$mbx.alias ;
                BadItemLimit=100 ;
                AcceptLargeDataLoss=$true ;
                suspend=$true ;
                batchname=$BatchName ;
                domaincontroller=$domaincontroller ;
                whatif=$($whatif) ;
            } ; 
            if($showDebug){Write-Verbose -Verbose:$true "$((get-date).ToString("HH:mm:ss")):Executing New-MoveRequest with settings:"; $spltMove|out-string } ;
            New-MoveRequest @spltMove ;
        } ; # loop-E
        if(!$whatif){
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):MOVEREQUESTS STATUS, BATCH:$(BatchName):" ;
            Get-MoveRequest -BatchName $BatchName -domaincontroller $dc | ft -auto DisplayName,Status,TargetDatabase;
        } ; 
    } else {
        throw "NO MATCHING MAILBOXES FOUND IN DB:$($DATABASE)" ; 
    } ; 
} else { 
    throw "specified db $($database) not found" ;  
} ; 
#*======^ END SUB MAIN ^======


