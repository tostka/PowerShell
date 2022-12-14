# append-x500LEDN-ToMbxEmailAddressess.ps1

<#
.SYNOPSIS
append-x500LEDN-ToMbxEmailAddressess.ps1 - Script to splice in historical LEDN onto migrated mailboxes, to support reply-to etc on PST migrated email. 
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
* 12:53 PM 1/7/2018 - updated, cleaned up, added Pshelp
* 9/25/2014 - initial version

PROCESS:
1. Export LEDN & WindowsEmailAddress from SOURCE Exchange mail system  1-liner:
#-=-=-=-=-=-=-=-=
get-mailbox | select alias,displayname,windowsemailaddress,legacyexchangedn | export-csv -path ".\[firm]-LEDNS-$(get-date -format 'yyyyMMdd-HHmmtt').csv" -notype ; 
#-=-=-=-=-=-=-=-=

2. Run match pass, to verify all are present in migrated system: 
Cleanup/update/remove perm-missing items till you get a clean pass with csv, 
The final script is  going to match on _migrated windowsemailaddress_ from SOURCE sys 
(because HR db frequently mangles displayname in migrations)
So ensure that all of the migrated email addresses are in place - watch out of maiden-name changes etc,
For users that had maiden name addresses in source sys, but migrated with married name. 
#-=matching test script  1-liner-=-=-=-=-=-=-=
$procMbxs=import-csv .\[firm]-LEDNS-.csv ;foreach($pMbx in $procMbxs) {  if($tmbx = get-mailbox -identity "$($pMbx.WindowsEmailAddress)" -ea 0 ){    write-host -foregroundcolor green  "$((get-date).ToString('HH:mm:ss')):MATCHED:$($pMbx.WindowsEmailAddress)($($pMbx.displayname))=>TO:$($tmbx.WindowsEmailAddress)($($tmbx.displayname))" ;  } else {    write-host -foregroundcolor red  "$((get-date).ToString('HH:mm:ss')):UNABLE TO MATCH:$($pMbx.WindowsEmailAddress)($($pMbx.displayname))" ;  }  ;} ; 
#-=-=-=-=-=-=-=-=

3. Prior to any updates,  1-liner to dump a backup copy of relevant material on the target mailboxes:
#-=-=-=-=-=-=-=-=
get-mailbox -OrganizationalUnit "OU=[dn path to target OU]" | select samaccountname,alias,displayname,windowsemailaddress,EmailAddresses | export-csv -notype -path ".\[firm]-migrated-[OU]-OU-mbxs-pre
-x500LEDN-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
#-=-=-=-=-=-=-=-=

#-=-=Resulting Input CSV format & fields:-=-=-=-=-=-=
Alias,DisplayName,WindowsEmailAddress,LegacyExchangeDN
#-=-=-=-=-=-=-=-=
Obviously the values above are the source mail system values, not migr dest system

Script below:
a. Performs a search of curr mail system against source system windowsemailaddress
b. Builds an x500: address from the src LEDN
c. Confirms the constructed x500Ledn doesn't pre-exist in the migr'd mbx emailaddresses list
d. Then '@{Add=xxx}' appends the new x500Ledn to the existing emailaddresses attrib on the migrated mailbox. 

Post-confirmation X500 dump 1-liner: 
#-=-=-=-=-=-=-=-=
import-csv .\pDE-LEDNS-.csv  | %{    $rcp=get-mailbox $_.windowsemailaddress ;    "==$($rcp.windowsemailaddress):updated X500s" ;    $rcp | Select -Expand EmailAddresses | Where {$_ -like "x500:*"} ;} ;
#-=-=-=-=-=-=-=-=

.DESCRIPTION
append-x500LEDN-ToMbxEmailAddressess.ps1 - Script to splice in historical LEDN onto migrated mailboxes, to support reply-to etc on PST migrated email. 
.PARAMETER  showDebug
Debugging Flag [-showDebug]
.PARAMETER  whatIf
Whatif Flag DEFAULTS TRUE! [-whatIf]
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.\append-x500LEDN-ToMbxEmailAddressess.ps1 -inputCSV .\XXX-LEDNS-.csv -whatif:$true ; 
Run a whatif pass - NOTE whatif defaults $true, requres explict -whatif:$false to run prod changes
.EXAMPLE
.\append-x500LEDN-ToMbxEmailAddressess.ps1 -inputCSV .\XXX-LEDNS-.csv -whatif:$false ; 
Prod pass, forcing the defaulted -whatif to false
.LINK
#>

Param(
    [Parameter(Mandatory=$True,HelpMessage="CSV file specifying Source mail system values for updates [-InputCSV l:\pathto\file.csv)]")]
    [string]$InputCSV,
    [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
    [switch] $showDebug,
    [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
    [switch] $whatIf
)  ; 

if(test-path $InputCSV) {
    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Loading Input file: $($InputCSV)" ;
    $procMbxs=import-csv $InputCSV ; 
    foreach($pMbx in $procMbxs) {
        if($tmbx = get-mailbox -identity "$($pMbx.WindowsEmailAddress)" -ea 0 ){
            write-host -foregroundcolor green  "$((get-date).ToString('HH:mm:ss')):MATCHED:$($pMbx.WindowsEmailAddress)($($pMbx.displayname))=>TO:$($tmbx.WindowsEmailAddress)($($tmbx.displayname))" ; 
            $x500LEDN = "X500:" + $pMbx.LegacyExchangeDN ; 
            if ($tmbx.EmailAddresses -notcontains $x500LEDN) {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):WHATIF:$($whatif):Adding LEDN as X500: `n$($x500LEDN) `nto existing emailaddresses:`n$($tmbx.emailaddresses)" ; 
                set-mailbox -identity $pMbx.WindowsEmailAddress -EmailAddresses @{Add=$x500LEDN} -whatif:$($whatif); 
            } else {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):SKIPPING:Mailbox $($tmbx.windowsemailaddress) already contains $($x500ledn)" ; 
            } ; 
        } else {
            write-host -foregroundcolor red  "$((get-date).ToString('HH:mm:ss')):UNABLE TO MATCH:$($pMbx.WindowsEmailAddress)($($pMbx.displayname))" ; 
        }  ; 
        "===========`n" ; 
    } ; 

} else {
    write-host -foregroundcolor red  "$((get-date).ToString('HH:mm:ss')):MISSING InputCsv:$($InputCSV)" ; 
} ; 

