# check-LyncSvcStatus.ps1

<# 
.SYNOPSIS
check-LyncSvcStatus.ps1 - Dynamically check status on all servers in local pool
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Change Log
# 11:06 AM 6/8/2015 - fixed Edge resolution & domain info (dom comes back as the box name on non-dom-joined edge boxes)
10:48 AM 6/8/2015 - updated to handle FE & Edge checks

.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
check-LyncSvcStatus.ps1
.LINK
*----------^ END Comment-based Help  ^---------- #>
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

# load-lms
$sName="Lync"; if (!(Get-Module | where {$_.Name -eq $sName})) {Import-Module $sName -ea Stop};


#*================v LYNC PER DOMAIN SPECS  v================
# 1:16 PM 1/16/2015 corrected typo in nontor settings for lab
# 2:26 PM 4/6/2015 add ordered PSv3 spec to hashes

switch -regex ($env:USERDOMAIN){
	"^DOMAIN-LAB$" {
		# add pretest for non-Lync lab
		$ucSettings=[ordered]@{
      registrarpool="POOL.DOMAINlab.com";
      Site="Site:Lync_SITE_Lab"
      sipaddresstype="EmailAddress";
      sipdomain="DOMAINlab.com"
		} ;
		$ucSettingsNonTOR=[ordered]@{
      registrarpool="POOL.DOMAINlab.com";
      sipdomain="DOMAINlab.com"
		};
		# 1:36 PM 1/12/2015 lab has no china ad child dom
		$domChina = ".global.ad.DOMAINlab.com";
		$dcChina = "DOMAINCONTROLLER2T.global.ad.DOMAINlab.com";
		$domUS = "global.ad.DOMAINlab.com";
		$dcUS="DOMAINCONTROLLER1T.global.ad.DOMAINlasb.com";
		$rootSipDom="DOMAINlab.com"
	}; # switch-E DOMAIN-lab

	"^DOMAIN$" {
		$ucSettings=[ordered]@{
		registrarpool="POOL.DOMAIN.com";
    Site="Site:LyncSITE";
		sipaddresstype="EmailAddress";
		sipdomain="DOMAIN.com"
		} ;
		
		$ucSettingsNonTOR=[ordered]@{
		registrarpool="POOL.DOMAIN.com";
		sipdomain="DOMAIN.com"
		};
		
		$domChina = "china.DOMAIN.com";
		$dcChina = "DOMAINCONTROLLER1.china.DOMAIN.com";
		$domUS = "US.DOMAIN.com";
		$dcUS="DOMAINCONTROLLER0.us.DOMAIN.com";
		
		$rootSipDom="DOMAINlab.com"
				
	}  # switch-E DOMAIN
	
	"^(\w{3}MS520\d(D)*)$" {
    
	} ; 
	
  default {
		Write-Warning "( $env:USERDOMAIN ) is not a recognized domain. Exiting!...";
		exit; 
  };   # switch-default
}; # switch-E

write-host -foregroundcolor yellow "`$ucSettings:";
foreach($row in $ucSettings) {foreach($key in $row.keys) {write-host "$($key): $($row[$key])"}} ;

$UpBoxes=[ordered]@{};
$DnBoxes=[ordered]@{};

#*================^ LYNC PER DOMAIN SPECS ^================

"`$ucSettings.registrarpool:$($ucSettings.registrarpool)"

if($($env:computername).substring(5,3) -eq "620"){
  # FE pool server
  write-verbose -verbose "$($env:computername) is an FE role server...";
  <# 9:44 AM 6/8/2015 older code that requires spec for local pool
  $PoolFqdn = "POOL.DOMAIN.com";
  $L13FE = @((Get-CsPool $PoolFqdn).Computers | Sort-Object); 
  #>
  #
  #$L13FEPool=(get-cspool| ?{($_.Services -like '*Registrar*') -AND ($_.Site -like "*$($ucSettings.registrarpool)*")})
  $L13FEPool=(get-cspool $($ucSettings.registrarpool))
  if(!($L13FEPool)){"NOMATCH!";EXIT}
  #$L13FEPool | select * ; 
  write-verbose -verbose "Pool:$($L13FEPool.fqdn):Checking whole pool...";
  #$L13FE = (get-cspool| ?{($_.Services -like '*Registrar*') -AND ($_.Site -like '*LyncSITE*')} | select computers);
  $ChkPool = ($L13FEPool | select computers).computers ;
  #"$ChkPool" ;
  #

} elseif($($env:computername).substring(5,3) -eq "123"){
  # edge pool server
  write-verbose -verbose "$($env:computername) is an Edge role server...";
  #$EdgSrvrs= (get-cspool| ?{$_.Services -like '*EdgeServer*' -AND ($_.Site -like '*LyncSITE*') } | select computers).computers ; 
  $ChkPool="localhost";
  #$L13Edg=$L13Edg.split(";") ;
} else {
  # non FE/Edge box
  write-verbose -verbose "$($env:computername) is an unrecognized role server. Aborting";
  Exit;
}  # if-E

  
# fully dynamic qry, filters on local site name
write-host ("=" * 15);
#$ChkPool;
foreach ($Srv in $ChkPool){
    write-host -fore yell ("`n" + (get-date).ToString("HH:mm:ss") + ": " + $Srv + " Svc Status:");
    Get-CsWindowsService -computername $Srv | ft name,status -auto ;
    $DeadSvc = (Get-CsWindowsService -computername $Srv | ?{$_.Status -ne 'Running'}  );
    if ($DeadSvc){    write-host -foregroundcolor red "STOPPED SVCS";
      $DeadSvc | ft Name,Status;
    } else {
      write-host "[no non-running svcs]"; 
    } ;
  write-host ("-" * 3);
};
write-host ("=" * 15);

