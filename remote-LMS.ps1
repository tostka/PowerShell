#c:\usr\work\ps\scripts\remote-EMS.ps1|remote-LMS.ps1|remote-PS.ps1|Remote-AD.ps1
#open-RemotePS.ps1
#open-EMSRemote-box.ps1

#*----------V Comment-based Help (leave blank line below) V---------- 

<# 
    .SYNOPSIS
open-PSRemote.ps1 - Remote an Exch EMS connection

    .NOTES
For LMS or Ps remote use, reset the default $sApp to Lync|AD|PS|ExLync|Exch
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Additional Credits: [REFERENCE]
Website:	[URL]
Twitter:	[URL]

Change Log
9:22 AM 11/26/2014 added color & title changes to id
8:47 AM 11/26/2014 chgd name to remote-EMS.ps1
9:21 PM 11/14/2014

    .DESCRIPTION
Remote PS into Exchk, Lync, or AD

    .PARAMETER  $sApp
Target App: PS|Lync|Exch|ExLync

    .PARAMETER  $sTargSrv
Override default host, specify Target FQDN


    .INPUTS
None. Does not accepted piped input.

    .OUTPUTS
None. Returns no objects or output.
System.Boolean
                True if the current Powershell is elevated, false if not.
[use a | get-member on the script to see exactly what .NET obj TypeName is being returning for the info above]

    .EXAMPLE
.\open-EMSRemote-box.ps1

    .EXAMPLE
.\open-emsremote.ps1 -a lync
Open a connection to a random Lync pool member

    .LINK
< name of a related topic, one .LINK keyword per related topic. Can also include a URI to online help (start with http:// or https://)>
*----------^ END Comment-based Help  ^---------- #>

<#
param(
[alias("a")]
[string] $sApp ="Exch",
[alias("f")]
[string] $TargSrv = $( if($sApp -eq "Exch") { "name.domain.com" } elseif($sApp -eq "Lync") { "pool.domain.com" } elseif($sApp -eq "AD") { "server.domain.com" } )
#>

param(
[alias("a")]
[string] $sApp ="Lync",
[alias("t")]
[string] $sTargSrv
)

[CmdletBinding()]

#$ExFqdn="name.domain.com";
$ExFqdn="lynms650.global.ad.domain.com";
$LyncFqdn="pool.domain.com"
$AdDc="server.domain.com"

if ($sApp.ToLower() -notmatch "^(lync|exch|ps|exlync|ad)$") {
   throw {"`$sApp:$sApp is not a suitable specification for the -a parameter [LYNC|EXCH|PS|AD|EXLYNC]. EXITING..."}
}

Write-Verbose "Prompting for credentials..." -verbose;
$cred = Get-Credential;
Write-Verbose "Connecting to $sTargSrv ($sApp)..." -verbose;

switch ($sApp){
  "Lync" {
    if ($sTargSrv) {
      ($LyncURI="https://" + $sTargSrv + "/OcsPowershell")
    } else {
      ($LyncURI="https://" + $LyncFqdn + "/OcsPowershell")
    };
    write-output "LYNC: Connecting to $LyncURI";
    <# 
    #Change the server or pool name in to a properly constructed URL
    $poolname = "https://" + $poolname + "/OcsPowershell"
    $cred = Get-Credential;
    $session = New-PSSession -ConnectionURI $poolname -Credential $cred;
    Import-PsSession $session
    
    #---http://tsoorad.blogspot.com/2013/10/lync-2013-remote-admin-with-powershell.html
    # set session options to bypass the PKI checks - I trust the far side
    $sessionoption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $session = New-PSSession -ConnectionUri https://somelyncfrontendserverFQDN.domain.com/ocspowershell -Credential  $credential -SessionOption $sessionOption
    #>
    
    #$CsSess = New-PSSession –ConnectionUri http://$LyncFqdn/OcsPowershell -Credential $cred;
    #$CsSess = New-PSSession –ConnectionUri $LyncURI -Credential $cred;
    # the below causes the Front End server IIS internal services certificate to be basically ignored, so if you connect to a domain server from a non-domain workstation
    $sessionoption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck;
    $CsSess = New-PSSession -ConnectionUri $LyncURI -Credential $cred -SessionOption $sessionOption
    Import-PSSession -session $CsSess;
    # set colors to white on light gray
    #$Host.UI.RawUI.BackgroundColor=8 ;$Host.UI.RawUI.ForegroundColor= 6;
    # try 56
    $Host.UI.RawUI.BackgroundColor=1 ;$Host.UI.RawUI.ForegroundColor= 7;
    clear-host;
    $host.ui.RawUI.WindowTitle = "REMOTE-LMS 2013"
  }
  
  "Exch" {
    <#
    if ($sTargSrv) {
      ($ExURI="https://" + $sTargSrv + "/powershell")
    } else {
      ($ExURI="https://" + $ExFqdn + "/powershell")
    };
    write-output "EXCH: Connecting to $ExFqdn";
    $ExSess = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri $ExURI -Credential $cred;
    Import-PSSession -session $ExSess;
    #>
    # below works, above doesn't
    if ($sTargSrv) {
      ($ExURI="http://$sTargSrv/powershell")
    } else {
      ($ExURI="http://$ExFqdn/powershell")
    };
    $ExSess = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri $ExURI -Credential $cred;
    Import-PSSession -session $ExSess;
    # set colors 
    $Host.UI.RawUI.BackgroundColor  = "Black";$Host.UI.RawUI.ForegroundColor="Gray"
    clear-host;
    $host.ui.RawUI.WindowTitle = "REMOTE-EMS 2010"
  }
  
  "AD" {
    if ($sTargSrv) {$AdDc=$sTargSrv};
    write-output "AD: Connecting to $AdDc";
    $AdSess = new-pssession -computer $AdDc;
    Invoke-Command -session $AdSess -script { Import-Module ActiveDirectory };
    Import-PSSession -session $AdSess -module ActiveDirectory -prefix Rem;
    # set colors to silver on light blue
    $Host.UI.RawUI.BackgroundColor=9 ;$Host.UI.RawUI.ForegroundColor= 7;
    clear-host;
    $host.ui.RawUI.WindowTitle = "REMOTE-EMS 2010"
  }
  
  "ExLync" {
    write-output "Ex/Lync Hybrid";
    write-output "LYNC: Connecting to $LyncFqdn";
    $CsSess = New-PSSession –ConnectionUri http://$LyncFqdn/OcsPowershell -Credential $cred;
    Import-PSSession -session $CsSess;
    write-output "EXCH: Connecting to $ExFqdn";
    $ExSess = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri http://$ExFqdn/powershell -Credential $cred;
    Import-PSSession -session $ExSess;

  }
  
  "PS" {
    if ($sTargSrv) {
      write-output "Powershell Remote";
      write-output "PS: Connecting to $sTargSrv";
      $PSSess = enter-pssession $sTargSrv -Credential $cred;
      #Import-PSSession -session $CsSess;
    } else {
      throw {"A suitable $sTargSrv must be specified with the -t parameter. EXITING..."}
    };
  }
  
} # swtch-E
  

