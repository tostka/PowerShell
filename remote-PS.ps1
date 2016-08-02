#c:\usr\work\ps\scripts\remote-PS.ps1
#open-RemotePS.ps1
#open-EMSRemote-box.ps1

#*----------V Comment-based Help (leave blank line below) V---------- 

<# 
    .SYNOPSIS
open-PSRemote.ps1 - Remote an Exch EMS connection

    .NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Additional Credits: [REFERENCE]
Website:	[URL]
Twitter:	[URL]

Change Log
8:48 AM 11/26/2014 chgd name to remote-PS.ps1
9:21 PM 11/14/2014

    .DESCRIPTION
Remote PS into Exchk, Lync, or AD

    .PARAMETER  $sApp
Target App: PS|Lync|Exch|ExLync

    .PARAMETER  $sTargSrv
Override default host, specify Target FQDN

    .NOTES
vers: 2:40 PM 11/17/2014

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
[string] $TargSrv = $( if($sApp -eq "Exch") { "mymailna.domain.com" } elseif($sApp -eq "Lync") { "pool.domain.com" } elseif($sApp -eq "AD") { "server.domain.com" } )
#>

param(
[alias("a")]
[string] $sApp ="Exch",
[alias("t")]
[string] $sTargSrv
)

[CmdletBinding()]

#$ExFqdn="mymailna.domain.com";
$ExFqdn="server.domain.com";
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
    
    # the below causes the Front End server IIS internal services certificate to be basically ignored, so if you connect to a domain server from a non-domain workstation
    $sessionoption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck;
    $CsSess = New-PSSession -ConnectionUri $LyncURI -Credential $cred -SessionOption $sessionOption
    Import-PSSession -session $CsSess;
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
  }
  
  "AD" {
    if ($sTargSrv) {$AdDc=$sTargSrv};
    write-output "AD: Connecting to $AdDc";
    $AdSess = new-pssession -computer $AdDc;
    Invoke-Command -session $AdSess -script { Import-Module ActiveDirectory };
    Import-PSSession -session $AdSess -module ActiveDirectory -prefix Rem;
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
  
