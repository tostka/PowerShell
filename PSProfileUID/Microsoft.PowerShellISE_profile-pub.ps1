#*================v NON-ADMIN ISE PROFILE: C:\Users\MyAccount\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1 v================
# Powershell ISE profile
# C:\C:\Users\MyAccount\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1
# notepad2 $profile # WITHIN ISE PS PROMPT
# 12:16 PM 11/5/2015 replaced inclpath detect with profile.ps1 version
# 10:35 AM 6/18/2015 it's trying to open:
# c:\users\MyAccount\documents\windowspowershell\microsoft.powershellise_profile.ps1, empty
# looks like EMS needs it's own
#*================================
<# creation cmds
# create the Microsoft.PowerShellISE_profile.ps1 file: FROM WITHIN THE ISE PS:
New-Item -path $PROFILE ; notepad2 $PROFILE ;
#>

# standard all machines should be ucase and all accts lower
<# 8:23 AM 12/30/2014 allow these to be defined in the MyAccount-prof.ps1 instead
$MyBox="MyComputer","MyComputer"
$AcctWAdmn="MYDOMAIN\user";
$AcctWUser="MYDOMAIN\user";
$AcctHAdmin = "MyAccount";
#>
#*======v CONTSTANTS v======
#$bDebug = $TRUE;
$bDebug = $FALSE;
# standard all machines should be ucase and all accts lower
# 8:25 AM 12/30/2014 don't define these in any other include files! (too confusing)
# define $binpath
if(Test-Path "c:\usr\local\bin"){
  $binpath="c:\usr\local\bin\"
} elseif(Test-Path "d:\scripts"){
  $binpath="d:\usr\local\bin\"
} elseif(Test-Path "c:\scripts"){
  $binpath="c:\scripts\"
} else {
  Write-Error -verbose ((Get-Date).ToString("HH:mm:ss") + ":NO LOCAL BINPATH FOUND. EXITING!:")
  Exit
} # if-E

# Computername contstants
$MyBox="MyComputer","MyComputer"
# 10:23 AM 11/5/2015 added explicit target varis for the above
$MyBoxW="MyComputer" ;
$MyBoxH="MyComputer"
$rgxProdL13Servers="^Server3]$" ; 
$rgxProdServers="^(SITE|SITE|SITE)XXX[0-9][0-9]$" ; 
$rgxLabL13Servers="^Server2]D$" ; 
$rgxLabServers="^(SITE|SITE|SITE)XXX[0-9][0-9]$" ; 
$rgxProdL13EdgSrvrs="^Server1$" ;
$rgxLabL13EdgSrvrs="^Server1$" ;

# User ID constants
$AcctWAdmn="MYDOMAIN\user";
# covers both prod & lab edge server names
$rgxAcctWAdmnEdg="^Server1\\MyAccount$"
#"Server1\MyAccount" ;
$AcctWUser="MYDOMAIN\user";
$AcctHAdmin = "MyAccount";

# Domain contstants
$DomainWork = "DOMAIN";
$DomHome = "DOMAIN";
$DomLab="DOMAIN";
# edge boxes show dom as themselves: ($env:USERDOMAIN) =Server0
# if the ($env:USERDOMAIN) -eq the ($env:COMPUTERNAME), you're on a non-domain-joined box
$DomL13EdgProd=$rgxProdL13EdgSrvrs ;
$DomL13EdgLab=$rgxLabL13EdgSrvrs ;

$LocalInclDir="c:\usr\home\db" ;
$LocalInclSIDDir = "c:\usr\work\ps\scripts";
# distrib shares:
$InclShareCent = "\\Server3.domain.com\USR$\MyAccount\dev\ps\scripts";
$InclShareL13 = "\\Server0.domain.com\SITEc_FS\scripts";
$inclShareLab = "\\Server0\e$\scripts";
# 12:53 PM 1/12/2015 lab SITEc support
$InclSIDDirL = "\\SITEcfs.domain.com\SITEcFileShare\scripts";

$smtpserver = 'Server0' # SMTP server you want to use to send email
$smtpserverport = 8111 ;

#*======^ END CONSTANTS  ^======

#*======v CALCULATED VARIABLES v======
<# profile search path
c:\usr\home\db\MyAccount-prof.ps1
\\Server3.domain.com\USR$\MyAccount\dev\ps\scripts
\\Server0.domain.com\SITEc_FS\scripts
\\Server0\e$\scripts\
#>

<# 10:59 AM 2/18/2015 below test-paths, are DEAD SLOW failingout of non-avail shares
better to test the domain & pc name, and switch to only the approp for the box, b4 actually testing
if (!(Test-Path $LocalInclDir)) {
  if (Test-Path $InclShareCent) {
    $LocalInclDir = $InclShareCent ; Write-Verbose -verbose "Using CentralShare Includes:$LocalInclDir";
  } elseif ($env:USERDOMAIN -eq "DOMAIN") {
    # add pretest for non-SITEc lab
    if(test-path $inclShareLab){
      $LocalInclDir = $inclShareLab ; Write-Verbose -verbose "Using LabShare Includes:$LocalInclDir";
    } elseif (test-path $InclSIDDirL){
      # and then defer into SITEcshare lab
      $LocalInclDir = $InclSIDDirL ; Write-Verbose -verbose "Using LabShareL13 Includes:$LocalInclDir";
    } else {
      Write-Warning "No available LAB include source found. Exiting..."; Exit;
    };
  } elseif (Test-Path $InclShareL13) {
    $LocalInclDir = $InclShareL13 ; Write-Verbose -verbose "Using L13 Includes:$LocalInclDir";
  } else {
    Write-Warning "No available include source found. Exiting..."; Exit;
  };
};
#>
#($env:USERDOMAIN)
switch ($env:USERDOMAIN){
    $DomLab {
        write-host "$DomLab Domain...";
        ($env:computername)
        switch -regex ($env:COMPUTERNAME){
            "$rgxLabL13Servers$" {
                $LocalInclDir = $InclSIDDirL ; Write-Verbose -verbose "Using LabShareL13 Includes:$LocalInclDir";
              };
            "$rgxLabServers" { 
                $LocalInclDir = $inclShareLab ; Write-Verbose -verbose "Using LabShare Includes:$LocalInclDir";
            };
            default { Write-Warning "No available LAB include source found. Exiting Profile load..."; Exit;};
        } # switch-ntry-E ($env:COMPUTERNAME)
    } # switch-ntry lab
    $DomainWork {
        write-host "$DomainWork Domain...";
        ($env:computername)
        
        switch -regex ($env:COMPUTERNAME){            
            "$rgxProdL13Servers" {
                $LocalInclDir = $InclShareL13 ; Write-Verbose -verbose "Using L13 Includes:$LocalInclDir";
            };
            "$rgxProdServers" { 
                $LocalInclDir = $InclShareCent ; Write-Verbose -verbose "Using CentralShare Includes:$LocalInclDir"; 
            };
            "$MyBoxW" { 
                $LocalInclDir = $LocalInclDir ; Write-Verbose -verbose "Using Existing Local Includes:$LocalInclDir"; 
            };
            default{ Write-Warning "No available PROD include source found. Exiting Profile load..."; Exit;};
        } # switch-ntry-E ($env:COMPUTERNAME)
    } # switch-ntry DOMAIN
    default{
        # no dom at home (workgroup), so just test compnames
        if($MyBox -contains $env:COMPUTERNAME){
            "$($env:computername): home network detected";
            # default spec at top will work: $LocalInclDir="c:\usr\home\db" ;
        }; # if-E
    } # switch-ntry-E default
} # switch-E ($env:USERDOMAIN)

#---------

Write-Verbose -verbose "`$LocalInclDir:$LocalInclDir";
Write-Verbose -verbose "`$LocalInclSIDDir:$LocalInclSIDDir";

Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":LOADING ISE INCLUDES:")
<# 7:43 AM 12/30/2014 interesting, the echo below was what was causing the ugly prompt screwup:
VERBOSE: post MyAccount-incl-home.ps1 
# with a global vari, by the time it executed finally, the $sLoad vari had been swapped to the last include load line, instead of where it started below.
#>
$sLoad=(join-path -path $LocalInclDir -childpath "ise-prof.ps1") ; ;if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;};


# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU0GCK2qp0lH9IT0PrbAqdakV4
# caygggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
# MCwxKjAoBgNVBAMTIVBvd2VyU2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdDAe
# Fw0xNDEyMjkxNzA3MzNaFw0zOTEyMzEyMzU5NTlaMBUxEzARBgNVBAMTClRvZGRT
# ZWxmSUkwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBALqRVt7uNweTkZZ+16QG
# a+NnFYNRPPa8Bnm071ohGe27jNWKPVUbDfd0OY2sqCBQCEFVb5pqcIECRRnlhN5H
# +EEJmm2x9AU0uS7IHxHeUo8fkW4vm49adkat5gAoOZOwbuNntBOAJy9LCyNs4F1I
# KKphP3TyDwe8XqsEVwB2m9FPAgMBAAGjdjB0MBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MF0GA1UdAQRWMFSAEL95r+Rh65kgqZl+tgchMuKhLjAsMSowKAYDVQQDEyFQb3dl
# clNoZWxsIExvY2FsIENlcnRpZmljYXRlIFJvb3SCEGwiXbeZNci7Rxiz/r43gVsw
# CQYFKw4DAh0FAAOBgQB6ECSnXHUs7/bCr6Z556K6IDJNWsccjcV89fHA/zKMX0w0
# 6NefCtxas/QHUA9mS87HRHLzKjFqweA3BnQ5lr5mPDlho8U90Nvtpj58G9I5SPUg
# CspNr5jEHOL5EdJFBIv3zI2jQ8TPbFGC0Cz72+4oYzSxWpftNX41MmEsZkMaADGC
# AWAwggFcAgEBMEAwLDEqMCgGA1UEAxMhUG93ZXJTaGVsbCBMb2NhbCBDZXJ0aWZp
# Y2F0ZSBSb290AhBaydK0VS5IhU1Hy6E1KUTpMAkGBSsOAwIaBQCgeDAYBgorBgEE
# AYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwG
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTwkSA5
# k26zcHumGVTeUfmoSXSJUTANBgkqhkiG9w0BAQEFAASBgIMpSzObmB+dB9gZg0HG
# AChCpCMRcdD4R3Lgo6FDYiTJj7HD8se9/KjGhoDQQyTZ1gobyak/hsnMC28rx0GT
# /BEtQZmr0s6n8JlslSV84qgf+OvI9KTfbyG1Ra5topzJT8sfplLTIGwsB167F5oa
# GsUqJ2Jt18CZ0HLjbgD6GIgE
# SIG # End signature block
