#*================v NON-ADMIN : C:\Users\MyAccount\Documents\WindowsPowerShell\profile.ps1 v================
# ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\MyAccount\Documents\WindowsPowerShell\profile.ps1

# NON-ADMIN acct $profile.CurrentUserAllHosts loc
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\MyAccount\Documents\WindowsPowerShell\profile.ps1
# NON-ADMIN acct $profile.CurrentUserCurrentHost loc
#C:\Users\kadriets\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1


# notepad2 $profile.CurrentUserAllHosts ;
#*================================

#*----------V Comment-based Help (leave blank line below) V---------- 

## 
#     .SYNOPSIS
# NON-ADMIN: C:\Users\MyAccount\Documents\WindowsPowerShell\profile.ps1 - 
# My primary non-admin Profile file
# 
#     .NOTES
# Written By: Todd Kadrie
# Website:	http://tinstoys.blogspot.com
# Twitter:	http://twitter.com/tostka
# 
# Additional Credits: [REFERENCE]
# Website:	[URL]
# Twitter:	[URL]
# 
# Change Log.
# 1:44 PM 5/24/2016 shifted storage from uhd => C:\sc\powershell\PSProfileUID
# 11:55 AM 11/5/2015 bring over $Inclpath detect from MyAccount-prof.ps1
# 11:37 AM 1/7/2015 replaced concat paths with join-path
# 9:37 AM 1/7/2015 codified local & central/remote share variable names
# 12:57 PM 12/31/2014 fixed missing \ on SITEc share
# 10:52 AM 12/31/2014 add support for SITEc shared scripts on fileshare "\\Server0.domain.com\SITEc_FS\scripts"
#12:40 PM 12/30/2014 updated to cover lab paths
# 8:28 AM 12/30/2014 minor format cleanup
# 2:00 PM 12/23/2014 split out into include files, shifted primary profile into include MyAccount-prof.ps1 as well


<# creation cmds
# create the CurrentUserAllHosts file:
New-Item -path $PROFILE.CurrentUserAllHosts -ItemType file -Force ; notepad2 $PROFILE.CurrentUserAllHosts ;
# create the CurrentUserCurrentHost (default) file, type:
New-Item -path $profile -itemtype file -force; notepad2 $PROFILE.CurrentUserCurrentHost ;
#>

write-host -foregroundcolor gray "$((get-date).ToString("HH:mm:ss")):====== EXECUTING: $(Split-Path -Leaf ((&{$myInvocation}).ScriptName) ) ====== " ; 

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
$rgxLabL13Servers="^Server2D$" ; 
$rgxLabServers="^(SITE|SITE|SITE)XXX[0-9][0-9]$" ;
$rgxProdL13EdgSrvrs="^Server1$" ;
$rgxLabL13EdgSrvrs="^Server1$" ;

# Domain contstants
$DomainWork = "DOMAIN";
$DomHome = "DOMAIN";
$DomLab="DOMAIN";
# edge boxes show dom as themselves: ($env:USERDOMAIN) =Server0
# if the ($env:USERDOMAIN) -eq the ($env:COMPUTERNAME), you're on a non-domain-joined box
$DomL13EdgProd=$rgxProdL13EdgSrvrs ;
$DomL13EdgLab=$rgxLabL13EdgSrvrs ;

#$LocalInclDir="c:\usr\home\db" ;
# 1:46 PM 5/24/2016 shift to github sc locations
$LocalInclDir="C:\sc\powershell\PSProfileUID" ;
#$LocalInclSIDDir = "c:\usr\work\ps\scripts";
$LocalInclSIDDir = "C:\sc\powershell\PSProfileSID";
# distrib shares:
$InclShareCent = "\\Server3.domain.com\USR$\MyAccount\dev\ps\scripts";
$InclShareL13 = "\\Server0.domain.com\SITE_FS\scripts";
$inclShareLab = "\\Server0\c$\scripts";
# 12:53 PM 1/12/2015 lab SITEc support
$InclSIDDirL = "\\SITEcfs.domain.com\SITEcFileShare\scripts";

$smtpserver = 'Server0' # SMTP server you want to use to send email
$smtpserverport = 25 ;

#*======^ END CONSTANTS  ^======

#*======v CALCULATED VARIABLES v======
<# profile search path
c:\usr\home\db\MyAccount-prof.ps1
\\Server3.domain.com\USR$\MyAccount\dev\ps\scripts
\\Server0.domain.com\SITEc_FS\scripts
\\Server0\e$\scripts\
#>

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
        <# 10:05 AM 11/5/2015 update regex to cover 630's & 4200 OWAS boxes#>
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


#*======^ END CALCULATED VARIABLES ^======

Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":LOADING INCLUDES:")
$sLoad=(join-path -path $LocalInclDir -childpath "MyAccount-prof.ps1") ; ;if (Test-Path $sLoad) {  Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad ; if ($bDebug) { Write-Verbose -verbose "Post $sLoad"};} else {  Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:"+ $sLoad + " EXITING...") ; exit;};

