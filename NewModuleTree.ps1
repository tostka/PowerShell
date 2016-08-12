# NewModuleTree.ps1
# debug: Clear-Host ;  .\NewModuleTree.ps1 -ModuleName INIHandler -ModuleDesc "Functions to Read/Write INI files (via Hash/OrderedDictionary)" -showdebug -whatif ;
# debug: populating batch tree BatScripts: .\NewModuleTree.ps1 -ModuleName VBCode -ModuleDesc "Visual Basic Projects" -showdebug -whatif ;

  <# 
  .SYNOPSIS
  NewModuleTree.ps1 - SC GitHub new module sub tree creation script
  .NOTES
  Written By: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka
  Change Log
  # 10:00 AM 7/21/2016 splice in support for new c:\sc\vb tree
  # 8:52 AM 7/21/2016 splice in support for new c:\sc\batch tree
  # 11:08 AM 6/10/2016 fixed typo, working
  # 9:48 AM 6/10/2016 added text-edting/replace of stock strings in templates woth ModuleName & ModuleDescription
  # 9:32 AM 6/10/2016 minor typo fixes, functional
  # 9:21 AM 6/10/2016 rewrote to use template file copies in root of C:\sc\powershell\
  # 7:31 AM 6/10/2016 ren: NewModuleTree.ps1 => NewModuleTree.ps1 (there's a new-Module native cmdlet, don't want risk of name overlap) & updated help
  # 2:37 PM 5/27/2016 - another wacky rewrite, path checking, self-path dyn generation, pre-path testing, and splat-based New-ModuleManifes
  # 2:01 PM 5/24/2016 - else'd $whatif, to force $bwhatif false on non-whatif
  # 9:07 AM 5/24/2016 - added ModuleName param support
  # 9:02 AM 5/24/2016 - added trailing symlink ref and next steps.
  # 8:25 AM 5/24/2016 - updated for Proc repos
  .DESCRIPTION
  NewModuleTree.ps1 - SC GitHub new module sub tree creation script
  Powershell module support files & include files emphasis, but written for ahk & proc non-PS SC tracking subdirs as well.
  .PARAMETER ModuleName
  New Module Name [string]
  .PARAMETER ModuleDesc
  New Module Description [string]
  .PARAMETER whatif
  Whatif Switch Flag
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  .\NewModuleTree.ps1
  Create specified sub-tree as per settings in NewModuleTree.ps1 file
  .EXAMPLE
  .\NewModuleTree.ps1 -whatif
  Create specified sub-tree as per settings in NewModuleTree.ps1 file
  .LINK
  #>

Param(
    [Parameter(HelpMessage='New Module Name [string]')]
    [string] $ModuleName
    ,[Parameter(HelpMessage='New Module Description [string]')]
    [string] $ModuleDesc
    ,[Parameter(HelpMessage='Debugging Flag [-showDebug]')]
    [switch] $showDebug
    ,[Parameter(HelpMessage='Whatif Flag  [-whatif]')]
    [switch] $whatIf
) # PARAM BLOCK END

#region INIT; # ------ 

$Author = 'Todd Kadrie' ; 
$cModuleName = 'PrcEx' ; 
$cModuleDesc = 'Procedure module to hold general MS Exchange procs' ; 
$scRoot="c:\sc" ; 
$sLowestPsVers='2.0' ; 

# pick up the bDebug from the $ShowDebug switch parameter
# SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
if ($ShowDebug) {$bDebug=$true ; $DebugPreference = "Continue" ; write-debug "(`$ShowDebug:$ShowDebug ;`$bDebug:$bDebug ;`$DebugPreference:$DebugPreference)" ; };
if ($Whatif){$bWhatif=$true ; Write-Verbose -Verbose:$true "`$Whatif is $true (`$bWhatif:$bWhatif)" ; }
else {$bWhatif=$false} ; 
if($bdebug){$ErrorActionPreference = 'Stop' ; write-debug "(Setting `$ErrorActionPreference:$ErrorActionPreference;"};
# scriptname with extension
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ; 
# Clear error variable
$Error.Clear() ; 

if (!$ModuleName) {
    # if no param, go hardcoded
    $ModuleName = $cModuleName ; 
    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Using defaulted `$ModuleName: $($ModuleName)" ; 
} else {
    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):`$ModuleName specified: $($ModuleName)" ; 
}; 

if (!$ModuleDesc) {
    # if no param, go hardcoded
    $ModuleDesc = $cModuleDesc ; 
    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Using defaulted `$ModuleDesc: $($ModuleDesc)" ; 
} else {
    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):`$ModuleDesc specified: $($ModuleDesc)" ; 
}; 

#$Path = "C:\sc\proc\$($ModuleName)"
# 9:08 AM 5/24/2016 drop the root up a level, avoid over nesting
#$Path = "C:\sc\proc\"

# 1:16 PM 5/27/2016determine path on split
# first validate we're in the sc tree
#if((get-location).path.tostring() -match "^C:\\sc\\(ahk|powershell|proc).*$"){
# 8:53 AM 7/21/2016 update for batch
# 10:00 AM 7/21/2016 update for vb
if((get-location).path.tostring() -match "^C:\\sc\\(ahk|powershell|proc|batch|vb).*$"){
    # then split the path and cd to the 2 element (works regardless of whatlevel issued at)
    #cd -path "$(c:\sc\$((get-location).path.tostring().split("\")[2]))" -ea stop ; 
    cd -path "$($scRoot)\$((get-location).path.tostring().split("\")[2])" -ea stop ; 
    $Path = (get-location).path.tostring() ;   
    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):User `$Path:$Path" ;   
}else{
    write-error "$((get-date).ToString("HH:mm:ss")):$((get-location).path) is not a directory within the $($scRoot) tree";
} ; 

# Create the module and private function directories
$subds="Private;Public;en-US".split(";") ; 
if(!(test-path -path "$Path\$ModuleName")) {mkdir $Path\$ModuleName -whatif:$($bWhatif) -ea stop }; 
if(!(test-path -path "$Path\Tests")){mkdir $Path\Tests -whatif:$($bWhatif) -ea stop ; } ;      # Tests scripts dir
<#
mkdir $Path\$ModuleName\Private -whatif:$($bWhatif) ; # for private un-exported functions/varis (if using psm1 as an include file)
mkdir $Path\$ModuleName\Public -whatif:$($bWhatif) ;  # for public exported functions/varis (if using psm1 as an include file)
mkdir $Path\$ModuleName\en-US -whatif:$($bWhatif) ;   # For about_Help files
#>
if(!$bWhatif){
    foreach ($subd in $subds){
        if(!(test-path -path "$Path\$ModuleName\$subd")){
            mkdir "$Path\$ModuleName\$subd" -whatif:$($bWhatif) -ea stop ; 
        } else {
            write-host -foregroundcolor yellow "Path: $("$Path\$ModuleName\$subd") already exists." ; 
        } ; 
    } # loop-E
    #Create the module and related files - for non-procedural sc repositories
    if(!($Path -match "^.*\\proc\\.*$")){
        <# 1:43 PM 5/27/2016 orig code
        New-Item "$Path\$ModuleName\$ModuleName.psm1" -ItemType File -whatif:$($bWhatif) ;
        New-Item "$Path\$ModuleName\$ModuleName.Format.ps1xml" -ItemType File -whatif:$($bWhatif) ;
        New-Item "$Path\$ModuleName\en-US\about_$($ModuleName).help.txt" -ItemType File -whatif:$($bWhatif) ;
        New-Item "$Path\Tests\$($ModuleName).Tests.ps1" -ItemType File -whatif:$($bWhatif) ;
        #>
        $mdir=join-path -path "$Path" -childpath "$ModuleName" ; 
        $mtFiles="XXMODNAMEXX.psm1;XXMODNAMEXX.Format.ps1xml;about_XXMODNAMEXX.help.txt;XXMODNAMEXX.Tests.ps1".split(";") ; 
        foreach ($mtF in $mtFiles){
            $tFile=(join-path -path $mdir -childpath $mtF.replace("XXMODNAMEXX","$($ModuleName)")) ; 
            if(!(test-path -path $tFile )){
                #New-Item "$($tFile)" -ItemType File -whatif:$($bWhatif) ;
                # 8:58 AM 6/10/2016 switch to copied in templates
                # C:\sc\powershell\MODULE.psm1_template                
                switch -regex ($tFile){
                    "^.*\\.*.psm1$" {
                        copy-item -path "$Path\MODULE.psm1_template" -Destination "$($tFile)" ;
                        # 9:42 AM 6/10/2016 replace out all '[MODULENAME]'=> "$($ModuleName)"
                        # 10:00 AM 6/10/2016 have to escape the closing square bracket, -replace works with regex syntax, and they are char set indicators
                        (Get-Content "$($tFile)") |
                            Foreach-Object { $_ -replace '\[MODULENAME\]', "$($ModuleName)" } |
                                Set-Content -path "$($tFile)"
                    } 
                    "^.*\\.*.\.help\.txt$" {
                        copy-item -path "$Path\about_MODULE.help.txt_template" -Destination "$($tFile)" ;
                    }
                    "^.*\\.*.\.(Format\.ps1xml|Tests\.ps1)$" {
                        New-Item "$($tFile)" -ItemType File -whatif:$($bWhatif) ;
                    }
                    default {
                        write-error "$((get-date).ToString("HH:mm:ss")):INVALID FILE SPEC!";
                    }
                } # swtch-E ; 
                
            } else { 
                write-host -foregroundcolor yellow "File: $($tFile) already exists." ; 
            }
        } ; 
        
        # 9:08 AM 6/10/2016 add to copy in the README.MD
        copy-item -path "$Path\Readme.md_template" -Destination "$(join-path -path $mdir -childpath "Readme.md")" ;
        # 9:42 AM 6/10/2016 replace out all '[MODULENAME]'=> "$($ModuleName)"
        # 10:54 AM 6/10/2016 stack -replace's
        
        (Get-Content -path "$(join-path -path $mdir -childpath "Readme.md")") |
            Foreach-Object { 
                $_ -replace '\[MODULENAME\]',"$($ModuleName)" -replace '\[MODULEDESC\]',"$($ModuleDesc)" 
            } | Set-Content -path "$(join-path -path $mdir -childpath "Readme.md")" -force ; 
        
        
        <#        
        $lns=(Get-Content -path "$(join-path -path $mdir -childpath "Readme.md")") ;
        foreach ($ln in $lns){
            $upd+= $ln -replace '\[MODULENAME\]',"$($ModuleName)" -replace '\[MODULEDESC\]',"$($ModuleDesc)" 
        } ; 
        $upd | out-string ;
        $upd | Set-Content -path "$(join-path -path $mdir -childpath "Readme.md")" -force  ; 
        #>
        

        <# 1:31 PM 5/27/2016 orig psv3 code
        New-ModuleManifest -Path $Path\$ModuleName\$ModuleName.psd1 `
                           -RootModule $Path\$ModuleName\$ModuleName.psm1 `
                           -Description $ModuleDesc `
                           -PowerShellVersion 3.0 `
                           -Author $Author `
                           -FormatsToProcess "$ModuleName.Format.ps1xml"
                           -whatif:$($bWhatif) ;
       #>
       # 2:12 PM 5/27/2016 whatif isn't supported with below
       <#
       New-ModuleManifest -Path $mdir\$ModuleName.psd1 `
                           -RootModule $mdir\$ModuleName.psm1 `
                           -Description $ModuleDesc `
                           -PowerShellVersion $sLowestPsVers `
                           -Author $Author `
                           -FormatsToProcess "$ModuleName.Format.ps1xml" ; 
        #>
        # 9:17 AM 6/10/2016 updated to ordered Psv2-compatible Dict
        if($host.version.major -ge 3){
            $spltManif=[ordered]@{
                Dummy = $null ; 
            } ;
        } else {
            $spltManif= New-Object Collections.Specialized.OrderedDictionary ; 
        } ;
        
        If($spltManif.Contains("Dummy")){$spltManif.remove("Dummy")} ; 
        # also, less code post-decl to populate the $hash with fields, post creation:
        $spltManif.Add("Path",$("$mdir\$ModuleName.psd1") ) ; 
        $spltManif.Add("RootModule",$("$mdir\$ModuleName.psm1") ) ; 
        $spltManif.Add("Author",$($Author)) ; 
        $spltManif.Add("CompanyName","(Private)") ; 
        $spltManif.Add("Description",$($ModuleDesc)) ; 
        $spltManif.Add("PowerShellVersion",$($sLowestPsVers)) ; 
        # 2:34 PM 5/27/2016 no drop it, throws errors if it's not populated:
        #$spltManif.Add("FormatsToProcess",$("$ModuleName.Format.ps1xml")) ; 
        New-ModuleManifest @spltManif

    } else {
        write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Procedural module, skipping .ps(d|m)1 creation" ; 
    }# if-E ; 

    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")): Copy the public/exported functions into the public folder, private functions into private folder. `nEdit the readme.md file.`nEdit the .psm1 file`nEdit the about_*help.txt file"
    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Where approp. create symlinks with the syntax:`nMKLINK [[/D] | [/H] | [/J]] Link Target`n# hard-link dir symlink:`n cmd /c mklink /j c:\usr\work\exch\scripts $scRoot\powershell\ExScripts"


} else { 
    write-verbose -verbose:$true  "(whatif pass: skipping subdir & file creation...)" ;
}



# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU3ejsbjwkPxH1xuB/EYUHFwfO
# TSWgggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQOK+gC
# EyRR7zrVOuH32fgLrd9gcjANBgkqhkiG9w0BAQEFAASBgFooWVpvUcgkVC6zmCCB
# gghEPbCP/GthD7feYkuNZf1KwDWILhwXLltj02oPQK94D774ZoO3aYVotYZjTksK
# dxglY9XfxxNV6XlT6yhNWsHvUinjbEdDw5DzeBvhrvPuOiEg3leWbIbz0yS2Sxpq
# 6QnRAJczgTD8C7y5zCpU7jCD
# SIG # End signature block
