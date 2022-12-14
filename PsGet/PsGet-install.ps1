# PsGet-install.ps1

# debug: . \\server\scripts\PsGet-install.ps1 -path "\\server\scripts\PsGet.psm1" -whatif -showdebug ;

<# 
    .SYNOPSIS
    PsGet-install.ps1 - Copy PsGet.psm1 to CU Modules directory (local-source-only)
    .NOTES
    Note, generally requires an up to date signature to function (if remotesigned is on).
    Written By: Tweaked local-only variant of the stock PsGet installation code - tweaked by Todd Kadrie
    Website:	https://github.com/psget/psget
    Change Log
    * 12:28 PM 7/18/2016 disabled the load-module PsGet and output a comment with syntax (suppress error)
    * 12:19 PM 7/18/2016 seems to work - to install, but attempts to import-module within the script fail. Though the installed .psm1 works fine afterward. 
    * 11:19 AM 7/18/2016 initial build
    .DESCRIPTION
    PsGet-install.ps1 - Copy PsGet.psm1 to CU Modules directory
    .PARAMETER  Path
    Path to .PSM1 file to be installed in CU Modules dir[c:\path-to\PsGet.psm1]
    .PARAMETER showDebug

    .PARAMETER  whatif
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    .LINK
#>


Param(
    [Parameter(Position=0,Mandatory=$True,HelpMessage="Local Path to PsGet.psm1 file [c:\path-to\PsGet.psm1]")]
    [ValidateNotNullOrEmpty()]
    [string]$Path
    ,[Parameter(HelpMessage='Debugging Flag [-showdebug]')]
    [switch] $showDebug
    ,[Parameter(HelpMessage='Whatif Flag  [-whatif]')]
    [switch] $whatIf
) # PARAM BLOCK END

if ($Whatif){$bWhatif=$true ; Write-Verbose -Verbose:$true "`$Whatif is $true (`$bWhatif:$bWhatif)" ; };

#*------v Function Install-PsGetLocal v------
function Install-PsGetLocal {
    # 11:56 AM 7/18/2016 ren -url to -path (for local use)
    param (
        [string]
        # URL to the respository to download PSGet from
        $path
    ) ; 

    # 11:55 AM 7/18/2016 validate the object is a leaf and ends with PsGet.psm1
    if (get-childitem -path $Path |?{$_.fullname -like '*\PsGet.psm1'}){
    
        $ModulePaths = @($env:PSModulePath -split ';') ; 
        # $PsGetDestinationModulePath is mostly needed for testing purposes,
        if ((Test-Path -Path Variable:PsGetDestinationModulePath) -and $PsGetDestinationModulePath) {
            $Destination = $PsGetDestinationModulePath ; 
            if ($ModulePaths -notcontains $Destination) {
                Write-Warning 'PsGet install destination is not included in the PSModulePath environment variable' ; 
            } ; 
        } else {
            $ExpectedUserModulePath = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath WindowsPowerShell\Modules ; 
            $Destination = $ModulePaths | Where-Object { $_ -eq $ExpectedUserModulePath } ; 
            if (-not $Destination) {
                $Destination = $ModulePaths | Select-Object -Index 0 ; 
            } ; 
        } ; 
        New-Item -Path ($Destination + "\PsGet\") -ItemType Directory -Force | Out-Null ; 
        #Write-Host ('Downloading PsGet from {0}' -f $url) ; 
        #Get-File -Url $url -SaveToLocation "$Destination\PsGet\PsGet.psm1" ; 
        # 11:54 AM 7/18/2016 pull from local path
        Copy-Item -Path $path "$Destination\PsGet\PsGet.psm1" ;
        $executionPolicy = (Get-ExecutionPolicy) ; 
        $executionRestricted = ($executionPolicy -eq "Restricted") ; 
        if ($executionRestricted) {
            Write-Warning @" 
Your execution policy is $executionPolicy, this means you will not be able import or use any scripts including modules.
To fix this change your execution policy to something like RemoteSigned.
PS> Set-ExecutionPolicy RemoteSigned ; 
For more information execute: ; 
PS> Get-Help about_execution_policies ; 
"@ ; 
    } ; 
    if (!$executionRestricted) {
        # ensure PsGet is imported from the location it was just installed to
        # 12:25 PM 7/18/2016 craps out on local from unc in lab, skip the load below
        #Import-Module -Name $Destination\PsGet ; 
        write-host "To load PsGet, type: Import-Module PsGet " ; 
    } ; 
    Write-Host "PsGet is installed and ready to use" -Foreground Green ; 
    Write-Host @"
USAGE: ; 
PS> import-module PsGet ; 
PS> install-module PsUrl ; 
For more details: ; 
get-help install-module ; 
Or visit http://psget.net
"@ ;
 
    } else {
         write-error "$((get-date).ToString("HH:mm:ss")):INVALID PATH TO PSGET.PSM1!:$($Path)"; break ;
    } # if-E ; 
} #*------^ END Function Install-PsGetLocal ^------; 


Install-PsGetLocal -path $Path ; 
