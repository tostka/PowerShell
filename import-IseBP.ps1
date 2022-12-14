# import-IseBP.ps1

<# 
.SYNOPSIS
import-IseBP.ps1 - import & set Line BreakPoints from matching XML for file currently open in the ISE
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
10:53 AM 6/7/2016 - added pretest for existing file before import attempt
9:10 AM 6/3/2016 - updated to only purge existing line BPs in current tab (not all tabs)
9:06 AM 6/3/2016 - updated to spot blank doc as Untitled1.ps1
8:34 AM 6/3/2016 - initial version
.DESCRIPTION
import-IseBP.ps1 - import & set Line BreakPoints from matching XML for file currently open in the ISE
.INPUTS
None. Leverages open file open in ISE to determine source
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.\import-IseBP.ps1 
.LINK
#>

# just reuse the existing open filename to get the xml
# 9:07 AM 6/3/2016just noticed on cold load, blank doc isn't blank, has: $psise.CurrentFile.FullPath: C:\sc\powershell\PSScripts\Untitled1.ps1
if ($psise -AND $psise.CurrentFile.FullPath -notmatch "^.*\\Untitled1.ps1$") { 
    $iFname=$psise.CurrentFile.FullPath.replace("ps1","xml").replace(".","-BP.") ; ;
    if(test-path -path $iFname){
    "*Importing BP file:$($iFname) and setting specified BP's for open file $($psise.CurrentFile.FullPath)" ;
    # purge any existing Line BP's
    # 9:09 AM 6/3/2016 --- IN THE CURRENT TAB!
    if($eBP=get-psbreakpoint |?{$_.Script -eq $($psise.currentfile.fullpath) -AND ($_.line)} ){ $eBP | remove-PsBreakpoint } ;  
    Import-Clixml -path $iFname | %{set-PSBreakpoint -script $_.script -line $_.line ; $iBP++} ; 
    "$($iBP) Breakpoints imported and set as per $($xFname)" 
    } else { 
        write-host "No matching BP.xml file found ($iFname)" ; 
    }  ; 
} else {  write-warning "This script only functions within PS ISE, with a script file open for editing" ; };


