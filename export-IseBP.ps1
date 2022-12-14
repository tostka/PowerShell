# export-IseBP.ps1

<# 
.SYNOPSIS
export-IseBP.ps1 - export currently set Line BreakPoints from open script in ISE, into matching XML
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
9:06 AM 6/3/2016 - updated to spot blank doc as Untitled1.ps1
8:53 AM 6/3/2016 - added filtering for JUST CURRENT TAB! (vs handling BP's for all tabs)
8:34 AM 6/3/2016 - initial version
.DESCRIPTION
export-IseBP.ps1 - export currently set Line BreakPoints from open script in ISE, into matching XML
.INPUTS
None. Leverages open file open in ISE to determine source
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.\export-IseBP.ps1 
.LINK
#>

# just noticed on cold load, blank doc isn't blank, has: $psise.CurrentFile.FullPath: C:\sc\powershell\PSScripts\Untitled1.ps1
if ($psise -AND $psise.CurrentFile.FullPath -notmatch "^.*\\Untitled1.ps1$") { 
  $xFname=$psise.CurrentFile.FullPath.replace("ps1","xml").replace(".","-BP.") ; ;
   "Creating BP file:$($xFname)" ;
   # 8:50 AM 6/3/2016 nope, below gets ALL TABS, 
  #Get-PSBreakpoint |?{$_.line}| Export-Clixml -Path $xFname;
  # 8:50 AM 6/3/2016 need to filter for the current tab
  $xBPs= get-psbreakpoint |?{$_.Script -eq $($psise.currentfile.fullpath) -AND ($_.line)} ; 
  $xBPs | Export-Clixml -Path $xFname ; 
  "$(($xBPs|measure).count) Breakpoints exported to $xFname" 
} else {  write-warning "This script only functions within PS ISE, with a script file open for editing" ; };


