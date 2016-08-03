#tor-incl-html-TOR-logo-graybar.ps1
#incl-html-TOR-logo-graybar.ps1

<# 
	.SYNOPSIS
tor-incl-html-TOR-logo-graybar.ps1 - Site-specific HTML/CSS-related includes boilerplate

	.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com


Change Log
* 8:37 AM 10/30/2015 trimmed length of --/==
* 8:58 AM 3/11/2015 updated the X log to pull from X.com web, no more attempts to host

	.DESCRIPTION
Default html formatting include settings.
Encapsulates $CSSFile path & $sHTMLPreLogo company logo pointer
Also includes the settings within $sHTMLhead (page top & <head>...</head> tags)
This gets pulled in via kadritss profile, into every script. 
If you run a script from another account, you'll need to explicitly add this include in manually

	.PARAMETER  <Parameter-Name>
<parameter comment>

	.INPUTS
None. Does not accepted piped input.

	.OUTPUTS
None. Returns no objects or output.
System.Boolean
				True if the current Powershell is elevated, false if not.
[use a | get-member on the script to see exactly what .NET obj TypeName is being returning for the info above]

	.EXAMPLE
.\[SYNTAX EXAMPLE]
[use an .EXAMPLE keyword per syntax sample]

	.LINK
< name of a related topic, one .LINK keyword per related topic. Can also include a URI to online help (start with http:// or https://)>
*------^ END Comment-based Help  ^------ #>

# Site-specific HTML/CSS-related includes boilerplate

# spec whether to inline the css code, or <link-include it>
# move this out to the main sub code
#$CSSInline=$true;

<# SAMPLE USAGE/LOAD/CALL
# ----- V INCLUDES BLOCK V------
$ScriptDir = (Split-Path -parent $MyInvocation.MyCommand.Definition) + "\";
write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ":LOADING INCLUDES:");
$sLoad=($ScriptDir + "incl-html-TOR.ps1") ; if (test-path $sLoad) {
  write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad } else {write-warning ((get-date).ToString("HH:mm:ss") + ":MISSING"+ $sLoad + " EXITING...") ; exit};
# ----- ^ INCLUDES BLOCK ^------
#>


$CssFile="tor-incl-base-graybar-borders.css"
#"tor-incl-base-graybar.css"




$sHTMLPreLogo=".\Xlogo.gif"
#*------v HEADER block v------
if($CSSInline){
  #$CSSInlineText=(get-content (join-path $LocalInclDir $CssFile))
  # shift to $Scriptdir
  $CSSInlineText=(get-content (join-path $ScriptDir $CssFile))
$sHTMLhead= @"

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$PageTitle</title><meta http-equiv="refresh" content="120" />
<style type=”text/css”>
<!–
$CSSInlineText
–>
</style>
</head>

"@
} else {
$sHTMLhead= @"

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$PageTitle</title><meta http-equiv="refresh" content="120" />
</head>

"@
} # if-E
#*------^ HEADER block ^------

if($bDebug){
  # dump out the $htmlHead value
  write-host "Dumping shtmlhead..."
  $sHTMLhead | out-file .\shtmlhead-test.html
};  # if-E
