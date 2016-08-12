#tor-incl-html.ps1
#incl-html-TOR.ps1

#*------V Comment-based Help (leave blank line below) V------ 

<# 
	.SYNOPSIS
tor-incl-html.ps1 - Default html formatting include settings

	.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com


Change Log
* 8:37 AM 10/30/2015 trimmed length of --/==
* 8:58 AM 3/11/2015 updated the log to pull from web, no more attempts to host

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

# HTML/CSS-related includes boilerplate

# spec cssfile (used with -cssuri $CssFile)
$CssFile="tor-incl-base.css"
$sHTMLPreLogo="http://www.site.com/Style%20Library/vdir/images/logo.gif"

#*------------v HEADER block v------------
$sHTMLhead= @"

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$PageTitle</title><meta http-equiv="refresh" content="120" />
</head>

"@
#*------------^ HEADER block ^------------
