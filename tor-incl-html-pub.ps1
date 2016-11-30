#tor-incl-html.ps1
#incl-html-TOR.ps1

# DOMAINspecific HTML/CSS-related includes boilerplate

# spec cssfile (used with -cssuri $CssFile)
$CssFile="tor-incl-base.css"
#11:24 AM 2/4/2015 updated for distr
#$sHTMLPreLogo=".\header.gif"
$sHTMLPreLogo=".\DOMAINlogo.gif"

#*----------------v HEADER block v----------------
$sHTMLhead= @"

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$PageTitle</title><meta http-equiv="refresh" content="120" />
</head>

"@
#*----------------^ HEADER block ^----------------


