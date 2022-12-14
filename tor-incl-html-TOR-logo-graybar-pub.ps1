#tor-incl-html-TOR-logo-graybar.ps1
#incl-html-TOR-logo-graybar.ps1

# DOMAINspecific HTML/CSS-related includes boilerplate

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


# spec cssfile (used with -cssuri $CssFile)
#$CssFile="TOR-base-graybar.css"
#11:24 AM 2/4/2015 updated for distr
$CssFile="tor-incl-base-graybar-borders.css"
#"tor-incl-base-graybar.css"




#$sHTMLPreLogo=".\header.gif"
$sHTMLPreLogo=".\DOMAINlogo.gif"
#*----------------v HEADER block v----------------
if($CSSInline){
  #$CSSInlineText=(get-content (join-path $LocalInclDir $CssFile))
  # shift to $Scriptdir
  $CSSInlineText=(get-content (join-path $ScriptDir $CssFile))
$sHTMLhead= @"

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$PageTitle</title><meta http-equiv="refresh" content="120" />
<style type=�text/css�>
<!�
$CSSInlineText
�>
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
#*----------------^ HEADER block ^----------------

if($bDebug){
  # dump out the $htmlHead value
  write-host "Dumping shtmlhead..."
  $sHTMLhead | out-file .\shtmlhead-test.html
};  # if-E

