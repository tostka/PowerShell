#Get-IMDBSearch.ps1 - Imdb title lookup function, uses text/html parsing. Posted by kristofdba
<#
.SYNOPSIS
Get-IMDBSearch.ps1 - Interactive Imdb title lookup function, uses xml (html) parsing. Returns and lists closest matches in menu, then returns details of selected choice
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Inspired by code By: Unknown (posted by kristofdba) & jerdub1993's xml element parsing example
Website:	https://kristofdba.wordpress.com/2013/03/05/imdb-powershell-function/
            https://www.reddit.com/r/PowerShell/comments/61kib6/query_imdb_for_movies_and_information/
Change Log
# 9:20 PM 9/12/2017 shifted a bunch of items to 2-stage: they were passing and handing back blank values, when matches failed. 
# 9:58 PM 6/6/2017 added code to support both hm & m duration (and cases where both aren't present), also added conversion of hm duration to mins
* 9:57 PM 6/4/2017 added out-null to array check, added support for year-only release dates (oldest items don't have full dates),
made SummaryLine dynamic, 3-line & 4-line to accommodate items wo mpaa ratings, added if/then tests on a variety of components 
(many fields drop completely from the page), added $sQueryMinInterval and dawdle to ensure google qrys are at least that 
far apart (30s), untype $Title, to let it acommoodate objects & arrays, making ratingvalue test (optional field), 
replacing -like with rgx -match
* 11:17 PM 6/3/2017 fixed the Released date (regex out actual date, dropped the regional designator), 
shifted to outputing a raw data obje in func, and handle formatting post-return, removed 
'field name: strings from data. Played with outputs a bit, added -Full param, to dump full CObj, 
otherwise dumps a summary, completely retooled, largely from scratch, leveraging xml elements parsing. 
* 9:47 AM 6/3/2017 port and cleanup, add pshelp, real param block, otb syntax
.DESCRIPTION
Get-IMDBSearch.ps1 - Interactive Imdb title lookup function, 
uses xml parsing. Returns and lists closest matches in menu, then returns details of selected choice
.PARAMETER  Title
Movie title string
.PARAMETER Full
Display full details (default is short summary)
.PARAMETER ShowDebug
Parameter to display Debugging messages [-ShowDebug switch]
.INPUTS
Accepts Title as an array of objects (default param 0)
.OUTPUTS
Outputs to console
.EXAMPLE
IMDBSearch $movie
Search movie title
.EXAMPLE
IMDBSearch -Title "The Nice Guys" ; 
Search movie title, defaults
.EXAMPLE
IMDBSearch -Title "The Nice Guys" -Full ; 
Search movie title with Full detail output. 
.EXAMPLE
.\Get-IMDBSearch.ps1 -Title "the wrecking crew","the terminator" -full
Run multiple searches, with full output
.LINK
https://kristofdba.wordpress.com/2013/03/05/imdb-powershell-function/
.LINK
https://www.reddit.com/r/PowerShell/comments/61kib6/query_imdb_for_movies_and_information/
#>

Param(
    [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Movie Title [-Title 'instersteller']")]
    [ValidateNotNullOrEmpty()]$Title,
    [Parameter(HelpMessage="Display Full Verbose Movie Info [-Full]")]
    [switch]$Full,
    [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
    [switch] $showDebug
) # PARAM BLOCK END

#region INIT; # ------ 
#*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
# pick up the bDebug from the $ShowDebug switch parameter
# SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
if ($ShowDebug) {$bDebug=$true ; $DebugPreference = "Continue" ; write-debug "(`$ShowDebug:$ShowDebug ;`$bDebug:$bDebug ;`$DebugPreference:$DebugPreference)" ; };
if ($Whatif){$bWhatif=$true ; Write-Verbose -Verbose:$true "`$Whatif is $true (`$bWhatif:$bWhatif)" ; };
if($bdebug){$ErrorActionPreference = 'Stop' ; write-debug "(Setting `$ErrorActionPreference:$ErrorActionPreference;"};

# scriptname with extension
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ; 
$ComputerName = $env:COMPUTERNAME ;
$sQot = [char]34 ; $sQotS = [char]39 ; 

# Clear error variable
$Error.Clear() ; 
#*======^ SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE ^======
#endregion INIT; # ------ 

#*------v Function Get-IMDBSearch v------
Function Get-IMDBSearch  {
    <#
    .SYNOPSIS
    Get-IMDBSearch() - Interactive Imdb title lookup function, uses xml parsing. Returns and lists closest matches in menu, then returns details of selected choice
    .NOTES
    Written By: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka
    Inspired by code By: Unknown, posted by kristofdba,  jerdub1993's xml element parsing code
    Website:	https://kristofdba.wordpress.com/2013/03/05/imdb-powershell-function/
                https://www.reddit.com/r/PowerShell/comments/61kib6/query_imdb_for_movies_and_information/
    Change Log
    * 9:15 PM 6/4/2017 updated where -> ? syntax throughout. 
    made SummaryLine dynamic, 3-line & 4-line to accommodate items wo mpaa ratings, added if/then tests on a 
    variety of components (many fields drop completely from the page), untype $Title, to let it acommoodate objects & arrays
    making ratingvalue test (optional field), replacing -like with rgx -match
    * 11:17 PM 6/3/2017 fixed the Released date (regex out actual date, dropped the regional designator), 
    shifted to outputing a raw data obje in func, and handle formatting post-return, removed 'field name: strings from data. 
    Played with outputs a bit, added -Full param, to dump full CObj, otherwise dumps a summary
    works fairly well, completely retooled, largely from scratch, leveraging xml elements parsing. 
    * 9:47 AM 6/3/2017 port and cleanup, add pshelp, real param block, otb syntax
    .DESCRIPTION
    Get-IMDBSearch() - Interactive Imdb title lookup script & function, 
    uses xml parsing. Returns and lists closest matches in menu, then returns details of selected choice
    .PARAMETER  Title
    Title search string [-Title 'instersteller'
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .INPUTS
    Single Title search string (default param 0)
    .OUTPUTS
    Returns output object to pipeline
    .EXAMPLE
    IMDBSearch "Bladerunner"
    Search movie title, defaults
    .EXAMPLE
    IMDBSearch -Title "The Nice Guys" ; 
    Search movie title, defaults
    .EXAMPLE
    IMDBSearch -Title "The Nice Guys" -Full ; 
    Search movie title with Full detail output. 
    .LINK
    https://kristofdba.wordpress.com/2013/03/05/imdb-powershell-function/
    .LINK
    https://www.reddit.com/r/PowerShell/comments/61kib6/query_imdb_for_movies_and_information/
    #>
    Param(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Title search string [-Title 'instersteller']")]
        [ValidateNotNullOrEmpty()][string]$Title,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) # PARAM BLOCK END

    # typical imdburl: http://www.imdb.com/title/tt1185418/
    $rgxImdbID="tt\d{7,8}" ; # imdbID: tt1185418

    # Uses Google to search for imdb pages, '?btnI' forces immed refresh to the first hit
    #$qryUrlRoot = "http://google.com/search?btnI=1&q=site:imdb.com/title" ; 
    # in this interactive case, we want all of the hits, and will output summary menu from which to pick specific target
    $qryUrlRoot = "http://google.com/search?q=site:imdb.com/title" ; 
    $url = "$($qryUrlRoot) $($Title.trim())" ; 
    
    $webPage = Invoke-WebRequest -Uri $url ; 
    $pageElems = $webPage.AllElements ;  

    # at this point, we have the google hits and need to collect and display them. 
    $matchedhits=($pageElems | ?{$_.class -eq "g"}) | select innertext ; 
    if($matchedhits){
        if($showDebug){write-verbose -verbose:$true  "Processing $($matchedhits.count) matches:"} ;
        $menu = [ordered]@{} ; 
        $menuentries=0 ; 
        $mnuItems = @() ; 
        foreach ($hit in $matchedhits){
            $fields=$hit.innerText.split("`n") ; 

            switch ($fields.Count) {
                "4" {
                    <# TV show 4 liner:
                    Demolition: The Wrecking Crew (TV Series 2015– ) - IMDb
                    www.imdb.com/title/tt4487606/‎Cached
                    SimilarDocumentary · Add a Plot » ... Demolition: The Wrecking Crew. 1h | Documentary 
                    | TV Series (2015– ) · Episode Guide. 3 episodes · Add a Plot » ...
                    #>
                    if($fields[3]){
                        $Summary = "$($fields[3].substring(0,[System.Math]::Min(50, $fields[3].Length)))..." ; 
                    } else {
                        $Summary = "(missing 3rd line)" ; 
                    } ; 
                }
                "5" {
                    <# movie 5-liner
                    The Wrecking Crew (1968) - IMDb
                    www.imdb.com/title/tt0065225/‎Cached
                    Similar  Rating: 5.9/10 - 1,601 votes
                    Action · The count has stolen enough gold to cause a financial crisis in the world 
                    markets so I.C.E. sends in ace spy Matt Helm to stop him. As Matt works alone, ...
                    #>
                    if($fields[3]){
                        $Summary = "$($fields[3].substring(0,[System.Math]::Min(50, $fields[3].Length)))..." ; 
                    } else {
                        $Summary = "(no data)" ; 
                    } ; 

                } 
            } ;  # swtch-E

            if($fields[1] -match $rgxImdbID ){  
                $imdbID=$matches[0] ; 
            } else { 
                write-warning "FAILED TO MATCH IMDBid FOR $(($fields | out-string).trim())`nSelection will not be openable without a valid IMDBid!";
                $imdbID="-" ; 
            } ; 

            if( $fields[0].tostring().trim() -match "^.*\s-\s(Full\sCast\s&\sCrew|Photo\sGallery|Trivia)\s-\sIMDb$" ){
                if($showDebug){Write-Host "Skipped: $($fields[0])..." ; } ; 
            } elseif($fields[0].tostring().trim() -match ".*\s-\s(Trivia|Full\sCast\s&\sCrew|Parents\sGuide|Awards|Plot\sSummary|Company\scredits|FAQ|Plot\skeywords|Photo\sGallery|Taglines|Filming\sLocations|Quotes|News|Synopsis|Soundtracks|External\sReviews|Crazy\sCredits|Connections|TV\sschedule|Release\sInfo|Technical\sSpecifications|Video\sGallery)\s-\sIMDb.*") {
                if($showDebug){Write-Host "Skipped: $($fields[0])..." ; } ; 
            } else {
                $props=[ordered]@{
                    'Title'=$($fields[0].tostring().trim()) ; 
                    'imdbID'=$($imdbID) ; 
                    'Summary'=$($Summary) ; 
                } ; 
                if(!($conflict = $mnuItems | ?{$_.imdbID -eq $props.imdbid} )){
                    $omnuEntry = New-Object PSObject -Property $props ; 
                    $mnuItems += $omnuEntry
                } else { 
                     if($showDebug){write-verbose -verbose:$true "$($props.Title) ($($props.imdbID))`n$($props.Summary)`n dupes existing entry $(($conflict|out-string).trim())" ;} ; 
                } ; 
            }  # if-E; 

        }  # loop-E; 

        # build a cobj of entries with names from the filtered $mnuItems
        write-host -ForegroundColor Yellow "Query: '$($Title)'" ; 
        $mnuItem=0 ; 
        foreach ($mnu in $mnuItems){
            $mnuItem++ ; 
            $menu.Add($($mnuItem),$($mnu.Title)) ; 
            # output 'visible' menu to console
            write-host "$($mnuItem). $($mnu.Title),$($mnu.Summary),$($Mnu.imdbID)" ; 
        } ; 
        # 6:33 PM 6/6/2017 add a final 'abort/exit' menu item in last position if -gt 1 hit
        # prompt for choice - default single-match menu's to the first item
        if($mnuItem -gt 1){
            $mnuItem++ ; 
            $mnuExitText= "[Abort & Exit]" ; 
            $menu.Add($($mnuItem),$($mnuExitText)) ; 
            # output 'visible' menu to console
            write-host "$($mnuItem). $($mnuExitText)" ; 
            [int]$choice = Read-Host 'Enter selection' ; 
        } else { 
            write-host "single-item menu, defaulting"
            [int]$choice = 1 ; 
        } ; 
        $selection = $menu.Item($choice-1) ; 
        # pull imdbID for matching $mnuItems entry for the choice
        if($selection -eq $mnuExitText){
            write-host -ForegroundColor green "Exiting..." ; 
            exit ; 
        } else { 
            $TImdbID = $mnuItems|?{$_.Title -eq $selection} | select -expand imdbID ; 
        } ; 
        # build moviedata hash (ordered under psv2 or 3+)
        if($host.version.major -ge 3){
            $moviedata=[ordered]@{Dummy = $null ; } ;
        } else {
            # psv2 Ordered obj (can't use with new-object -properites)
            $moviedata = New-Object Collections.Specialized.OrderedDictionary ; 
            <# or use an UN-ORDERED psv2 hash:
            $moviedata=@{
                Dummy = $null ; 
            } ;
            #>
        } ;
        If($moviedata.Contains("Dummy")){$moviedata.remove("Dummy")} ; 
        # Populate the $moviedata with fields, post creation (can't create [ordered] without members)
        $hashfields="Type","MpaaRating","Genres","UsrRatingsStmt","UsrRatingsScore","UsrRatingsCount","RuntimeMinutes","Country","Language","Color","Title","Released","Director","Writers","Stars","Description","Storyline","PlotkeywordsKey","imdbID","imdbURL" ; 
        $hashfields |%{$moviedata.Add("$($_)",$($null)) ; } ;  
        
        # now load the target $TImdbID
        $url = "http://www.imdb.com/title/$($TImdbID)" ; 
        write-host -foregroundcolor green "Opening selection: '$($selection)'`nimdbID:$($TImdbID) : $($url)..." ; 
        $webPage = Invoke-WebRequest -Uri $url ; 
        $pageElems = $webPage.AllElements ;  


        # 1:22 PM 6/3/2017 assign to vari
        $SummaryLine=($pageElems | ?{$_.class -eq "subtext"})[0].innertext.split("|").trim() ; 
        
        $moviedata.Title = (($pageElems | ?{$_.itemprop -eq "name"})[0].innerHTML -split "&nbsp;")[0] ;
        
        <# typical summary line examples: 
        #3 dead stoned:           "1h 38min | Documentary, Biography, Comedy | 25 September 2015 (USA) "
        # wonder woman:          "PG-13 | 2h 21min | Action, Adventure, Fantasy | 2 June 2017 (USA) "
        # Battlestar Galactica : "TV-14 | 44min | Action, Adventure, Drama | TV Series (2004–2009) "
        # doctor who 63:         "TV-PG | 45min | Adventure, Drama, Family | TV Series (1963–1989) "
        #3 Les Elkes champions du Cake-Walk (1903): "1min | Short | 1903 (France) "
        # 4count, have 0:MpaaRating ; 1:Duration ; 2:Genre ; 3:releasedate (region)
        # 3 count have 0:Duration ; 1:Genre ; 2:releasedate (region)
        #>

        switch($SummaryLine.count){
            "3" {
                # 3 count have 0:Duration ; 1:Genre ; 2:releasedate (region)
                #3 dead stoned:           "1h 38min | Documentary, Biography, Comedy | 25 September 2015 (USA) "
                #3 Les Elkes champions du Cake-Walk (1903): "1min | Short | 1903 (France) "
                $moviedata.MPAARating = "-" ; # always a blank rating on a 3count
                $moviedata.Genres = $($SummaryLine[1].Trim()) ;
                
                $matches = $null; 
                If ($SummaryLine[2] -match "^TV\sSeries.*$") {
                    $moviedata.Type = "TV Series" ;
                    $matches = $null ; 
                    if($summaryline[2] -match "^TV\sSeries\s\((\d{4}).*"){
                        try {
                            $moviedata.Released = get-date -Year $matches[1] -month 1 -Day 1 -Format "yyyy" ;
                        } catch {
                            $moviedata.Released = "-" ; 
                        } ; 
                    } else {
                        $moviedata.Released = "-";
                    } ; 
                } else { 
                    $moviedata.Type = "Movie" ;
                    # typical movie on [2]:"25 September 2015 (USA)"
                    $matches = $null ; 
                    if($summaryline[2] -match ".*(\d{1,2}\s\w*\s\d{4}).*" ) {

                        # lookabehind the (word) and get-date that captured string
                        try {
                            $SummaryLine[2] -match ".*(?=\s\(\w*\))" ; 
                            $moviedata.Released = get-date $matches[0] -format "MM/dd/yyyy";
                        } catch {
                            $moviedata.Released = "-" ; 
                        } ; 
                    } elseif($summaryline[2] -match "(\d{4})\s\(\w*\)" ) {
                        # lookabehind the (word) and get-date that captured string
                        try {
                            #$moviedata.Released = get-date $matches[1] -format "MM/dd/yyyy";
                            $moviedata.Released = get-date -Year $matches[1] -month 1 -Day 1 -Format "yyyy" ;
                        } catch {
                            $moviedata.Released = "-" ; 
                        } ; 
                    } else {  $moviedata.Released = "-"; }   ; 
                }  ; 


            }# swtch-3-E ; 

            "4" {
                # 4count, have 0:MpaaRating ; 1:Duration ; 2:Genre ; 3:releasedate (region)
                # wonder woman:          "PG-13 | 2h 21min | Action, Adventure, Fantasy | 2 June 2017 (USA) "
                # Battlestar Galactica : "TV-14 | 44min | Action, Adventure, Drama | TV Series (2004–2009) "
                # doctor who 63:         "TV-PG | 45min | Adventure, Drama, Family | TV Series (1963–1989) "
                $moviedata.MPAARating = $SummaryLine[0].tostring().trim() ;
                $moviedata.Genres = $($SummaryLine[2].Trim()) ;

                $matches = $null; 
                If ($SummaryLine[3] -match "^TV\sSeries.*$") {
                    $moviedata.Type = "TV Series" ;
                    $matches = $null ; 
                    if($summaryline[3] -match "^TV\sSeries\s\((\d{4}).*"){
                        try {
                            $moviedata.Released = get-date $matches[1] -format "MM/dd/yyyy";
                        } catch {
                            $moviedata.Released = "-" ; 
                        } ; 
                    } else {
                        $moviedata.Released = "-";
                    } ; 
                } else { 
                    $moviedata.Type = "Movie" ;
                    # typical movie on [3]:"25 September 2015 (USA)"
                    $matches = $null ; 
                    if($summaryline[3] -match ".*(\d{1,2}\s\w*\s\d{4}).*" ) {
                        #lookabehind the (word) and get-date that captured string
                        try {
                            $SummaryLine[3] -match ".*(?=\s\(\w*\))" ; 
                            $moviedata.Released = get-date $matches[0] -format "MM/dd/yyyy";
                        } catch {
                            $moviedata.Released = "-" ; 
                        } ; 
                    } else {  $moviedata.Released = "-"; }   ; 
                } # if-E TV/Movie ; 

            }  # swtch-4-E ; 

        } ; 
       
       # 9:12 PM 9/12/2017force blank to -
       if(!$moviedata.Released){$moviedata.Released = "-" ; } ; 

        # 12:22 PM 6/4/2017 ratingValue is optional, pre-test for presence
        if(($pageElems | ?{$_.class -eq "ratingValue"})){
            $moviedata.UsrRatingsStmt=($pageElems | ?{$_.class -eq "ratingValue"})[0].innerhtml.split('"')[1].tostring().trim() ; 
            $moviedata.UsrRatingsScore=($moviedata.UsrRatingsStmt -split("\sbased\son\s"))[0].tostring().trim() ; 
            $moviedata.UsrRatingsCount=($moviedata.UsrRatingsStmt -split("\sbased\son\s"))[1].tostring().replace(" user ratings","").trim() ; 
        }else{
            $moviedata.UsrRatingsStmt="-" ; 
            $moviedata.UsrRatingsScore="-" ; 
            $moviedata.UsrRatingsCount="-" ; 
        } ; 
        # 1:47 PM 6/4/2017 films freq don't have writers|Dir|Stars, which drops the missing item, and shifts rest up. 
        #     so can't safely be hard coded for position, do them dyn via filter
        $crSum=($pageElems | ?{$_.class -eq "credit_summary_item"})| select innertext ; 
        if($Dir=($crSum|?{$_ -like '*Director:*'}).innerText){
            $moviedata.Director= $Dir.tostring().replace("Director: ","").trim() ; 
        } else {
            $moviedata.Director="-" ; 
        }; 
        if($Writers=($crSum|?{$_ -like '*Writers:*'}).innerText){
            # 2:08 PM 6/4/2017 split out : '| 1 more credit' » 
            if($Writers -match ".*\|.*"){
                $moviedata.Writers= $Writers.tostring().split("|").trim()[0].replace("Writers: ",""); 
            } else { 
                $moviedata.Writers= $Writers.tostring().trim().replace("Writers: ",""); 
            } ; 
            if($Writers -match "Writers:\s.*\|\s\d{1,2}\smore\scredits.*"){ $moviedata.Writers+="..."} ; 
        } else {
            $moviedata.Writers="-" ; 
        };
        if($Stars=($crSum|?{$_ -like '*Stars:*'}).innerText){
            if($stars -match ".*\|.*"){
                $moviedata.Stars= $Stars.tostring().split("|").trim()[0].replace("Stars: ","") ;
            } else {
                $moviedata.Stars= $Stars.tostring().trim().replace("Stars: ","") ;
            }  ; 
        } else {
            $moviedata.Stars="-" ; 
        };
        # 6:15 PM 9/12/2017 2-step it, some come back with no summary_text
        $TempResult = $null ; 
        $TempResult = $pageElems | ?{$_.class -eq "summary_text"} ; 
        if($TempResult){
            $moviedata.Description = $TempResult[0].innertext.tostring().trim() ; 
        } ; 
        if(($moviedata.Description -match "^Add\sa\sPlot\s.*") -OR (!$moviedata.Description)){
            $moviedata.Description = "-" ; 
        } ; 
        # 3:22 PM 6/4/2017 Storyline doesn't always appear, pretest
        if(($pageElems | ?{$_.class -eq "inline canwrap"})){
            $moviedata.Storyline = ($pageElems | ?{$_.class -eq "inline canwrap"})[0].innertext.tostring().trim() ; 
        } else {
            $moviedata.Storyline = "-" ; 
        } ; 
        # 3:28 PM 6/4/2017 looks like Plot Keywords aren't consistently availalble either, pretest
        # 6:10 PM 9/12/2017 it's passing even when there's no proper match, 2 step it and then test
        $TempResult = $null ; 
        $TempResult = $pageElems | ?{$_.class -eq "see-more inline canwrap"} ; 
        #if(($pageElems | ?{$_.class -eq "see-more inline canwrap"})[0].outerText -match "(Plot\sKeywords:\s.*\s)\|\sSee\sAll\s\(\d*\)\s.*" ){
        if($TempResult){
            if($TempResult[0].outerText -match "(Plot\sKeywords:\s.*\s)\|\sSee\sAll\s\(\d*\)\s.*" ){
                $moviedata.PlotkeywordsKey = $matches[1].tostring().trim() ; 
            } ; 
        } ; 
        if(!$moviedata.PlotkeywordsKey){
            $moviedata.PlotkeywordsKey = "-" ; 
        } ; 

        # 8:49 PM 9/12/2017 2-step
        $TempResult = $null ; 
        $TempResult = $pageElems | ?{$_.id -eq "titleDetails"} ; 
        #if(($pageElems | ?{$_.class -eq "see-more inline canwrap"})[0].outerText -match "(Plot\sKeywords:\s.*\s)\|\sSee\sAll\s\(\d*\)\s.*" ){
        if($TempResult){
            if(($TempResult)[0].innertext -match ".*(Country:\s.*)"){
                # also replace out the pipe with comma
                $moviedata.Country=($matches[1] -replace "\s\|\s","," -replace("Country: ","")).tostring().trim(); 
            } ;
        } ;  
        if(!$moviedata.Country){ $moviedata.Country="-" } ; 
        
        # 8:57 PM 9/12/2017 2 stage
        $tempResult = $pageElems | ?{$_.id -eq "titleDetails"} ; 
        if($tempREsult){
            if(($pageElems | ?{$_.id -eq "titleDetails"})[0].innertext -match ".*(Language:\s.*)" ){
                # 8:07 PM 6/4/2017 replace pipe->comma
                $moviedata.Language=$matches[1].tostring().replace("Language: ","").trim() -replace "\s\|\s",","; 
            } ;
        } ; 
        if(!$moviedata.Language){  $moviedata.Language="-" } ; 

        $tempResult= $pageElems | ?{$_.id -eq "titleDetails"} ; 
        if($tempResult){
            if(($tempResult)[0].innertext -match ".*(Color:.*)" ){
                # 9:58 PM 6/6/2017 sub-out pipe for comma
                $moviedata.Color=$matches[0].tostring().replace("Color: ","").trim() -replace("\|\s",","); 
            } ; 
        } ; 
        if(!$moviedata.Color){ $moviedata.Color = "-" } ; 

        $moviedata.imdbID=$($TImdbID) ; 
        $moviedata.imdbURL=$($URL) ; 

        if($duration=($pageElems | ?{$_.itemprop -eq "duration"})){
            # 6:46 PM 6/6/2017 sometimes there are several, and sometimes there's just 1: 169 min, some use mins, some h&m
            $matches=$null ; 
            switch -regex ($duration[-1].innertext) {
                "(\d{1,2}h\s\d{1,2}min)" {
                    if($duration[-1].innertext -match "(\d{1,2}h\s\d{1,2}min)"){
                        $timestamp = $matches[0].tostring().trim().replace(" ","") ; 
                        $moviedata.RuntimeMinutes =  "$([int]($timestamp.split('h')[0])*60 +[int]($timestamp.split('h')[1]).replace('min',''))min" ; 
                    } ; 
                } ; 
                "(\d{1,3}\smin)" {
                    # 1h 27min 
                    if($duration[-1].innertext -match "(\d{1,3}\smin)" ){
                        $moviedata.RuntimeMinutes = $matches[0].tostring().trim().replace(" ","") ;  ; 
                    } ; 
                } ; 
                default{ $moviedata.RuntimeMinutes = "-" } ; 
            } ; 
        } else { $moviedata.RuntimeMinutes = "-" } ; 

        # dump hash into pipeline (formatting should be handled on receiving end, we just do source data in this func :D)
        $moviedata | write-output ; 

    } else { 
        write-host "No matches on qry:$($Title)" ; 
    } ; #if-E matchedhits ; 

#    } ; 
} #*------^ END Function Get-IMDBSearch ^------; 

#*======v SUB MAIN v======

$sQueryMinInterval=30 ; # seconds min between queries (avoid google bot throttling)
foreach ($movie in $Title){    
    $sw = [Diagnostics.Stopwatch]::StartNew() ; 
    $bRet = Get-IMDBSearch -Title $movie ; 
    # output formatting:
    write-host "=== v $($bRet.Title) v ===" ; 
    if($Full){
        ($bRet | out-string).trim()  ; 
    } else { 
        "Type:$($bRet.Type)" ; 
        "MPAARating:$($bRet.MPAARating)" ; 
        "Genres:$($bRet.Genres)" ; 
        "ImdbRating:$($bRet.UsrRatingsScore) for $($bRet.UsrRatingsCount)" ; 
        "RuntimeMinutes:$($bRet.RuntimeMinutes)" ; 
        "Country:$($bRet.Country)" ; 
        "Language:$($bRet.Language)" ; 
        "Color:$($bRet.Color)" ; 
        if($bRet.Released -eq "-" ){
            "Title:$($bRet.Title) ($($bRet.Released))" ; 
        } else { 
            "Title:$($bRet.Title) ($(get-date $bRet.Released -format 'yyyy'))" ; 
        }
        "Director:$($bRet.Director)" ; 
    } ; 
    write-host "=== ^ $($bRet.Title) ^ ===`n" ; 

    $sw.Stop() ; 
    write-host -foregroundcolor green "Elapsed Time: (HH:MM:SS.ms)" $sw.Elapsed.ToString() ;
    if($Title -is [system.array] | out-null ){
        if($sw.Elapsed.Seconds -lt $sQueryMinInterval){
            $shortSecs=($sQueryMinInterval - $sw.Elapsed.Seconds)
            write-host "(waiting $($shortSecs) more secs for next query)"
            start-sleep -Seconds $shortSecs ; 
        } ; 
    } ; 
    
} ; # parsing them into info

# SCRIPT-CLOSE MATERIAL TO CLEAR THE UNDERLYING $DBGPREF & $EAPREF TO DEFAULTS:
if ($ShowDebug -OR ($DebugPreference = "Continue")) {
        Write-Verbose -Verbose:$true "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
        $bDebug=$false
        # 8:41 AM 10/13/2015 also need to enable write-debug output (and turn this off at end of script, it's a global, normally SilentlyContinue)
        $DebugPreference = "SilentlyContinue" ;
} # if-E ; 
if($ErrorActionPreference -eq 'Stop') {$ErrorActionPreference = 'Continue' ; write-debug "(Restoring `$ErrorActionPreference:$ErrorActionPreference;"};

#*======^ END SUB MAIN ^======
