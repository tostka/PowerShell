# create-MediaNFO.ps1
# debg cmd: clear-host ; C:\usr\work\ps\scripts\create-MediaNFO-20160110-0800PM.ps1 -IFile "I:\videos\Movie\comedy\Movie.mp4"
# debg cmd: Clear-Host ; .\create-MediaNFO.ps1 -ifile "I:\videos\Movie\all-flick-paths.csv" -showDebug -Whatif ; 

<# create-NFOFile.ps1
  .SYNOPSIS
  create-MediaNFO.ps1 - Pull url & xml summary of range of movies. Input via pipeline
  
  .NOTES
  Written By: ? Updated by Todd Kadrie
  Website:    https://ketarin.org/forum/topic/658-want-to-rename-your-movies-with-resolution-title-year-ketarin-can-do/
  Change Log
    * 7:28 PM 1/10/2016 - rewrote and added support for single-path spec in $ifile, in place of csv file
    * 6:00 PM 12/28/2015 appears functional
    * 10:34 AM 12/26/2015 rewrite from movlu & movlu2
  .DESCRIPTION
  .PARAMETER  IFile
  CSV File of local Movie paths, or a single path to a movie-containing directory.
  .PARAMETER showProgress
  ShowProgress [switch]
  .PARAMETER showDebug
  Debugging Flag [switch]
  .PARAMETER showDebug
  Whatif Flag  [switch]
  .PARAMETER whatIf
  Whatif Flag  [switch]
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  .LINK
  *---^ END Comment-based Help  ^--- #>


#Requires -version 3.0
Param(
    [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="CSV File of local Movie paths, or a single path to a movie-containing directory.")][ValidateNotNullOrEmpty()]
    [string]$IFile="I:\videos\Movie\all-flick-paths.csv",
    [Parameter(HelpMessage='ShowProgress [switch]')]
    [switch] $showProgress,
    [Parameter(HelpMessage='Debugging Flag [switch]')]
    [switch] $showDebug,
    [Parameter(HelpMessage='Whatif Flag  [switch]')]
    [switch] $whatIf
) # PARAM BLOCK END

# pick up the bDebug from the $ShowDebug switch parameter
if ($ShowDebug) {$bDebug=$true ; $DebugPreference = "Continue" ; };
if ($whatIf) {$bWhatIf=$true};

write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):=== PASS STARTED ===";

if(!(test-path $iFile)){ 
    write-error "$((get-date).ToString("HH:mm:ss")):INVALID iFile parameter, aborting. ";
    Exit
} else {
    # write-debug vers (requires :$DebugPreference = "Continue", and must be set back to SilentlyContinue on exit)
    if($bDebug){
        Write-Debug "`$iFile:$($iFile)"; 
    } ;
    
}# if-block end

 
#*----------------v Function Resolve-ImdbId() v----------------
function Resolve-ImdbId {
<#
    .Synopsis
        Converts an integer, string of numbers, or an object with the property "imdbID" to an IMDb (Internet Movie Database) formatted identifier in the form of "tt#######".
        If the input cannot be resolved, it is returned as is.

    .Parameter Id
        An integer, string of numbers, or an object with the property "imdbID".
        No characters are interpreted as wildcards.

    .NOTES
        Author: Benjamin Lemmond
        Email : benlemmond@codeglue.org

    .EXAMPLE
        Resolve-ImdbId 1234

        This example accepts an integer for the Id and returns a string value of "tt0001234".

    .EXAMPLE
        12, '345', '6879', 'BadId' | Resolve-ImdbId

        This example accepts four piped inputs, an integer and three strings, and returns the following string values:
        tt0000012
        tt0000345
        tt0006879
        BadId

    .EXAMPLE
        Get-ImdbTitle 'The Office' | imdbid

        This example uses Get-ImdbTitle to retrieve a PSCustomObject which has a property named 'imdbID'.
        This object is piped to imdbid (Resolve-ImdbId) and returns a string value of "tt0386676".
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [object[]]
        $Id
    )

    process {
      $Id | foreach {
          try {
              if ($_) {
                  if ($_ -is [int] -or $_ -match '^\s*\d+\s*$') {
                      return 'tt{0:0000000}'-f [int]$_
                  } # if-block end

                  if ($_.psobject.Properties['imdbID']) {
                      return $_.imdbID
                  } # if-block end
              } # if-block end
              $_
          } # try block end
          catch {
              Write-Error -ErrorRecord $_
          } # catch block end
      } # for-loop end
    } # process block end
}#*----------------^ END Function Resolve-ImdbId ^----------------

#*----------------v Function Get-ImdbTitle v----------------
function Get-ImdbTitle {
<#
    .Synopsis
        Retrieves IMDb (Internet Movie Database) information using the OMDb (Open Movie Database) API by Brian Fritz.
        Changes to the OMDb API may break the functionality of this command.

    .Parameter Title
        If no wildcards are present, the first matching title is returned.
        If wildcards are present, the first 10 matching titles are returned.

    .Parameter Year
        The year of the title to retrieve (optional).

    .Parameter Id
        An integer, string of numbers, or an object with the property "imdbID" that represents the IMDb ID of the title to retrieve.
        No characters are interpreted as wildcards.

    .NOTES
        Author: Benjamin Lemmond
        Email : benlemmond@codeglue.org

    .EXAMPLE
        Get-ImdbTitle 'True Grit'

        This example returns a PSCustomObject respresenting the 2010 movie "True Grit".

    .EXAMPLE
        Get-ImdbTitle 'True Grit' 1969

        This example returns a PSCustomObject respresenting the 1969 movie "True Grit".

    .EXAMPLE
        'True Grit' | Get-ImdbTitle -Year 1969

        Similar to the previous example except the title is piped.

    .EXAMPLE
        65126 | imdb

        This example also returns a PSCustomObject respresenting the 1969 movie "True Grit".
#>

    [CmdletBinding(DefaultParameterSetName='Title')]
    param (
        [Parameter(ParameterSetName='Title', Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string[]]
        $Title,

        [Parameter(ParameterSetName='Title', Position=1, ValueFromPipelineByPropertyName=$true)]
        [int]
        $Year,

        [Parameter(ParameterSetName='Id', Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [object[]]
        $Id
    )

    process {
      try {
          if ($PSBoundParameters.ContainsKey('Id')) {
              $queryStrings = $Id | Resolve-ImdbId | foreach { "i=$_" }
          } # for-loop end
          else {
              $yearParam = ''
              if ($Year) {
                  $yearParam = "&y=$Year"
              } # for-loop end

              $queryStrings = $Title | foreach {
                  $key = 't'
                  if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Title)) {
                      $key = 's'
                  } # for-loop end
                  "$key=$_$yearParam"
              } # for-loop end
          } # if-block end

          $uriRoot = 'http://www.omdbapi.com/?'
          $webClient = New-Object System.Net.WebClient

          $queryStrings | foreach {
              try {
                  $result = $webClient.DownloadString("$uriRoot$_") | ConvertFrom-Json

                  if ($result.psobject.Properties['Error']) {
                      throw [System.Management.Automation.ItemNotFoundException]$result.Error
                  } # if-block end

                  if (-not $result.psobject.Properties['Search']) {
                      return $result
                  } # if-block end

                  $result.Search | Resolve-ImdbId | foreach { $webClient.DownloadString("${uriRoot}i=$_") } | ConvertFrom-Json
              } # try block end
              catch {
                  Write-Error -ErrorRecord $_
              } # catch block end
          } # for-loop end
      } # try block end
      catch {
          Write-Error -ErrorRecord $_
      } # catch block end
    } # process block end
}#*----------------^ END Function Get-ImdbTitle ^----------------

#*----------------v Function Open-ImdbTitle v----------------
function Open-ImdbTitle {
<#
    .Synopsis
        Opens the IMDb (Internet Movie Database) web site to the specified title using the default web browser.
        Some features of this command are achieved via the OMDb (Open Movie Database) API by Brian Fritz.
        Changes to the OMDb API may break the functionality of this command.

    .Parameter Title
        If no wildcards are present, the first matching title is opened.
        If wildcards are present, the first 10 matching titles are opened.

    .Parameter Year
        The year of the title to open (optional).

    .Parameter Id
        An integer, string of numbers, or an object with the property "imdbID" that represents the IMDb ID of the title to open.
        No characters are interpreted as wildcards.

    .NOTES
        Author: Benjamin Lemmond
        Email : benlemmond@codeglue.org

    .EXAMPLE
        Open-ImdbTitle 'True Grit'

        This example opens the IMDb page for the 2010 movie "True Grit".

    .EXAMPLE
        Open-ImdbTitle 'True Grit' 1969

        This example opens the IMDb page for the 1969 movie "True Grit".

    .EXAMPLE
        'True Grit' | Open-ImdbTitle -Year 1969

        Similar to the previous example except the title is piped.

    .EXAMPLE
        65126 | imdb.com

        This example also opens the IMDb page for the 1969 movie "True Grit".
#>

    [CmdletBinding(DefaultParameterSetName='Title')]
    param (
        [Parameter(ParameterSetName='Title', Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string[]]
        $Title,

        [Parameter(ParameterSetName='Title', Position=1, ValueFromPipelineByPropertyName=$true)]
        [int]
        $Year,

        [Parameter(ParameterSetName='Id', Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [object[]]
        $Id
    )

    process {
      try {
        if ($PSBoundParameters.ContainsKey('Id')) {
            $imdbId = $Id | Resolve-ImdbId
        } # if-block end
        else {
            $imdbResult = Get-ImdbTitle @PSBoundParameters
            if (-not $imdbResult) { return $imdbResult }
            $imdbId = $imdbResult | foreach { $_.imdbID }
        } # if-block end

        $imdbId | foreach { Start-Process "http://imdb.com/title/$_" }
      } # try block end
      catch {
          Write-Error -ErrorRecord $_
      } # catch block end
    } # process block end
}#*----------------^ END Function Open-ImdbTitle ^----------------

<#
New-Alias imdbid   Resolve-ImdbId -Force
New-Alias imdb     Get-ImdbTitle  -Force
New-Alias imdb.com Open-ImdbTitle -Force


Export-ModuleMember -Function *-* -Alias *
#>
<#
Get-ImdbTitle 'True Grit' 1969
#>


#*----------------v Function PickList v----------------
function PickList(){

    <# 
    .SYNOPSIS
    PickList() - Console dynamic picklist 
    .NOTES
    Written By: Paul Westlake
    Website:    https://quickclix.wordpress.com/2014/10/30/stringarrayasmenu/
  
    Change Log
    #12:02 PM 12/26/2015 retooling for my own use, defaulting to passed in array, with csv detect & convert
    30th October 2014 - posted version
    .DESCRIPTION
    PickList() - Console dynamic picklist. Takes an array or CSV string and presents the user with a list of choices and returns their selection for later use.
    .PARAMETER  MnuOptions
    Array or semi-colon-delimited string of values to be displayed and picked from.
    .PARAMETER  MnuPrompt
    Prompt text to display with the menu
    .PARAMETER  MnuChoiceReq
    Forces selection before exit (no cancel)
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    $collection = "gmail.com;hotmail.com;yahoo.com;mail.com;outlook.com" ; 
    $MenuChoiceText = "Please select your Domain" ; 
    $selectedDomain = PickList $collection $MenuChoiceText $true ; 
    Write-Host "`n`t`tYou Selected $selectedDomain`n" -Fore Yellow ; 
    Pick a domain from a delimeted string list
    .EXAMPLE
    $collection = "Server1","Server2","Server3","Server4" ; 
    $MenuChoiceText = "Please select your Server"
    $selectedServer = PickList $collection $MenuChoiceText $false
    Write-Host "`n`t`tYou Selected $selectedServer`n" -Fore Yellow
    Select a server from the list, Selection is not required.
    .LINK
    https://quickclix.wordpress.com/2014/10/30/stringarrayasmenu/
    *---^ END Comment-based Help  ^--- #>

  
    param(
        [String]$MnuOptions,
        [String]$MnuPrompt,
        [alias("PickReq")]
        [Boolean]$MnuChoiceReq
    )
    $i=0
    $xPickListSelection=$null

    # set-up Empty array
    $xPickListArray = @()
    
    # if not array, and contains ;
    #if( ($MnuOptions -isnot [system.array]) -AND ($MnuOptions -is [string]) -AND ($MnuOptions.indexof(";")) ){
    if( (($MnuOptions | measure).count -eq 1) -AND ( $MnuOptions.indexof(";") -gt 0) ){
        # Convert CSV list to hash table
        $MnuOptions = $MnuOptions.split(";") ; 
    } # if-E

    # Convert hash table to multi dimensional array
    foreach($entry in $MnuOptions){
        # for each entry add to the array and increment the counter
        $array += (,($i,$entry)) ; 
        $i=$i+1 ; 
    } # loop-E

    # Write the Prompt Header
    Write-Host "`n`t`t$MnuPrompt`n" ; 

    # for each entry display on the console screen
    foreach ($arrayentry in $array){
        Write-Host $("`t`t`t"+$arrayentry[0]+".`t"+$arrayentry[1]) ; 
    }

    # Select prompt type depending if required selection or not
    if( $MnuChoiceReq ){
        # Loop while the user has not made a selection
        while (!$xPickListSelection){
            # advise that selection is required
            Write-Host "`n`t`tRequired" -Fore Red -NoNewline ; 
            # Prompt for and record pick list choice
            $xPickListSelection = Read-Host "`tEnter Option Number" ; 
        } # loop-E
        # Return the operators chosen option.
        return $array[$xPickListSelection][1] ; 
    }else{
        # advise that selection is not required
        Write-Host "`n`t`tNot Required" -Fore White -NoNewline ; 
        $xPickListSelection = Read-Host "`tEnter Option Number" ; 
        # If the user selected an option return it else $null will return
        if($xPickListSelection){
            # Return the operators chosen option.
            # this returns the picklist name
            #return $array[$xPickListSelection][1] ; 
            # instead we want to return the record number
            write-output $xPickListSelection
        } ; # if-E
    } ; # if-E 
}#*----------------^ END Function PickList ^----------------
<# demo
$xCSVList = "gmail.com;hotmail.com;yahoo.com;mail.com;outlook.com" ; 
$MenuChoiceText = "Please select your Domain" ; 
$selectedDomain = PickList $collection $MenuChoiceText $true ; 
Write-Host "`n`t`tYou Selected $selectedDomain`n" -Fore Yellow ; 
#>

#*------v Function Test-Transcribing v------
# Author: Oisin Grehan
# URL: http://poshcode.org/1500
# Tests for whether transcript (start-transcript) is already running
# usage:
#   if (Test-Transcribing) {stop-transcript} ;
function Test-Transcribing {
  $externalHost = $host.gettype().getproperty("ExternalHost",
        [reflection.bindingflags]"NonPublic,Instance").getvalue($host, @())

  try {
    $externalHost.gettype().getproperty("IsTranscribing",
        [reflection.bindingflags]"NonPublic,Instance").getvalue($externalHost, @())
  } catch {
     write-warning "This host does not support transcription."
  }
} #*------^ END Function Test-Transcribing ^------


#*----------------v Function Remove-Chars v----------------
function Remove-ForiegnChars {
    <# 
    .SYNOPSIS
    Remove-ForiegnChars() - Replace Foreign Language special characters into base ascii equivelents (from comment at http://www.powershellmagazine.com/2014/08/26/pstip-replacing-special-characters/)
    .NOTES
    Written By: Dirk
    Website:    http://www.powershellmagazine.com/2014/08/26/pstip-replacing-special-characters/
    Change Log
    * 7:43 PM 12/27/2015 cleaned up added help etc
    * August 27, 2014 at 4:27 am     posted vers
    .DESCRIPTION
    .PARAMETER  src
    Source string of text with characters to be normalized
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Remove-ForiegnChars 'Zaz=lc GESLA jazn'
    .LINK
    http://www.powershellmagazine.com/2014/08/26/pstip-replacing-special-characters/
    *---^ END Comment-based Help  ^--- #>

    param ([String]$src = [String]::Empty) ; 
    #replace diacritics
    $normalized = $src.Normalize( [Text.NormalizationForm]::FormD ) ; 
    $sb = new-object Text.StringBuilder ; 
    $normalized.ToCharArray() | % {
        if( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {[void]$sb.Append($_) } ; 
    } ; 
    $sb=$sb.ToString() ; 
    #replace via code page conversion
    $NonUnicodeEncoding = [System.Text.Encoding]::GetEncoding(850) ; 
    $UnicodeEncoding = [System.Text.Encoding]::Unicode ; 
    [Byte[]]$UnicodeBytes = $UnicodeEncoding.GetBytes($sb);
    [Byte[]]$NonUnicodeBytes = [System.Text.Encoding]::Convert($UnicodeEncoding, $NonUnicodeEncoding , $UnicodeBytes);
    [Char[]]$NonUnicodeChars = New-Object -TypeName "Char[]" -ArgumentList $($NonUnicodeEncoding.GetCharCount($NonUnicodeBytes, 0, $NonUnicodeBytes.Length)) ; 
    [void]$NonUnicodeEncoding.GetChars($NonUnicodeBytes, 0, $NonUnicodeBytes.Length, $NonUnicodeChars, 0);
    [String]$NonUnicodeString = New-Object String(,$NonUnicodeChars) ; 
    write-output $NonUnicodeString ; 
} ; #*----------------^ END Function Remove-ForiegnChars ^----------------


#*================v SUB MAIN v================

$nl = "`n" ; 
# Constants
# Following can be edited to include other video file types
Set-Variable videofiletypes -Option Constant -Value @("*.avi", "*.mpg", "*.flv", "*.wmv", "*.mp4", "*.mkv") ; 

# non-diacrit 
# 5:32 PM 12/28/2015 looks like apostrophe has to be dbld to be escaped
Set-Variable rgxForeignChars -Option Constant -Value '[^a-zA-Z0-9\s\$:\(\)\.,!'']' ; 
<#$string = 'Le pFre Nodl est une ordure' ; $string -match $rgxForeignChars ; $matches;$matches[0];
returns True, if diacriticals in string, and places first matched diacrit char into $matches[0]

Name                           Value
----                           -----
0                              F
#>

# gen filename from script and start/stop
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
$TimeStampNow = get-date -uformat "%Y%m%d-%H%M"
$outtransfile=$ScriptDir + "logs"
if (!(test-path $outtransfile)) {Write-Host "Creating dir: $outtransfile" ;mkdir $outtransfile ;} ;
$outtransfile+="\" + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.log" ;
#stop-transcript -ErrorAction SilentlyContinue
# stop transcript,trap any error & eat complaint
# note, this will suppress all errors coming out of the transcript commands - even one's you WANT to see:
#Trap {Continue} Stop-Transcript | Out-Null ;
# alt:     
if (Test-Transcribing) {stop-transcript} ;
start-transcript -path $outtransfile ;

# count infos generated
[int]$nfos = 0
$paths = $null ; 
$paths = @()  ; 

# test the inbound 'iFile' = get-childitem -path "$path\*" -include $videofiletypes | ?{!($_.PSIsContainer)}
# if it's a csv, process it, if it's not and is a vid file, use it solely
if ((get-childitem -path $iFile).Extension -eq ".csv"){
    # we're going to run a csv of problem movies paths on i:
    $paths = import-csv $iFile | select path ; 
} else {
    #$paths = (get-childitem -path $iFile.tostring()).fullname ; 
    # 8:39 PM 1/10/2016 no, we don't want a vid file, just it's directory (the csv is dirs, not file paths)
    $paths = (get-childitem -path $iFile.tostring()).Directory ; 
    # $states.Add("Alaska", "Fairbanks")
    #$paths.add("path",(get-childitem -path $iFile.tostring()).fullname);
} ; 
#foreach ($path in $paths.path){
foreach ($path in $paths){

    if($bDebug){
        Write-Debug "`$path:$($path)"; 
    } ;

    if($path -match ".*est\sune.*"){
       IF($bDebug){write-host "STINKER!"}
    } ; 
    # get vid file pull -recurse, should be in top dir
    #if($vidfile = get-childitem -path $path -include $videofiletypes | ?{!($_.PSIsContainer)}) { 
    # 5:01 PM 12/27/2015 since we're halting the recurse, we need to append \* to end of path and hunt the local dir
    # 8:37 PM 1/10/2016 we've either got a path or a directory name here, not a video file
    if($vidfile = get-childitem -path "$path\*" -include $videofiletypes | ?{!($_.PSIsContainer)}) { 
    
        $movs = $null ; 
        # 11:56 AM 12/27/2015 problem, some dirs have files named with extra identifiers: I:\videos\Movie\christmas\Christmas Carol, A (2009, Jim Carrey): 
        $movlu=(split-path $path -leaf).replace("}","")    

        <#$string = 'Le pFre Nodl est une ordure' ; $string -match $rgxForeignChars ; $matches;$matches[0];
        returns True, if diacriticals in string, and places first matched diacrit char into $matches[0]

        Name                           Value
        ----                           -----
        0                              F
        #>
        # check for and sub-out spec chars
        if($movlu -match $rgxForeignChars){
            if($bDebug){"first matched diacrtical char: $($matches[0])" } ;      
            $movlu=Remove-ForiegnChars -src $movlu ; 
        }
        # get-imdbtitle supports -title -year -id, split it
        if(($movlu.indexof("(")) -AND ($movlu.indexof(")"))){
            #$movlu=$movlu.split("(").replace(")","")
            #if($movlu -match "(.*)\s\((\d{4})\)"){
            # update the rgx to accommodate (2009, Jim Carrey)
            if($movlu -match "(.*)\s\((\d{4})((,)*).*\)"){
                # we should have $matches[1]=title, and $matches[2]=yyyy, $movlu[0]=original string
                #$matches
                if($bDebug){
                    $matches ; 
                } ;
                write-host -foregroundcolor green "===Looking up title/year:$($matches[1])/$($matches[2])...";
                # 12:09 PM 12/27/2015 need to massage "title, A" => "A title"
                $titlelu=$($matches[1]);
                [int]$titleyr=$($matches[2]);
                if($titlelu -match "(.*)\,\s((A|The))"){
                    $titlelu="{1} {0}" -f $matches[1],$matches[2]
                }    
            
                $movs = (Get-ImdbTitle -title "$($titlelu)" -year "$titleyr"  | sort year ) 
                # 12:27 PM 12/27/2015 retry on fail
                if(!$movs){
                    # retry no year wildcard
                    write-warning "$((get-date).ToString("HH:mm:ss")):Failed title/year search, retrying simplified wildcard" ;
                    $movs = (Get-ImdbTitle -title "$($titlelu)*"  | sort year ) 
                }
            } # if-E
        } else {
            # just run it all ; 
            write-host -foregroundcolor green "Looking up $movlu...";
            $movs=$null ; 
            $movs = (Get-ImdbTitle "$movlu"  | sort year )  ; 
        } ; 

        if(($movs | measure).count -gt 0){
            if($bDebug){
                Write-Debug "Matches Count:$(($movs|measure).count)"; 
            } ;
            # build a hash of title (year) combos for the picklist
            [array]$movlist=$null ; 
            foreach($mov in $movs) {
                $movlist+= ($mov.title + " [" + $mov.year +"]")
            } # loop-E
            # mult matches use PickList to prompt for a choice
            if(($movs | measure).count -gt 1){


                $MenuChoiceText = "Please select your movie"
                $MovN = PickList -MnuOptions $movlist -MnuPrompt $MenuChoiceText -MnuChoiceReq $false ; 
           
            } else {
                $MovN = 0 ; 
            } # if-E
            Write-Host -ForegroundColor Yellow "Selected item# $($MovN): $($movs[$MovN].Title) $($movs[$MovN].year)" ; 

            $movbuf=@{
                Title=$null;
                imdbRating=$null;
                year=$null;
                imdbVotes=$null;
                Plot=$null;
                Runtime=$null;
                Rated=$null;
                imdbID=$null;
                genre=$null;
                Director=$null;
                actors=$null;
                url=$null;
            } ; 
            $movbuf.Title=$($movs[$MovN].Title) ;
            $movbuf.imdbRating=$($movs[$MovN].imdbRating) ;
            $movbuf.year=$($movs[$MovN].year) ;
            $movbuf.imdbVotes=$($movs[$MovN].imdbVotes) ;
            $movbuf.Plot=$($movs[$MovN].Plot) ;
            $movbuf.Runtime=$($movs[$MovN].Runtime) ;
            $movbuf.Rated=$($movs[$MovN].Rated) ;
            $movbuf.imdbID=$($movs[$MovN].imdbID) ;
            $movbuf.genre=$($movs[$MovN].genre) ;
            $movbuf.Director=$($movs[$MovN].Director) ;
            $movbuf.actors=$($movs[$MovN].actors) ;
            $movbuf.url="http://www.imdb.com/title/$($movs[$MovN].imdbID)" ;
        
            if($bdebug){
                $movbuf | out-string ; 
            } ; 
            # build xml content here
            [string]$s  = "<?xml version=""1.0"" encoding=""utf-8""?>" + $nl
            $s += "  <movie>" + $nl
            $s += "    <title>$($movbuf.Title)</title>" + $nl
            $s += "    <originaltitle>$($movbuf.Title)</originaltitle>" + $nl
            $s += "    <sorttitle>$($movbuf.Title)</sorttitle>" + $nl
            $s += "    <set></set>" + $nl
            $s += "    <rating>$($movbuf.imdbRating)</rating>" + $nl
            $s += "    <year>$($movbuf.year)</year>" + $nl
            $s += "    <top250>0</top250>" + $nl
            $s += "    <votes>$($movbuf.imdbVotes)</votes>" + $nl
            $s += "    <outline>$($movbuf.Plot)</outline>" + $nl
            $s += "    <plot>$($movbuf.Plot)</plot>" + $nl
            $s += "    <tagline></tagline>" + $nl
            $s += "    <runtime>$($movbuf.Runtime)</runtime>" + $nl ;
            $s += "    <thumb></thumb>" + $nl ;
            $s += "    <mpaa>$($movbuf.Rated)</mpaa>" + $nl ;
            $s += "    <playcount>0</playcount>" + $nl ;
            $s += "    <id>$($movbuf.imdbID)</id>" + $nl ;
            $s += "    <filenameandpath></filenameandpath>" + $nl ; 
            $s += "    <trailer></trailer>" + $nl ; 
            # genre needs to be split
            $genres = $movbuf.genre.split(",").trim() ;  
            foreach ($genre in  $genres) {
                if($genre.length){
                    $s += "    <genre>$($genre)</genre>" + $nl ;             
                }
            }  ; # loop-E
            #$s += "    <genre>$($movs[$MovN].Genre)</genre>" + $nl ; 
            $s += "    <credits></credits>" + $nl ; 
            $s += "    <fileinfo>" + $nl ; 
            $s += "         <streamdetails></streamdetails>" + $nl ; 
            $s += "    </fileinfo>" + $nl ; 
            $s += "    <director>$($movbuf.Director)</director>" + $nl ; 
        
            # split actors and assign out
            $actors = $movbuf.actors.split(",").trim()
            foreach ($actor in $actors ) {
                <# 
                <actor>
                    <name>Paul Begala</name>
                    <role>Himself</role>
                </actor>
                #>
                if($actor.length){
                    $s += "    <actor>" + $nl ; 
                    $s += "         <name>$($actor)</name>" + $nl ; 
                    $s += "         <role></role>" + $nl ; 
                    $s += "    </actor>" + $nl ; 
                } # if-E
            }  ; # loop-E
            $s += "  <movie>" + $nl
            # add url: http://www.imdb.com/title/tt0103002/
            $s += "  $($movbuf.url)"

            if($bDebug){
                Write-Debug "--- v Xml output v ---";
                Write-Debug $s
                Write-Debug "--- ^ Xml output:^ ---";
            } ;

            [string]$outfile = (join-path -path $path -childpath ($vidfile.BaseName + ".nfo")) ; 
            If(!(test-path $outfile)){
                Write-Host "Creating NFO file: $($outfile)"
                if(!$whatif){
                    try {
                      Set-Content -Encoding UTF8 -Path $outfile -Force -Value $s
                      $nfos++
                    } catch {
                        Write-Host $("ERROR: Unable to create NFO file """ + $outfile + """. Do you have permissons and is there sufficient disk space?") -Foreground Red
                        Exit
                    } # if-E
                } else {
                    write-host -foregroundcolor green "Whatif:Skipping execute...";
                }
            } ; 

            <#
            [String]$MnuOptions,
            [String]$MnuPrompt,
            [alias("PickReq")]
            [Boolean]$MnuChoiceReq
            #>
            # then summarize and dump the movies returned:
            <#
            foreach ($mov in $movs) {
                write-host -fore yell ($mov.Title + " [" + $mov.year + "]" );
                $mov  | select director,Genre,Rated,Type,imdbid,@{Name='iRating';
                Expression={ ($_.imdbRating.ToString() + " [" + $_.ImdbVotes + "]")}},Actors,Runtime,Country ;
                "Mov FileName:`n$($mov.Title) $($mov.year)"
                # http://www.imdb.com/title/tt0103002
                "IMDBUrl:`nhttp://www.imdb.com/title/$($mov.imdbid)"
            };
            #>
        } else {
            write-warning "No match returned ($movlu[1] $movlu[2])" ;
        }
    
        $i++ ; 

    } else {
        write-warning "$((get-date).ToString("HH:mm:ss")):No video file found for`n $path`n--SKIPPING---" ;
        
    }  ; # if-E vidfile
    
} # for-loop end

if($bdebug){"`$i:$i"} ; 

write-host -foregroundcolor green "$($nfos) .nfo files created"
write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):=== PASS COMPLETED ===";
stop-transcript ;
if ($ShowDebug) {
    $bDebug=$true ; 
    $DebugPreference = "SilentlyContinue" ;
 };


#*================^ END SUB MAIN ^================
