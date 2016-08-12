# c:\usr\home\db\tsk-incl-Home.ps1.ps1


#*======v NON-ADMIN: C:\Users\Account\Documents\WindowsPowerShell\profile.ps1 v======
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\Account\Documents\WindowsPowerShell\profile.ps1
# NON-ADMIN acct $profile.CurrentUserCurrentHost loc
#C:\Users\kadriets\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1

# ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\Accounts\Documents\WindowsPowerShell\profile.ps1
# notepad2 $profile.CurrentUserAllHosts ;
#*======

## 
#     .SYNOPSIS
# NON-ADMIN: C:\Users\Account\Documents\WindowsPowerShell\profile.ps1 - 
# My primary non-admin Profile file
# 
#     .NOTES
# Written By: Todd Kadrie
# Website:	http://tinstoys.blogspot.com
# Twitter:	http://twitter.com/tostka
# 
# Additional Credits: [REFERENCE]
# Website:	[URL]
# Twitter:	[URL]
# 
# Change Log.
#* 8:37 AM 10/30/2015 trimmed length of --/==
# 9:46 AM 12/23/2014 - split out Toys to include file
# 8:13 AM 12/23/2014 constructed SERVER-SIDE CORE subset of functions appropr for profile on server desktops

write-host -foregroundcolor gray "$((get-date).ToString("HH:mm:ss")):====== EXECUTING: $(Split-Path -Leaf ((&{$myInvocation}).ScriptName) ) ====== " ; 

#           =========
#*======v *** HOME FUNCTIONS *** v======
# moved to c:\usr\home\db\tsk-incl-Home.ps1.ps1

write-host "loading imdb functions"
#*------v Function Resolve-ImdbId() v------
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
    PARAM (
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
}#*------^ END Function Resolve-ImdbId ^------

#*------v Function Get-ImdbTitle v------
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
    PARAM (
        [Parameter(ParameterSetName='Title', Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string[]]$Title
        ,[Parameter(ParameterSetName='Title', Position=1, ValueFromPipelineByPropertyName=$true)]
        [int]$Year
        ,[Parameter(ParameterSetName='Id', Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [object[]]$Id
    ) ; 

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
}#*------^ END Function Get-ImdbTitle ^------

#*------v Function Open-ImdbTitle v------
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
    PARAM (
        [Parameter(ParameterSetName='Title', Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string[]]$Title
        ,[Parameter(ParameterSetName='Title', Position=1, ValueFromPipelineByPropertyName=$true)]
        [int]$Year
        ,[Parameter(ParameterSetName='Id', Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [object[]]$Id
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
}#*------^ END Function Open-ImdbTitle ^------
<#   
New-Alias imdbid   Resolve-ImdbId -Force
New-Alias imdb     Get-ImdbTitle  -Force
New-Alias imdb.com Open-ImdbTitle -Force
#>


#*======^ *** HOME FUNCTIONS *** ^======
#           =========

