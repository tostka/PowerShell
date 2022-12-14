# create-NFOFileNotInXBMC.ps1

  <# create-NFOFile.ps1
  .SYNOPSIS
  create-NFOFile.ps1 - Mossywell's NFO File Creator
  .NOTES
  Written By: mossywell
  Website:	http://forum.kodi.tv/showthread.php?tid=147530
  Change Log
	6:03 PM 12/24/2015 Todd Kadrie, rework to be command-line nfo roller, drop the sql/xbmc features
	20121216 Version: 0.2 - mossy's posted version 
	0.1 - First test release onto the XBMC forums.
	0.2 - Stricter type handling.
	 - Added Frodo support via enumeration (probably unnecessary frippery!)	
	# To Do:  - Remove the duplicated code - there's far too much of it! - DONE PARTIALLY
				- Regional strings.
				- Catches should be more specific.
				- Support for TV shows. (I know the NFO format, but not the database format because I don't have any TV shows!)
  .DESCRIPTION
  .PARAMETER  Path
  
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  .LINK
  *---^ END Comment-based Help  ^--- #>

<# ----------- disabled 6:11 PM 12/24/2015 pullling xbmc code
# Eunms (see http://connect.microsoft.com/PowerShell/feedback/details/742760/add-type-cannot-add-type-the-type-name-already-exists)
try {
Add-Type -TypeDefinition @"
   public enum XBMCVersion
   {
      Null,
      Eden,
      Frodo
   }
"@
}
catch {
  # Try already exists (see HTTP link above)
}
----------- disabled 6:11 PM 12/24/2015 pullling xbmc code 
#>

$ErrorActionPreference = "Stop"

<# ---------- disabled 6:11 PM 12/24/2015 pullling xbmc code
#*----------------v Function getSQLiteData v----------------
Function getSQLiteData($profile, $query) {
  # Return: System.Data.DataTable as System.Object[]
  
  [string]$sqlitedll = [Environment]::GetEnvironmentVariable("ProgramFiles") + "\System.Data.SQLite\2008\bin\System.Data.SQLite.dll"
  try {
    [void][System.Reflection.Assembly]::LoadFrom($sqlitedll)
  }
  catch {
    Write-Host $("ERROR: Unable to load the SQLite DLL """ + $sqlitedll + """. Is the ADO.NET adapter for SQLite installed?") -Foreground Red
    Exit
  }
  
  # Which DB file?
  switch($xbmcversion) {
    ([XBMCVersion]::Eden) {
      $dbfile = "MyVideos60.db"
    }
    ([XBMCVersion]::Frodo) {
      $dbfile = "MyVideos75.db"
    }    
  }

  # Create the connection object
  $conn = New-Object -TypeName System.Data.SQLite.SQLiteConnection

  # Create the connection string, query, data table, adapter
  if($profile -eq "Master user") {
    $connectionstring = "Data Source=" + $([Environment]::GetEnvironmentVariable("UserProfile")) + "\AppData\Roaming\XBMC\userdata\Database\$dbfile"
  } else {
    $connectionstring = "Data Source=" + $([Environment]::GetEnvironmentVariable("UserProfile")) + "\AppData\Roaming\XBMC\userdata\profiles\" + $profile + "\Database\$dbfile"
  }
  $conn.ConnectionString = $connectionstring
  [System.Data.DataTable]$table = New-Object System.Data.DataTable
  $adapter = New-Object System.Data.SQLite.SQLiteDataAdapter($query, $conn)

  # Let's do it - the table will contain the results
  try {
    $adapter.Fill($table) > $null
  } catch {
    Write-Host $("ERROR: The SQL query failed for connection string """ + $connectionstring + """. Does the file exist and is it accessible?") -Foreground Red
    Exit
  }
  $conn.Close()
  
  # Return the table
  $table
}#*----------------^ END Function getSQLiteData ^----------------

#*----------------v Function IsInMovieLibrary v----------------
Function IsInMovieLibrary($path) {
  # Default is not found
  [bool]$found = $false
  
  foreach($row in $moviepaths) {
    # Case insensitive
    if($path -eq $row[0]) {
      $found = $true
      break
    }
  }
  
  $found
}#*----------------^ END Function IsInMovieLibrary ^----------------
----------- disabled 6:11 PM 12/24/2015 pullling xbmc code
#>

#*----------------v Function generateNfos v----------------
Function generateNfos([string]$searchpath, [bool]$recurse) {
  <# 
  .SYNOPSIS
  generateNfos - Generate xml NFO file(s)
  .NOTES
  Written By: TMossywell's NFO File Creator
  .NOTES
  Written By: mossywell
  Website:	http://forum.kodi.tv/showthread.php?tid=147530
  Change Log
  10:17 PM 12/25/2015 - 
  .DESCRIPTION
  .PARAMETER  <Parameter-Name>
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  .LINK
  *---^ END Comment-based Help  ^--- #>
  # Constants
  # Following can be edited to include other video file types
  Set-Variable videofiletypes -Option Constant -Value @("*.avi", "*.mpg", "*.flv", "*.wmv", "*.mp4", "*.mkv")
  
  # Variables
  [int]$nfos = 0
  
  # For each directory...
  if($recurse) {
    $dirs = Get-ChildItem -Path $searchpath -Recurse | where {$_.PSIsContainer}
  } else {
    $dirs = Get-ChildItem -Path $searchpath | where {$_.PSIsContainer}
  }
  
  if($dirs -ne $null) {    
    foreach($dir in $dirs) {    
      # Only create an NFO if it's not in the library already (directory must have a "\" at the end)
      #if(-not (IsInMovieLibrary($dir.FullName + "\"))) {
	  # 10:31 PM 12/25/2015 switch to if no .nfo
	  if( !(test-path -path "$($dir.FullName)\*.nfo")){
        # Only do it if there's a VIDEO_TS subfolder - it's a DVD
		
		<#
        if(Test-Path -Path $($dir.FullName + "\VIDEO_TS")) {
          [string]$s  = "<?xml version=""1.0"" encoding=""utf-8""?>" + $nl
          $s += "  <movie>" + $nl
          $s += "    <title>$($dir.BaseName)</title>" + $nl
          $s += "    <id>-1</id>" + $nl
          $s += "  </movie>"
          [string]$outfile = $dir.FullName + "\movie.nfo"
          # Output to the host NOT the pipeline!
          Write-Host $("Creating NFO file: """ + $outfile + """")
          try {
            Set-Content -Encoding UTF8 -Path $outfile -Force -Value $s
            $nfos++
          }
          catch {
            Write-Host $("ERROR: Unable to create NFO file """ + $outfile + """. Do you have permissons and is there sufficient disk space?") -Foreground Red
            Exit
          } # if-block end
        } # # if dvd
		#>
		
      } # if-block end
    } # for-loop end
  } # if-block end

  # For each video file...
  if($recurse) {
    $files = Get-ChildItem -Path $($searchpath + "*") -Recurse -Include $videofiletypes | where {!$_.PSisContainer}
  } else {
    $files = Get-ChildItem -Path $($searchpath + "*") -Include $videofiletypes | where {!$_.PSisContainer}
  }
  
  if($files -ne $null) {    
    foreach($file in $files) {
      # Only create an NFO if it's not in the library already
      #if(-not (IsInMovieLibrary($file.FullName))) {    
	  # only create an nfo if there's no existing nfo
	  If(!(test-path $file.FullName\*.nfo)){
        [string]$s  = "<?xml version=""1.0"" encoding=""utf-8""?>" + $nl
        $s += "  <movie>" + $nl
        $s += "    <title>$($file.BaseName)</title>" + $nl
        $s += "    <id>-1</id>" + $nl
        $s += "  </movie>"
        [string]$outfile = $file.DirectoryName + "\" + $file.BaseName + ".nfo"
        # Output to the host NOT the pipeline!
        Write-Host $("Creating NFO file: """ + $outfile + """")
        try {
          Set-Content -Encoding UTF8 -Path $outfile -Force -Value $s
          $nfos++
        }
        catch {
          Write-Host $("ERROR: Unable to create NFO file """ + $outfile + """. Do you have permissons and is there sufficient disk space?") -Foreground Red
          Exit
        }
      }
    }
  }
  
  $nfos
}#*----------------^ END Function generateNfos ^----------------

#*----------------v Function getProfiles v----------------
Function getProfiles() {
  # Return: System.Object[]
  
  # Load the profiles.xml file
  [string]$profilefile = [Environment]::GetEnvironmentVariable("UserProfile") + "\AppData\Roaming\XBMC\userdata\profiles.xml"
  try {
    # Recast
    [xml]$profilefile = Get-Content $profilefile
  }
  catch {
    Write-Host $("ERROR: Unable to access """ + $profilefile + """.") -Foreground Red
    Exit
  }
  
  # Pass to the pipeline the profile names into an array
  @($profilefile.profiles.profile | % {$_.name})
}#*----------------^ END Function getProfiles ^----------------

#*----------------v Function getInstallLocation v----------------
Function getInstallLocation() {
  # Return: [string] Location of the XBMC loation from the registry
  [string]$location = ""
  try {
    $location = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\XBMC" | % { $_.InstallLocation }
  }
  catch {
    Write-Host "ERROR: Unable to get the XBMC installation location. Is it installed?" -Foreground Red
    Exit
  }
  
  # Return
  $location
}#*----------------^ END Function getInstallLocation ^----------------

#*----------------v Function getXBMCVersion v----------------
Function getXBMCVersion([string]$location) {
  # Return: [XBMCVersion] The Exe major version from the file properties
  [XBMCVersion]$xbmcver = [XBMCVersion]::Null
  
  [string]$exefile = $location + "\XBMC.exe"

  [string]$ver = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exefile).FileVersion
  $ver = ($ver.Split("."))[0]
  
  switch($ver) {
    "11" {
      $xbmcver = [XBMCVersion]::Eden
    }
    "12" {
      $xbmcver = [XBMCVersion]::Frodo
    }
  }
  
  # Return
  $xbmcver
}#*----------------^ END Function getXBMCVersion ^----------------

#*================v SUB MAIN v================
# Constants
Set-Variable nl -Option Constant -Value $([Environment]::NewLine)
# Variables
[int]$nfocount = 0
[string]$query = ""

# Get the install location
[XBMCVersion]$xbmcversion = (getXBMCVersion (getInstallLocation))
"XBMC Major Version: " + $xbmcversion.ToString()

# Which user are we interested in?
$profiles = getProfiles

# Ask the user which profile we are interested in
[int]$i = 1
$($nl + "The following XBMC profiles were found:")
foreach($profile in $profiles) {
  # $profile is [string]
  
  $i.ToString() + ": " + $profile
  $i++
}
Do {
	$selectedprofile = Read-Host "Enter the user you wish to create NFO files for and press Return"
} while((1..$($profiles.Count)) -notcontains $selectedprofile)

# Now grab the movie paths and search paths for that user
$query = "SELECT strPath,strContent,strHash,scanRecursive FROM path WHERE strContent = ""movies"" OR (strContent = ""tvshows"" AND strHash IS NOT NULL)"
$searchpaths = getSQLiteData $profiles[$selectedprofile - 1] $query
$query = "SELECT c22 FROM movie"
$moviepaths = getSQLiteData $profiles[$selectedprofile - 1] $query

# Iterate each search path
foreach($path in $searchpaths) {
  # $path is System.Data.DataRow
  
  # Check if the original path contained the "recurse" instruction
  if($path[3] -gt 0) {
    $retval = generateNfos ($path[0].ToString()) $true
  } else {
    $retval = generateNfos ($path[0].ToString()) $false
  }
  $nfocount += $retval
} # for-loop end

"NFO files created: " + $nfocount
