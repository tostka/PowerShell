# pull-gitBOD.ps1

  <# 
  .SYNOPSIS
  pull-gitBOD.ps1 - Beginning of day pull of all changes in each GitHub/Bitbucket remote repository down to local repositories.
  .NOTES
  Written By: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka
  Change Log
  # 12:05 PM 5/26/2016 added push/pull code for GH/BB
  # 11:55 AM 5/26/2016 based on commit-gitEOD.ps1
  # 2:46 PM 5/25/2016 - initial version
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  .\pull-gitBOD.ps1 ; 
  .LINK
#>

$ScRoot="c:\sc" ; 
cd $ScRoot ;
$reps=get-childitem -path $ScRoot -Directory |?{$_.BaseName -ne 'tmp'} ; 
foreach($rep in $reps){
    cd "$($rep.FullName)" ; 
    write-host -foregroundcolor yellow "$((get-date).ToString("HH:mm:ss")):=$(pwd) Status:" ; 
    git status -s ; 
    write-host -foregroundcolor green "Pulling changes from Remote Origin Repo...`n$(git remote -v):" ; 
    # git push -u origin master ;
    git pull origin master ;
    write-host -foregroundcolor yellow "===" ;
} ; 
cd $ScRoot ;
