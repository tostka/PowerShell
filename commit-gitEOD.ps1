# commit-gitEOD.ps1

  <# 
  .SYNOPSIS
  commit-gitEOD.ps1 - End of day forced commit of everything in each Repository in Git
  .NOTES
  Written By: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka
  Change Log
  # 12:05 PM 5/26/2016 added push/pull code for GH/BB
  # 2:46 PM 5/25/2016 - initial version
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  .\commit-gitEOD.ps1 ; 
  .LINK
#>

$ScRoot="c:\sc" ; 
cd $ScRoot ;
$reps=get-childitem -path $ScRoot -Directory |?{$_.BaseName -ne 'tmp'} ; 
foreach($rep in $reps){
    cd "$($rep.FullName)" ; 
    write-host -foregroundcolor yellow "$((get-date).ToString("HH:mm:ss")):=$(pwd) Status:" ;
    git status -s ; 
    $Memo="$(get-date -format 'yyyyMMdd-HHmmtt'):EOD closing commit" ; 
    write-host -foregroundcolor gray "ExcCmd: git add -A ; git commit -m $($Memo);" ; 
    write-host -foregroundcolor gray "=$(pwd) Adds:" ; 
    git add -A ;
    write-host -foregroundcolor gray "=$(pwd) Commit:$($Memo)" ; 
    git commit -m "$($Memo)" ;
    write-host -foregroundcolor green "Pushing changes to remote origin...`n$(git remote -v):" ; 
    git push -u origin master ;
    write-host -foregroundcolor yellow "===" ;
} ; 
cd $ScRoot ;