# c:\usr\home\db\tsk-incl-Toys.ps1 - Toy Function Includes

#*======v NON-ADMIN: C:\Users\Account\Documents\WindowsPowerShell\profile.ps1 v======
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
# NON-ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\Account\Documents\WindowsPowerShell\profile.ps1
# NON-ADMIN acct $profile.CurrentUserCurrentHost loc
#C:\Users\Account\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1

# ADMIN acct $profile.CurrentUserAllHosts loc
#C:\Users\Accounts\Documents\WindowsPowerShell\profile.ps1
# notepad2 $profile.CurrentUserAllHosts ;
#*======

#*------V Comment-based Help (leave blank line below) V------ 

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

write-host -foregroundcolor gray "$((get-date).ToString("HH:mm:ss")):====== EXECUTING: $(Split-Path -Leaf ((&{$myInvocation}).ScriptName) ) ====== " ; 
#           ==============
#*======v *** TOY FUNCTIONS *** v======

#*------v Function get-fortune v------
function get-fortune {
  # read fortune.txt in profile path
  # vers: 11:20 AM 5/14/2014 have it output a string and handle the string externally
  # vers: 8:41 AM 5/1/2014
  #[System.IO.File]::ReadAllText((Split-Path $profile)+'\fortune.txt') -replace "`r`n", "`n" -split "`n%`n" | Get-Random
  # or read fortune.txt from c:\usr\local\bin
  if (test-path c:\usr\local\bin\fortune.txt) {
    $utildrive = "c"
  } # if-block end
  # 8:34 AM 5/1/2014 shift into a variable
  # replace is replacing CrLf with Lf, and splitting quotes on the `n%`n pattern
  $fortune=[System.IO.File]::ReadAllText($utildrive + ':\usr\local\bin\fortune.txt') -replace "`r`n", "`n" -split "`n%`n" | Get-Random
  
  # 9:32 AM 5/14/2014 output a value
  return $fortune
} #*------^ END Function get-fortune ^------

#*------v Function get-Excuse v------
function get-excuse {
  # vers 9:36 AM 5/14/2014, switched to returning a string, do formatting externally
  # adapted online fortune/excuse source BOFH http://occasionalutility.blogspot.com/2013/10/everyday-powershell-part-2-bofh-speaks.html
  $path = "http://www.cs.wisc.edu/~ballard/bofh/excuses"
  $localfile = "c:\usr\local\bin\excuse.html"
  download-file $path $localfile
  $excuse = Get-Content $localfile | get-random
  # 9:32 AM 5/14/2014 output a value
  return $excuse
  
} #*------^ END Function get-Excuse ^------

#*======^ *** TOY FUNCTIONS *** ^======
#           ==============
