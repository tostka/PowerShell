# restart-s4b.ps1

#*------v Function restart-s4b v------
Function restart-s4b {
    <#
    .SYNOPSIS
    restart-s4b - Close all lync processes, and reslaunch any *.ahk.lnk files in the "$env:APPDATA\...Startup" folder
    .NOTES
    Written By: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka
    Change Log
    * 8:29 AM 10/21/2016 minor format cleanup, genericizing, tests functional @work
    * 8:20 AM 10/21/2016 tests functional for s4b from work
    * 7:53 AM 10/21/2016 genericized, swapped varis in for hard-coded rgx's and values, ren'd $tRunRoot=>$aRunRoot, added $aRunLnkName rgx, ren $iUnc=>$oUnc, ren $tUnc=>$aUnc
    * 5:06 PM 8/9/2016 tshot the rev at home, now works on home host
    * 7:46 AM 8/9/2016 add -showdebug, force launch/clear objects into single obj strongly typed as array, conv $PMs/$FFs etc to single $oRun array obj
    * 5:06 PM 8/8/2016 stuck a test-path & dawdle for logonscript path completion prior to firing the target
    * 1:06 PM 8/7/2016 subbed out missing get-datetimestamp for pstf code
    # 7:21 AM 6/12/2015 port in support for the -NoRestart switch
    # 7:11 AM 5/18/2015 port back the restart-* enhancements
    # 7:26 AM 4/30/2015 port to profile funct and rename [verb]-restart.ps1 to restart-s4b()
    # 8:47 AM 1/15/2015
    # 7:46 AM 12/12/2014
    .DESCRIPTION
    restart-s4b - Close all lync processes, and relaunch any targeted .lnk files in the "$env:APPDATA\...Startup" or Quick Launch folder
    .PARAMETER NoRestart
    Parameter to suppress re-open[-NoRestart switch]
    .PARAMETER showDebug
    Switch parameter that indicates to display Debugging output[-showdebug]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    .\restart-s4b
    .LINK
    #>
    Param([Parameter(HelpMessage='NoRestart [$switch]')]
    [switch] $NoRestart
    ,[switch]$showDebug=$false  )  ;
    if ($showDebug) {
        write-debug "`$showDebug is $true. `nSetting `$DebugPreference = 'Continue'" ;
        $DebugPreference = "Continue" ;
        $bDebug=$true ;
    } else {
        $DebugPreference = "SilentlyContinue";
    };
    # constants & system-generated variables.
    $ScriptName=$myInvocation.ScriptName ;
    $StartFldr="$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"; #gci $StartFldr;
    $QckFldr="$env:userprofile\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch" ; 
    $rgxDrvLtr="^([a-zA-Z]{1}\:).*$" ; 
    $TargAppName="lync" ; # generic app descriptive name for echos
    $TargAppProc="^lync$" ; # specific app process (get-process) using a rgx, has -match below
    $aRunLnkName="^Skype\sfor\sBusiness\s\d{4}.lnk$" ; # launch shortcut name regex (uses -match)
    # below are dependent drive maps to test, set $null/blank if tests not needed
    $oRunRoot=$null ; # "I:\output" ; 
    $aUnc="\\MyComputer\apps" ; # app exe dependent unc map
    $oUnc="\\MyComputer\p2p" ; # output dependent unc map
    $RunBoxes="MyComputer".split(";") ; # semi-colon delimd machines from which this should be run
    
    if($oRunRoot){
        # back out the drive letter (fancier than it has to be, .substring(0,2) would work fine, but this pre-validates format of the string)
        if($oRunRoot -match $rgxDrvLtr){$oDrvLtr= $matches[1] }
        else {throw "invalid `$oRunRoot:$($oRunRoot)"}; 
    } ; 
    if($aRunRoot){
        $aRunRoot=$null ; # "T:\appexedrive" 
        if($aRunRoot -match $rgxDrvLtr){$aDrvLtr= $matches[1] }
        else {throw "invalid `$aRunRoot:$($aRunRoot)"}; 
    } ;       
    # 7:06 AM 8/9/2016 pre strongly type all processing objects into array (to permit += use), and one vari for kills, one for runs, pre-purge any prior use
    [array]$prcs = @() ;
    # 7:07 AM 8/9/2016 replace $PMs & $FFs with single array-type launch variable
    [array]$oRuns = @() ;
    $procs = $null ; $oRuns = $null ;
    if ($NoRestart) {
          write-verbose -verbose:$true "$((get-date -format "HH:mm tt");): -NoRestart specified." ;
      } ;
    write-verbose -verbose:$true "$((get-date -format "HH:mm tt");): --PASS STARTED:$ScriptName --" ;
    write-verbose -verbose:$true "killing $($TargAppName)" ;
    $prcs=get-process -ea silentlycontinue| ?{$_.name -match $TargAppProc} ;
    $prcs | select Name ;
    $prcs | stop-process -verbose ;
    start-sleep 2 ;
    $prcs=get-process -ea silentlycontinue| ?{$_.name -like $TargAppProc}  ;
    if ($prc) {write-verbose -verbose:$true "ZOMBIES:" ;$prc} ;
    if ($NoRestart) {
        # drop
    } else {
          write-verbose -verbose:$true "Launching Startup $TargAppName .lnks..." ;
          #$oRuns += get-childitem $StartFldr | ?{($_.Extension -like '*.lnk') -AND ($_.Name -like $aRunLnkName)} ;
          
          if($RunBoxes -contains ($env:COMPUTERNAME)) {
              # relaunch it
              # only run the 2ndary 'public' on work box
              # lnk in Start
              #$oRuns += get-childitem $StartFldr | ?{($_.Extension -like '*.lnk') -AND ($_.Name -like $aRunLnkName)} ;
              # lnk in QLaunch
              $oRuns += get-childitem $QckFldr | ?{($_.Extension -like '*.lnk') -AND ($_.Name -match $aRunLnkName)}
          } else {
              write-error "$((get-date).ToString("HH:mm:ss")):This app can only be run from $($RunBoxes -join " ")!";
          } ;  # if-E MyComputer
          if($bDebug){Write-Debug "`$oRuns:"; $oRuns | out-string } ;
          if($oRuns){
              # 7:35 PM 8/8/2016 running from sched task, i: doesn't seem present, test network drives & map as needed
              # 8:17 AM 10/21/2016 make situational
              if($oRunRoot){
                  if(test-path -path $oRunRoot -ea 0){"$($aDrvLtr) present"} 
                  else {"mapping missing $($aDrvLtr) to $oUnc" ; (New-Object -ComObject WScript.Network).MapNetworkDrive($($aDrvLtr),$($oUnc),$true);} ; 
              } ; 
              if($aRunRoot){
                  if(test-path -path $aRunRoot -ea 0){"$($aDrvLtr) present"} 
                  else {"mapping missing $($aDrvLtr) to $aUnc" ; (New-Object -ComObject WScript.Network).MapNetworkDrive($($aDrvLtr),$($aUnc),$true);} ; 
              } ; 
              # 8:17 AM 10/21/2016 make situational
              if($oRunRoot -OR $aRunRoot){
                  # 5:03 PM 8/8/2016 dawdle until both net drives mounted - logonscript at home can take awhile to run/complete.
                  write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):checking for/waiting-for mount of target: $($aRunRoot)`nand `$oRunRoot $($oRunRoot)" ; 
                  Do {write-host "." -NoNewLine;Start-Sleep -m (1000 * 5)} Until ( (test-path -path $aRunRoot -ea 0) -AND (test-path -path $oRunRoot -ea 0) ) ;
              } ; 
              $oRuns | %{$_.Name ; invoke-item $_.FullName }
          } else {
              write-warning "$((get-date).ToString("HH:mm:ss")):(no launchable configured)" ;
          } ; 
    } # if-E No-NoRestart
    
    if ($ShowDebug -OR ($DebugPreference = "Continue")) {
        Write-Verbose -Verbose:$true "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
        $bDebug=$false
        # 8:41 AM 10/13/2015 also need to enable write-debug output (and turn this off at end of script, it's a global, normally SilentlyContinue)
        $DebugPreference = "SilentlyContinue" ;
      } # if-E ;
    write-verbose -verbose:$true "$((get-date -format "HH:mm tt");): --PASS COMPLETED --"
} ;  # if-func block
# for profile use: config an alias for it
new-Alias rut restart-s4b
# for freestanding restart-s4b.ps1 use, just call the function:
#restart-s4b
#*------^ END Function restart-s4b ^------

restart-s4b


