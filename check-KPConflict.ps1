# check-KPConflict.ps1
# kp-conflict-chk.ps1
# debug syntax: Clear-Host ; . C:\usr\work\ps\scripts\check-KPConflict.ps1 -showDebug ;
<#
.SYNOPSIS
check-KPConflict.ps1 - Check keepass kdbx files for sync-service-based conflicts (gdrv & dbox)
.NOTES
Written By: Todd Kadrie
Website:    http://tinstoys.blogspot.com
Twitter:    http://twitter.com/tostka
Change Log
# 7:57 AM 3/14/2016: added hanging prompt on the window
# 7:26 AM 3/14/2016: gmail/whatever replaced quite a few dashes(-) with û. Subbed back out. 
# 11:40 AM 3/12/2016: did some recoding on the home v work dirs, and standardized to a root path
# 3:16 PM 3/11/2016 Ln440: clear error if you didn't use the traytip
# 7:40 AM 3/10/2016: Ren kp-conflict-chk.ps1|cmd => check-KPConflict.ps1|cmd, move ulb=>uwps. Switch from cmd to wrapperless (doc'd under Config: below).
# 7:15 AM 3/10/2016: Confirmed that both conflict variants rgx still works: google sample: 1st gen: tin[Conflict].kdbx ; 2nd generated: C:\usr\home\gdrv\db\tin[Conflict 1].kdbx.
# 12:11 PM 9/17/2015 gdrv conflicst aren't just '[Conflict #]' - 1st one has no #: tin[Conflict].kdbx, regx needs to accommodate: ^(?i:(.*\[Conflict(\s\d+)*\].*\.(KDB|KDBX)))$
# 11:54 AM 9/17/2015 loop both dbox & gdrv.
# * 9:28 AM 8/24/2015 - roughed in code to accommodate variant gdrv dir path @home from @work
# rev: 8:16 AM 8/18/2015 - shift to gdrv, by putting a flag in to indicate if in dbox or gdrv
# rev: 9:23 AM 4/10/2014
===========
Config:
3:08 PM 3/11/2016 flipped args from -command {...} (using -noexit to demo fire)
to -file: -noprofile -NoLogo -executionpolicy bypass -noexit -file C:\usr\work\ps\scripts\check-KPConflict.ps1
(Keepass Triggers are backed up to codedb...)
KeePass: Tools\Triggers, Triggers dialog:
  dbl-click Sync to Dbox trigger
    Actions: dblclick Execute Command line / URL:
      File/URL:
      origainal was the -ps1.cmd wrapper file.
      7:25 AM 3/10/2016: new Wrapperless cmdline:
      File/Url: %WINDIR%\system32\windowspowershell\v1.0\powershell.exe
      Arguments:  -noprofile -NoLogo -executionpolicy bypass -file C:\usr\work\ps\scripts\check-KPConflict.ps1
      [x]Wait for exit
      For testing: add -noexit to the args above, will cause the ps win to remain open for review after exit.
===========
.DESCRIPTION
check-KPConflict.ps1 - Check keepass kdbx files for sync-service-based conflicts (gdrv & dbox)
# check ...
# C:\usr\home\gdrv\db
# 9:19 AM 8/24/2015 home has a diff name, something like C:\usr\home\GoogleDrv\db
# need to recode below to accommodate by hostname
# C:\usr\home\Dropbox\db
# for kdbx or kdb files with instr ' conflicted copy' (dbox) or '[Conflict #]'(gdrv)
.PARAMETER ShowDebug
Parameter to display Debugging messages [-ShowDebug switch]
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.EXAMPLE
.LINK
*----------^ END Comment-based Help  ^---------- #>
<# #*======v TASK WRAPPER v====== (3:12 PM 3/11/2016 dropped went to cmd-line wrapperless launch above) #>
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Param(
    [Parameter(HelpMessage='Debugging Flag [$switch]')]
    [switch] $showDebug
) # PARAM BLOCK END
# pick up the bDebug from the $ShowDebug switch parameter
if ($ShowDebug) {$bDebug=$true; $DebugPreference = "Continue" ; write-debug "$((get-date).ToString('HH:mm:ss')):`$bDebug:$bDebug"; };
# constants
$ShowTime=30 ; # TrayTip ShowTime (seconds)

#11:15 AM 3/12/2016 redefine
$HmGdrvRoot="C:\path-to\gdrv\" ;
$WkGdrvRoot="C:\path-to\gdrv\" ;
$HmDbxRoot="C:\path-to\Dropbox\" ;
$WkDbxRoot="C:\path-to\Dropbox\" ;
$HmIcoDir="c:\path-to\grfx\icons\" ;
$WkIcoDir="c:\path-to\grfx\icons\" ;

if($($env:computername) -eq "computer1"){
     $GDrvDir=$WkGdrvRoot ;
     $DbxDir=$HmDbxRoot ;
     $IcoDir=$HmIcoDir ;
} elseif($($env:computername) -eq "computer2"){
    $GDrvDir=$HmDbxRoot ;
    $DbxDir=$WkDbxRoot ;
    $IcoDir=$WkIcoDir ;
} elseif($($env:computername) -eq "computer3"){
    $GDrvDir=$HmDbxRoot ;
    $DbxDir=$WkDbxRoot ;
    $IcoDir=$WkIcoDir ;
} ;
# 11:43 AM 3/12/2016 validate the dirs are in place
if(!(test-path (join-path -path $GDrvDir -ChildPath "\db")) ){ write-host "`g" ; write-error "MISSING $($GDrvDir). ABORTING!`nPress any key to EXIT. . ." ; $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp") | out-null ;} ;
if(!(test-path (join-path -path $DbxDir -ChildPath "\db"))){ write-host "`g" ; write-error "MISSING $($GDrvDir). ABORTING!`nPress any key to EXIT. . ." ; $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp") | out-null ;} ;
if(!(test-path (join-path -path $DbxDir -ChildPath "\db")) ){ write-host "`g" ; write-error "MISSING $($GDrvDir). ABORTING!`nPress any key to EXIT. . ." ; $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown,IncludeKeyUp") | out-null ;} ;
#"Dbox" ;
#"Gdrv"
$rgxConflictDbox= "^(?i:(.*conflicted\scopy.*\.(KDB|KDBX)))$"
#$rgxConflictGdrv= "^(?i:(.*\[Conflict\s\d+\].*\.(KDB|KDBX)))$"
# updated 12:12 PM 9/17/2015 to accommodate both tin[Conflict].kdbx & tin[Conflict #].kdbx' - 1st one has no #: tin[Conflict].kdbx
$rgxConflictGdrv= "^(?i:(.*\[Conflict(\s\d+)*\].*\.(KDB|KDBX)))$"

$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName)) ;

#*======v FUNCTIONS v======

#*----------------v Function show-TrayTip v----------------
function show-TrayTip {
    <#
    .SYNOPSIS
    Show-TrayTip() - Display popup System Tray Tooltip
    .NOTES
    Written By: Todd Kadrie
    Website:    http://tinstoys.blogspot.com
    Twitter:    http://twitter.com/tostka
    Change Log
    * 2:23 PM 3/10/2016 reworked the $TrayIcon validation, to permit either a valid path to an .ico, or a variable of Icon type (of the type pulled from shell32.dll by Extract-Icon())
    * 1:31 PM 3/10/2016 debugged and functional in check-kpconflict.ps1
    * 11:53 AM 3/10/2016 - added some concepts from the src Pat used: Dr. Tobias Weltner, http://www.powertheshell.com/balloontip/ (Pat was in the comments asking questions on the subject)
    * 9:39 AM 3/10/2016 - added some concepts from Pat Richard (pat@innervation.com),http://www.ehloworld.com/1038
    * 11:19 AM 3/6/2016 - unknown original, updating with formatting, pshelp and updated params
    c:\windows\system32\shell32.dll icon indexes:
    Warning
    .DESCRIPTION
    Show-TrayTip() - Display popup System Tray Tooltip
    .PARAMETER Type
    Tip Icon type [Error|Info|Warning|None]
    .PARAMETER Text
    Tip Text to be displayed [string]
    .PARAMETER title
    Tip Title [string]
    .PARAMETER ShowTime
    Tip Display Time (secs, default:2)[int]
    .PARAMETER TrayIcon
    Specify variant Systray icon (defaults per type)
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    show-TrayTip -type "error" -text "$computer is still ONLINE; Check that reboot is initiated properly" -title "Computer is not rebooting"
    Show TrayTip with default (powershell) Systray Icon, Error-type balloon icon, and balloon title & text specified, for 30 seconds
    .EXAMPLE
    show-TrayTip -type "error" -title "CONFLICT!" -text "CONFLICTED KEEPASS DB FOUND!" -ShowTime 30 -TrayIcon $TrayIcon ;
    Show TrayTip with custom Systray Icon, Error-type balloon icon, and balloon title & text specified, for 30 seconds
    .EXAMPLE
    show-TrayTip -type info -text "PowerShell script has finished processing" -title "Completed"
    Basic Example using parameter names (rest defaults)
    show-TrayTip info "PowerShell script has finished processing" "Completed"
    Basic Example using positional parameters
    .LINK
    *---^ END Comment-based Help  ^--- #>
    # 9:21 AM 3/10/2016 added None Icon option. switched ShowTime range from 10-30, to 1-30
    # 2:11 PM 3/10/2016 pulled $trayicon validatescript: [ValidateScript({Test-Path $_ | ?{-not $_.PSIsContainer}})]
    # 2:15 PM 3/10/2016: pulled $trayicon [string] type , to accommodate either a path or an object(icon)
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param(
        [Parameter(Position=0,Mandatory=$True,HelpMessage="Tip Icon type [Error|Info|Warning|None]")][ValidateSet("Error","Info","Warning","None")]
        [string]$Type
        ,[Parameter(Position=1,Mandatory=$True,HelpMessage="Tip Text to be displayed [string]")][ValidateNotNullOrEmpty()]
        [string]$Text
        ,[Parameter(Position=2,Mandatory=$True,HelpMessage="Tip Title [string]")]
        [string]$Title
        ,[Parameter(HelpMessage="Tip Display Time (secs, default:2)[int]")][ValidateRange(1,30)]
        [int]$ShowTime=2
        ,[parameter(HelpMessage = "Specify variant Systray icon (defaults to Powershell)")]
        $TrayIcon
    )  ;
    
    # 2:24 PM 3/10/2016 shift out of validation params, and manually check between trayicon =file path, and =icon pointer
    # Icon type variables have the following type name: $trayicon.gettype().name : Icon
    #>
    
    if($TrayIcon){
        if( (test-path $TrayIcon) -OR ($TrayIcon.gettype().name -eq 'Icon') ){ }
        else {
            write-warning "Invalid TrayIcon, resetting to Default Icon" ;
            $TrayIcon =$null ;
        }
    } ;
    if(!($NoTray)){
        #load Windows Forms and drawing assemblies
        [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null ; # used for TrayTip tips
        [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null ; # used for icon extraction
        
        #define an icon image pulled from PowerShell.exe
        #$icon=[system.drawing.icon]::ExtractAssociatedIcon((join-path $pshome powershell.exe)) ;
        
        # load the TrayTip
        # KEY POINT; don't create a new icon, if one exists, reuse it!
        if ($script:TrayTip -eq $null) {  $script:TrayTip = New-Object System.Windows.Forms.NotifyIcon } ;
        <# TrayIcon (BalloonTip): configurable property's:
          # the systray icon to be displayed (extracted from the PS path here)
          $path                    = Get-Process -id $pid | Select-Object -ExpandProperty Path ;
          $TrayTip.Icon            = [System.Drawing.Icon]::ExtractAssociatedIcon($path) ;
          # the following configure settings _within_ the balloon popup
          $TrayTip.BalloonTipIcon  = $Icon ;
          $TrayTip.BalloonTipText  = $Text ;
          $TrayTip.BalloonTipTitle = $Title ;
          # finally show the BalloonTip, with a specified timeout.
          $TrayTip.Visible         = $true ;
          $TrayTip.ShowBalloonTip($Timeout) ;
        #>
        if ($TrayIcon) { $TrayTip.Icon = $TrayIcon  }
        else {
            # use the extracted Powershell process icon
            $Path = Get-Process -id $pid | Select-Object -ExpandProperty Path ;
            $TrayTip.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) ;
        }
        $TrayTip.BalloonTipIcon  = $Type ;
        $TrayTip.BalloonTipText  = $Text ;
        $TrayTip.BalloonTipTitle = $Title ;
        # show it
        $TrayTip.Visible         = $true ;
        # set timeout - convert the spec'd ShowTime seconds to ms
        $TrayTip.ShowBalloonTip($ShowTime*1000) ;
        # 3/7/2016 - added debug echo
        if($bdebug){
            Write-Debug "$((get-date).ToString('HH:mm:ss')):ToolTip:`$ShowTime:$ShowTime`n`$text:$text`n`$type:$type`n`$TrayIcon:$TrayIcon`n`$Type:$Type`n`$Text:$Text`n`$Title:$Title`n`$ShowTime:$ShowTime"
            #Write-Debug "$((get-date).ToString('HH:mm:ss')):ToolTip:`$ShowTime:$ShowTime`n`$text:$text`n`$type:$type`n`$TrayIcon:$TrayIcon`n$Type:$Type`n$Text:$Text`n$Title:$Title`n$ShowTime:$ShowTime"
        };
        <# Cleanup code that should be used at script-end to cleanup the objects
            if($script:TrayTip) { $script:TrayTip.Dispose() ; Remove-Variable -Scope script -Name TrayTip ; }
        #>
    }  # if-E $NoTray ;
} #*----------------^ END show-TrayTip Function  ^---------------- ;

#*------v Function extract-Icon v------
Function extract-Icon {
    <#
    .SYNOPSIS
       Exports an ico from a given source to a given destination (file, if OutputIconFilename specified, to pipeline, if not)
    .Description
    Exports an ico from a given source to a given destination
    .PARAMETER SourceFilePath
    Source Exe/DLL to extract Icon from
    .PARAMETER IconIndex
    Optional Icon Index Number (only used for DLL's)
    .PARAMETER OutputIconFileName
    Optional Output path for extracted .ico file (if blank, returns the extracted Icon object)[c:\path-to\test.ico]
    .PARAMETER ExportIconResolution
    Optional Icon Output ExportIconResolution [264|128|48|32|16]
    .NOTES
    Versions:
    * 2:34 PM 3/10/2016 at this point, it's substantially expanded, it's got DLL extract code from Chrissy LeMaire, and stock [System.Drawing.Icon]::ExtractAssociatedIcon EXE-extract code.
        It also will either hand back an Icon-type variable, or if you spec an OutputIconfFileName, it writes the extracted icon file out and returns the $OutputIconFileName to confirm success.
    * 10:03 AM 3/10/2016: TSK: retooled
    * 1.1 2012.03.8 posted version
    .EXAMPLE
    extract-Icon -SourceFilePath  -IconIndex  ExportIconResolution
    .EXAMPLE
    $TrayIcon=extract-Icon -SourceFilePath (join-path -path $($env:WINDIR) "System32\shell32.dll") -IconIndex 238 -ExportIconResolution 16 ;  
    Grab the shell32.dll's 238'th icon into a variable that can be reassigned to a Traytip icon:
    #>
    Param (
        [parameter(Mandatory = $true,HelpMessage="Source Exe/DLL to extract Icon from")][ValidateScript({Test-Path $_ | ?{-not $_.PSIsContainer}})]
        [string]$SourceFilePath
        ,[parameter(Mandatory = $false,HelpMessage="Optional Output path for extracted .ico file (if blank, returns the extracted Icon object)[c:\path-to\test.ico]")]
        [string]$OutputIconFileName
        ,[parameter(HelpMessage="Optional Icon Index Number (only used for DLL's)")]
        [int]$IconIndex
        ,[parameter(HelpMessage="Optional Icon Output ExportIconResolution [264|128|48|32|16]")][ValidateSet(256,128,48,32,16)]
        [int]$ExportIconResolution
    ) ;
    $error.clear() ;
    # code that provides the DLL-extracting functions
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
namespace System
{
    public class IconExtractor
    {
     public static Icon Extract(string file, int number, bool largeIcon)
     {
      IntPtr large;
      IntPtr small;
      ExtractIconEx(file, number, out large, out small, 1);
      try
      {
       return Icon.FromHandle(largeIcon ? large : small);
      }
      catch
      {
       return null;
      }
     }
     [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
     private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
    }
}
"@
    TRY {
        #Grab the shell32.dll's 238'th iconinto a pointer that can be reassigned to a Traytip:
        # $TrayIcon=extract-Icon -SourceFilePath (join-path -path $($env:WINDIR) "System32\shell32.dll") -IconIndex 238 -ExportIconResolution 16 ;  
        # extract-Icon -SourceFilePath (join-path -path $($env:WINDIR) "System32\shell32.dll") -IconIndex 238  ;
        # dll's contain arrays of icons, need to pick one
         
        if(test-path -path $SourceFilePath){
            If ( $SourceFilePath.ToLower().Contains(".dll") ) {
                If(!($IconIndex)){
                    $IconIndex = Read-Host "Missing IconIndex param: Enter the icon index: " ;
                } ;
                # load the DLL extract code from the here string
                Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing ;
                $Icon = [System.IconExtractor]::Extract($SourceFilePath, $IconIndex, $true)  ;
            } Else {
                #
                [void][Reflection.Assembly]::LoadWithPartialName("System.Drawing") ;
                [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") ;
                #
                $Image = [System.Drawing.Icon]::ExtractAssociatedIcon("$($SourceFilePath)").ToBitmap() ;
                # image needs to be converted to bitmap and then into icon
                $Bitmap = new-object System.Drawing.Bitmap $image ;
                # defaulting res, not sure if it's a limit, or just defaulted for the guy's orig purp.
                $Bitmap.SetResolution($ExportIconResolution ,$ExportIconResolution ) ;
                $Icon = [System.Drawing.Icon]::FromHandle($Bitmap.GetHicon()) ;
            } ;
            if($OutputIconFileName){
                 write-verbose -verbose:$true  "Exporting Source File Icon..." ;
                $stream = [System.IO.File]::OpenWrite("$($OutputIconFileName)") ;
                $Icon.save($stream) ;
                $stream.close() ;
                write-verbose -verbose:$true  "Icon file can be found at $OutputIconFileName" ;
                # return 0 non-error status
                #0 | Write-Output ;
                # return path to exported file
                $OutputIconFileName | write-output ;
            } else {
                # extract & reuse command
                # or return the actual icon object
                $Icon | write-output ;
            }  # if-E
        } else {
          write-error "$((get-date).ToString('HH:mm:ss')):Non-existent `$SourceFilePath:$SourceFilePath. Aborting!";
        } # if-E ;
    } CATCH {
        Write-Error "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)"
        #Write-Error "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
        Exit #Opts: STOP(debug)|EXIT(close)|Continue(move on in loop cycle)
    } # try/catch-E ;
}#*------^ END Function extract-Icon ^------

#*======^ END FUNCTIONS ^======
#*======v SUB MAIN v======
<#    $GDrvDir
    $DbxDir
    $IcoDir
#>

<# 7:58 AM 3/14/2016 need a distinctive window, to find the damn thing
Administrator: Windows PowerShell
ahk_class ConsoleWindowClass
ahk_exe powershell.exe
#>
# set distinctive *FINDABLE* Window Title
$host.ui.RawUI.WindowTitle = "Running: $($ScriptBaseName)"

$SyncSvcs="Gdrv","Dbox" ;
foreach($SyncSvc in $SyncSvcs){
    if ($bDebug) {write-Debug "$((get-date).ToString('HH:mm:ss')):`$SyncSvc:$SyncSvc" } ;
    switch ($SyncSvc){
        "Dbox" {
            $KpSyncDir=(join-path -path $DbxDir -ChildPath "\db") ;
            write-host "Using Dbox directory:$($KpSyncDir)" ;
            $rgxConflict = $rgxConflictDbox
        }
        "Gdrv" {
            $KpSyncDir=$(join-path -path $GDrvDir -childpath "\db") ;
            write-host "Using Gdrv directory:$($KpSyncDir)" ;
            $rgxConflict = $rgxConflictGdrv ;
        }
        default {
            write-error"$((get-date).tostring('HH:mm:ss')):UNRECOGNIZED `$Syncsvc:$SyncSvc. ABORTING!";
          exit ;
        }
    }  # swtch-E ;
 
    if ($bDebug) {
      write-Debug  "$((get-date).ToString('HH:mm:ss')):`$KpSyncDir:$KpSyncDir"  ;
      write-Debug  "$((get-date).ToString('HH:mm:ss')):`$rgxConflict:$rgxConflict"  ;
    } ;
    
    
    <# $rgxConflictDbox= "^(?i:(.*conflicted\scopy.*\.(KDB|KDBX)))$"
    $rgxConflictGdrv= "^(?i:(.*\[Conflict\s\d+\].*\.(KDB|KDBX)))$"
    #>
    $Conflict=(gci $KpSyncDir | ?{$_.fullname -match $rgxConflict})
    if ($bDebug) {
      write-Debug  "$((get-date).ToString('HH:mm:ss')):`n`$Conflict:"
      $Conflict| format-list | out-default ;
    };
    if($Conflict) {
      # bing
      write-host "`a";
      # exclam
      
      # 12:02 PM 3/10/2016 updated version
      $TrayIcon=(join-path -path $IcoDir -childpath "explorer-redx.ico") ;
      if(!(test-path -path $TrayIcon)){write-verbose -verbose:$true  "$((get-date).ToString("HH:mm:ss")):Invalid `$TrayIcon:$TrayIcon. Defaulting to Powershell.exe icon" ;$TrayIcon=$null ;  } ;
      
      # 1:44 PM 3/10/2016 lets try the extract to pull an icon out of shell32.dll
      #$TrayIcon=extract-Icon -SourceFilePath (join-path -path $($env:WINDIR) "System32\shell32.dll") -IconIndex 238 -ExportIconResolution 16 ;
      #show-TrayTip -type error -title "CONFLICT!" -text "::$($SyncSvc)::CONFLICTED KEEPASS DB FOUND!`n$Conflict`nManually Sync the file to avoid DATALOSS!!!" -ShowTime $ShowTime -TrayIcon $TrayIcon ;
      
      # 1:26 PM 3/10/2016: shift to a @splat for params into function...
      $traysplat=@{
            type="error" ;
            title="CONFLICT!" ;
            text="::$($SyncSvc)::CONFLICTED KEEPASS DB FOUND!`n$Conflict`nManually Sync the file to avoid DATALOSS!!!" ;
            ShowTime=$ShowTime ;
            TrayIcon=$TrayIcon ;
      } ;
      
      # try the splat
      show-TrayTip @traysplat ;
      
      # without custom icon
      #New-BalloonTip -icon error -title "CONFLICT!" -text "CONFLICTED KEEPASS DB FOUND!`n$Conflict`nManually Sync the file to avoid DATALOSS!!!" -showtime 30
      if ($bDebug) {
          write-Debug  "$((get-date).ToString('HH:mm:ss'))::$($SyncSvc)::Conflict Detected!"
      } ;

      write-host -foregroundcolor red "`n`n$((get-date).ToString('HH:mm:ss'))::*** $($SyncSvc)::Conflict Detected! ***`n";          
      $bRet=Read-Host "Enter YYY to continue. Anything else will exit" 
      if ($bRet.ToUpper() -eq "YYY") {
           Write-host "Moving on"
      }  ; 
    } else {
      if ($bDebug) {
          write-Debug  "$((get-date).ToString('HH:mm:ss')):::$($SyncSvc)::No Conflict Detected"
      } ;

    } # if-block end
} ;  # loop-E
sleep 5 ;
# Cleanup code that should be used at script-end to cleanup the objects
# 3:16 PM 3/11/2016 clear error if you didn't use the traytip
if($script:TrayTip) { $script:TrayTip.Dispose() ; Remove-Variable -Scope script -Name TrayTip ; }
if ($bDebug) {
  write-Debug "$((get-date).ToString('HH:mm:ss')):`$SyncSvc:$SyncSvc"
  # set back to SilentlyContinue on exit)
  $DebugPreference = "SilentlyContinue" ;
} ;
#*======^ END SUB MAIN ^======
