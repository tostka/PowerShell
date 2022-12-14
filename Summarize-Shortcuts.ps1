# Summarize-Shortcuts.ps1

# summarizing shortcuts into csv
#*------v Function Summarize-Shortcuts v------
function Summarize-Shortcuts {
    <# 
    .SYNOPSIS
    Summarize-Shortcuts - Summarize a lnk file or folder of lnk files (COM wscript.shell object).
    .NOTES
    Written By: Todd Kadrie (based on sample by MattG)
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka
    Change Log
    8:29 AM 4/28/2016 rewrote and expanded, based on MattG's orig concept
    2013-02-13 18:38:52 MattG posted version
    .DESCRIPTION
    .PARAMETER  <Parameter-Name>
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    $Output = Summarize-Shortcuts -path "C:\test" ;
    Summarize a folder of .lnk files. 
    .EXAMPLE
    $Output = Summarize-Shortcuts -path "C:\test\Pale Moon PubProf -safe-mode.lnk" ;
    Summarize a single .lnk file.
    .LINK
    https://powershell.org/wp/forums/topic/find-shortcut-targets/
    *---^ END Comment-based Help Summarize-Shortcuts ^--- #>

    Param(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage='Path [c:\path-to\]')]
        [ValidateScript({ 
          if(-not (Test-Path -LiteralPath $_)) { throw "Path '${_}' does not exist. Please provide the path to a file or folder on your local computer and try again." } ; 
          $true ; 
        })] 
        [string]$Path
        ,[Parameter(HelpMessage='Debugging Flag [$switch]')]
        [switch] $showDebug
        ,[Parameter(HelpMessage='Whatif Flag  [$switch]')]
        [switch] $whatIf
    ) # PARAM BLOCK END

    # pick up the bDebug from the $ShowDebug switch parameter
    if ($ShowDebug) {$bDebug=$true};
    if ($whatIf) {$bWhatIf=$true};

    
    if(test-path $Path -pathtype Container){        
            # get all #lnks
            $Lnks = @(Get-ChildItem -Recurse -path $Path -Include *.lnk) ; 
    } elseif(test-path $Path -pathtype Leaf){    
            # get single .lnk
            $Lnks = @(Get-ChildItem -path $Path -Include *.lnk) ; 
    } else {
          write-error "$((get-date).ToString("HH:mm:ss")):INVALID -PATH: NON-FILESYSTEM CONTAINER OR FILE OBJECT";
          break ; 
    } 
    
    $Shell = New-Object -ComObject WScript.Shell ; 
    <# exposed props/methods of the $shell:
        $shell.CreateShortcut($lnk) | select *
        FullName         : C:\Users\kadrits\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\Net\Firefox PrivProf.lnk
        Arguments        : -no-remote -P "ToddPriv"
        Description      :
        Hotkey           :
        IconLocation     : C:\usr\home\grfx\icons\tin_omen\guy_fawkes_mask.ico,0
        RelativePath     :
        TargetPath       : C:\Program Files (x86)\Mozilla Firefox\firefox.exe
        WindowStyle      : 1
        WorkingDirectory : C:\Program Files (x86)\Mozilla Firefox
    #>
    if($host.version.major -lt 3){
        $Props = @{
            ShortcutName = $null ; 
        } ;
    } else { 
        $Props = [ordered]@{
            ShortcutName = $null ; 
        } ;
    } ; 
    $Props.Add("TargetPath",$null) ;
    $Props.Add("Arguments",$null) ;
    $Props.Add("WorkingDirectory",$null) ;
    $Props.Add("IconLocation",$null) ;

    foreach ($Lnk in $Lnks) {
        $Props.ShortcutName = $Lnk.Name ; 
        $Props.TargetPath = $Shell.CreateShortcut($Lnk).targetpath ; 
        $Props.Arguments = $Shell.CreateShortcut($Lnk).Arguments ;
        $Props.WorkingDirectory= $Shell.CreateShortcut($Lnk).WorkingDirectory;
        $Props.IconLocation= $Shell.CreateShortcut($Lnk).IconLocation;
        # 7:41 AM 4/28/2016 added explicit write summary obj to pipeline
        New-Object PSObject -Property $Props | write-output ; ;
    }  # loop-E;
    # unload the Wscript.shell obj
    [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null ;
}#*------^ END Function Summarize-Shortcuts ^------

# this will break running sid (wrong path)
$Output = Summarize-Shortcuts -path "$($env:APPDATA)\Microsoft\Internet Explorer\Quick Launch\Net" ;
$Output = Summarize-Shortcuts -path "C:\Users\kadrits\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\Net" ;
$Output | out-string ;
$Output | select ShortcutName,TargetPath,Arguments | out-string ;
$Output = Summarize-Shortcuts -path "C:\Users\kadrits\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\Net\Pale Moon PubProf -safe-mode.lnk" ;
$Output | out-string ;
$Output | select ShortcutName,TargetPath,Arguments | out-string ;
#$Output | Export-Csv -LiteralPath .\Lnks_targets.csv
