# Remove-Comments.ps1
# dbg cmd: Clear-Host ; .\remove-Comments.ps1 -path C:\usr\work\exch\scripts\DAGNAn0-Stat-loop.ps1 -whatif -showdebug ;

<# 
.SYNOPSIS
Remove-Comments.ps1 - Strips Comments from specified -path file, and outputs updated copy of origin script as XXXXX-STRIPPED-yymmdd-hhmm.ps1
.NOTES
Written By: Todd Kadrie, leveraging Matthew Graeber (@mattifestation)'s Remove-Comments function
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Change Log
* 9:09 AM 1/22/2016 fixed bug in '[IO.File]::ReadAllText()' use in Remove-Comments(): it throws up on unicode chars; specifically any #requires line crashes it
            Replaced it with new Psv3 get-content -raw -path XXX command
* 8:54 AM 1/22/2016 - completed initial build
* 8:32 AM 1/22/2016 - initial pass

# 1-liner to rename the STRIPPED-yyyymmdd-hhmm back out of the file
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
get-childitem *.ps1 | ?{$_.name -like '*-STRIPPED-*'} | %{$_; $_ -match '(.*)(.*-STRIPPED-\d{8}-\d{4})(.*)' ; "$($matches[1])$($matches[3])" ; rename-item -path $_ -newname "$($matches[1])$($matches[3])" -whatif } ; 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

.DESCRIPTION
Remove-Comments.ps1 - Strips Comments from specified -path file, and outputs updated copy of origin script as 'XXXXX-STRIPPED-yymmdd-hhmm.ps1'. 
Does retain semicolon line ends. 
This strips all comments, along with all empty lines, and left justifies the code (unindenting). 
The tokenizer approach used targets just the excutable content in the script, dropping all other text.

.PARAMETER  Path
Path to PS1 to be stripped of comments[c:\path-to\script.ps1]
.PARAMETER  showDebug
'Debugging Flag [$switch]
.PARAMETER  Whatif
Whatif Flag  [$switch] (console-only no file output)
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Outputs a revised version of the input script as xxxx
.EXAMPLE
.\remove-Comments.ps1 -path C:\usr\work\exch\scripts\DAGNAn0-Stat-loop.ps1;
Remove comments from specified file
.EXAMPLE
.\remove-Comments.ps1 -path C:\usr\work\exch\scripts\DAGNAn0-Stat-loop.ps1 -whatif;
Whatif pass on Removing comments from specified file
.EXAMPLE
.\remove-Comments.ps1 -path C:\usr\work\exch\scripts\DAGNAn0-Stat-loop.ps1 -whatif -showdebug ;
Whatif pass on Removing comments from specified file, with Debugging output displayed.
.LINK
*---^ END Comment-based Help  ^--- #>


Param(
    [Parameter(HelpMessage='Path to PS1 to be stripped of comments[c:\path-to\script.ps1]')]    
    [ValidateScript({Test-Path $_})]
    [string]$Path,
    [Parameter(HelpMessage='Debugging Flag [$switch]')]
    [switch] $showDebug,
    [Parameter(HelpMessage='Whatif Flag  [$switch]')]
    [switch] $whatIf
) # PARAM BLOCK END

# pick up the bDebug from the $ShowDebug switch parameter
if ($ShowDebug) {$bDebug=$true};
if ($whatIf) {$bWhatIf=$true};

#*------v Function Remove-Comments v------
function Remove-Comments {

    <#
    .SYNOPSIS
    Remove-Comments() - Strips comments and extra whitespace from a script.
    PowerSploit Function: Remove-Comments
    Author: Matthew Graeber (@mattifestation)
    License: BSD 3-Clause
    Required Dependencies: None
    Optional Dependencies: None
	Web URL: https://github.com/PowerShellMafia/PowerSploit/blob/master/ScriptModification/Remove-Comments.ps1
	
    Change Log
    * 9:09 AM 1/22/2016 fixed bug in '[IO.File]::ReadAllText()' use, it throws up on unicode chars; specifically any #requires line crashes it
            Replaced it with new Psv3 get-content -raw -path XXX command

    .DESCRIPTION
    Remove-Comments strips out comments and unnecessary whitespace from a script. This is best used in conjunction with Out-EncodedCommand when the size of the script to be encoded might be too big.
    A major portion of this code was taken from the Lee Holmes' Show-ColorizedContent script. You rock, Lee!
    .PARAMETER ScriptBlock
    Specifies a scriptblock containing your script.
    .PARAMETER Path
    Specifies the path to your script.
    .EXAMPLE
    C:\PS> $Stripped = Remove-Comments -Path .\ScriptWithComments.ps1
    .EXAMPLE
    C:\PS> Remove-Comments -ScriptBlock {
    ### This is my awesome script. My documentation is beyond reproach!
    Write-Host 'Hello, World!' ### Write 'Hello, World' to the host
    ### End script awesomeness
    }
    Write-Host 'Hello, World!'
    .EXAMPLE
    C:\PS> Remove-Comments -Path Inject-Shellcode.ps1 | Out-EncodedCommand
    Description
    -----------
    Removes extraneous whitespace and comments from Inject-Shellcode (which is notoriously large) and pipes the output to Out-EncodedCommand.
    .INPUTS
    System.String, System.Management.Automation.ScriptBlock
    Accepts either a string containing the path to a script or a scriptblock.
    .OUTPUTS
    System.Management.Automation.ScriptBlock
    Remove-Comments returns a scriptblock. Call the ToString method to convert a scriptblock to a string, if desired.
    .LINK
    http://www.exploit-monday.com
    http://www.leeholmes.com/blog/2007/11/07/syntax-highlighting-in-powershell/
    #>
    
[CmdletBinding( DefaultParameterSetName = 'FilePath' )] Param (
	[Parameter(Position = 0, Mandatory = $True, ParameterSetName = 'FilePath' )]
	[ValidateNotNullOrEmpty()]
	[String] $Path,
	[Parameter(Position = 0, ValueFromPipeline = $True, Mandatory = $True, ParameterSetName = 'ScriptBlock' )]
	[ValidateNotNullOrEmpty()]
	[ScriptBlock] $ScriptBlock
) # if-PARAM

Set-StrictMode -Version 2
if ($PSBoundParameters['Path']) {
	Get-ChildItem $Path -ErrorAction Stop | Out-Null
	#$ScriptBlockString = [IO.File]::ReadAllText((Resolve-Path $Path))
    <# above throws errors on requires blocks
    Exception calling "Create" with "1" argument(s): "At line:44 char:11
+ #Requires �Modules ActiveDirectory
    #>
    # use of ReadAllText returns one string for whole file
    # get-content returns array of lines
    # but in Psv3 you can approximate it with: Get-Content "FileName.txt" -Raw
    $ScriptBlockString = Get-Content -Raw -path $(Resolve-Path $Path)  ; 
    # should probably sub-out all of them before the READALLText
	$ScriptBlock = [ScriptBlock]::Create($ScriptBlockString)
} else {
	# Convert the scriptblock to a string so that it can be referenced with array notation
	$ScriptBlockString = $ScriptBlock.ToString()
} # if-E

# Tokenize the scriptblock and return all tokens except for comments
$Tokens = [System.Management.Automation.PSParser]::Tokenize($ScriptBlock, [Ref] $Null) | Where-Object { $_.Type -ne 'Comment' } ; 
$StringBuilder = New-Object Text.StringBuilder
# The majority of the remaining code comes from Lee Holmes' Show-ColorizedContent script.
$CurrentColumn = 1
$NewlineCount = 0

foreach($CurrentToken in $Tokens) {
		# Now output the token
		if(($CurrentToken.Type -eq 'NewLine') -or ($CurrentToken.Type -eq 'LineContinuation')) {
			$CurrentColumn = 1
			# Only insert a single newline. Sequential newlines are ignored in order to save space.
			if ($NewlineCount -eq 0) {
				$StringBuilder.AppendLine() | Out-Null
			} # if-E
			$NewlineCount++
		} else {
			$NewlineCount = 0
			# Do any indenting
			if($CurrentColumn -lt $CurrentToken.StartColumn) {
					# Insert a single space in between tokens on the same line. Extraneous whiltespace is ignored.
					if ($CurrentColumn -ne 1) {
						$StringBuilder.Append(' ') | Out-Null
					} # if-E
			} # if-E
			# See where the token ends
			$CurrentTokenEnd = $CurrentToken.Start + $CurrentToken.Length - 1
			# Handle the line numbering for multi-line strings
			if(($CurrentToken.Type -eq 'String') -and ($CurrentToken.EndLine -gt $CurrentToken.StartLine)) {
				$LineCounter = $CurrentToken.StartLine ; 
				$StringLines = $(-join $ScriptBlockString[$CurrentToken.Start..$CurrentTokenEnd] -split '`r`n') ; 
				foreach($StringLine in $StringLines) {
					$StringBuilder.Append($StringLine) | Out-Null ; 
					$LineCounter++ ; 
				} # loop-E ; 
			} else {
				# Write out a regular token
				$StringBuilder.Append((-join $ScriptBlockString[$CurrentToken.Start..$CurrentTokenEnd])) | Out-Null
			}
			# Update our position in the column
			$CurrentColumn = $CurrentToken.EndColumn
		} # if-E
} # loop-E
Write-Output ([ScriptBlock]::Create($StringBuilder.ToString()))
} #*------^ END Function Remove-Comments ^------ ; 

#*======v SUB MAIN  v======
if(test-path $Path){
    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):`$Path specified: $($Path)...";
	$ifile = $path;
	$ifile = get-childitem $ifile ;
	if($ifile.Extension -eq '.ps1'){
		$ofile= join-path -path ($ifile.Directory) -childpath ( $ifile.BaseName.replace(".","-") + "-STRIPPED-" + (get-date -uformat "%Y%m%d-%H%M") + $ifile.Extension ) ;
		"`$ifile:$ifile";
		"`$ofile:$ofile";
		$Stripped = Remove-Comments -Path $Path ; 
		write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Lines Processed $((gc $Path | Measure -Line).count )";
		write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Outputting Stripped file to:`n$($ofile)";
        if($bWhatIf){
            write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):`n-whatif pass, dumping stripped output to console:";
            write-host -foregroundcolor yellow "# *------v $($path) v------ "; 
            $Stripped | out-string
            write-host -foregroundcolor yellow "# *------^ $($path) ^------ "; 
        } else {
            if($bDebug) { 
                write-host -foregroundcolor yellow "# *------v $($path) v------ "; 
                $Stripped | out-string
                write-host -foregroundcolor yellow "# *------^ $($path) ^------ "; 
            }  ; 
		    $Stripped | out-file -FilePath $ofile ; 
        }  # if-E
	} else {
		write-error "$((get-date).ToString("HH:mm:ss")):($path) specified is not a Powershell .ps1 file. EXITING";
	} # if-E
} else {
	write-error "$((get-date).ToString("HH:mm:ss")):Invalid/missing -path ($path) specified. EXITING";
}

#*======^ END SUB MAIN ^======
