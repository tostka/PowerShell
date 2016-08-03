#check-LyncSvcStatusRpt.ps1
# clean out the set-csuserpin overhead code

# *** REGION SETUP MARKER
#region SETUP ; 

#*----------V Comment-based Help (leave blank line below) V---------- 

<# 
.SYNOPSIS
check-LyncSvcStatusRpt.ps1 - mailed-report version of the svcstatus script
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Change Log
* 9:42 AM 12/14/2015 updated Send-EmailNotification, tested for & tweaked psv2 params, which Psv3 has, and v2 does not. Also updated the attachment code to properly  test for status. And finally switched @$email to base model, and tested to add extra attributes, rather than keeping multiple concurrent full versions
* 12:51 PM 3/20/2015 2nd vers

.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Outputs a text transcript and emails the report to $rptAdmin (emailadmin@domain.com)
.EXAMPLE
.\check-LyncSvcStatusRpt.ps1
.LINK
*----------^ END Comment-based Help  ^---------- #>

#region TASKWrapper ; 
<#*================v TASK WRAPPER v================
#*================^ END TASK WRAPPER ^================
#>
#endregion TASKWrapper ;


[CmdletBinding()]
Param();


# pick up the bDebug from the $ShowDebug switch parameter
if ($ShowDebug) {$bDebug=$true};
If ($Whatif){write-host "`$Whatif is $true" ; $bWhatif=$true}; 
#$bDebug = $false;
#$bDebug = $true;

#$ProgInterval= 500 ; # write-progress wait interval in ms

# 1:13 PM 2/4/2015: add a $TemplatePath spec
#$TemplateCustomFile="LyncPINTemplate.html";

# 12:15 PM 2/9/2015 add an SMTP retry limit (per user attempted)
[int]$WelcomeRetryLimit=4;
[int]$RetryDelay=20;    # wait time after failure
# 1:57 PM 2/18/2015
$AbortPassLimit = 4;    # maximum failed users to abort entire pass
# 9:49 AM 2/17/2015 SMTP Priority level[Normal|High|Low]
$Priority="Normal";
# SMTP port (default is 25)
$SMTPPort = 25 ; 

# 12:23 PM 2/20/2015 add gui vb prompt support
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null ; 
# 11:00 AM 3/19/2015 should use Windows.Forms where possible, more stable

# 2:10 PM 2/4/2015 shifted to here to accommodate include locations
$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName)) ;
$SendAddr = (($ScriptBaseName.replace(".","-")) + "@domain.com") ; 
$SmtpServer="smtp.domain.com";
$SMTPLab="LYNMS6200D.global.ad.domainlab.com";
$DataStore = "COBJ";
if($bdebug){Write-Host -ForegroundColor Yellow "DataStore:$DataStore";};


$ComputerName = $env:COMPUTERNAME ;



# 12:48 PM 3/11/2015 detect -noprofile runs (in case you need to add profile content/functions to get to function)
$NoProf=[bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); 
# if($NoProf){# do this};

$DomLab="domain-LAB";

#*================v HTML BOILERPLATE (USE IN SUB MAIN) v================
# browser title:
$PageTitle="Lync Server Status Report";
#above the table Legend
#$PreTitle = $PageTitle + "<br>Processing Results $(get-timestampnow)" ;
# 2:30 PM 2/13/2015 obso function
#$PreTitle = $PageTitle + "<br>Processing Results $(get-timestamp)" ;
# 2:30 PM 2/13/2015 no, revise the dep out, use raw call
$PreTitle = $PageTitle + "<br>Processing Results $((get-date).ToString("HH:mm:ss"))" ;
$Author = "Todd Kadrie" ;
write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ":LOADING INCLUDES:");

# move CSSInline spec up before the include load
$CSSInline=$true;

# stock non-graybar
#$sLoad=(join-path $LocalInclDir "tor-incl-html.ps1") ; 
# switch it to the $ScriptDir
$sLoad=(join-path $ScriptDir "tor-incl-html.ps1") ; 
# graybar
#$sLoad=(join-path $LocalInclDir "tor-incl-html-TOR-logo-graybar.ps1") ; 
$sLoad=(join-path $ScriptDir "tor-incl-html-TOR-logo-graybar.ps1") ;
#$sLoad
# 2:30 PM 3/20/2015 disable this POS!
#*================^ HTML BOILERPLATE (USE IN SUB MAIN) ^================

# 2:01 PM 2/5/2015
$sQot = [char]34
$sQotS = [char]39

if($bdebug){
  $ErrorActionPreference = 'Stop';
  # Write-Debug will write messages to the screen. You need to set, $DebugPreference = "continue", so they actually appear.
  # 2:13 PM 1/21/2015 looks like this is the source of the mass of DEBUG: messages coming out of AD!
  #$DebugPreference = "continue";
};

# Clear error variable
$Error.Clear()

#*================v FUNCTIONS  v================
#*----------------v Function Get-TimeStamp v----------------
function Get-TimeStamp {

    <# 
    .SYNOPSIS
    Get-TimeStamp - Return "HH:mm:ss" Timestamp
    .NOTES
    Written By: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka

    Change Log
    * 11:23 AM 3/19/2015 cleanup added pshelp
    * 2:06 PM 12/3/2014 fixed, lms says  non-existent func
    * 8:43 AM 10/31/2014 - simple timestamp echo

    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns "HH:mm:ss" Timestamp
    .EXAMPLE
    Write "$(Get-TimeStamp):MESSAGE"
    .EXAMPLE
    write-host "$(Get-TimeStamp):MESSAGE" –verbose ;
    .EXAMPLE
    write-host -foregroundcolor yellow  "$(Get-TimeStamp):MESSAGE";
    .LINK
    *----------^ END Comment-based Help  ^---------- #>

	Get-Date -Format "HH:mm:ss";	
} #*----------------^ END Function Get-TimeStamp ^----------------

#*----------------v Function Test-TranscriptionSupported v----------------
function Test-TranscriptionSupported {
  <#
  .SYNOPSIS
  Tests to see if the current host supports transcription.
  .DESCRIPTION
  Powershell.exe supports transcription, WinRM and ISE do not.
  .Example
  #inside powershell.exe
  Test-Transcription
  #returns true
  Description
  -----------
  Returns a $true if the host supports transcription; $false otherwise
  #>
  
  #($Host.Name -eq 'ServerRemoteHost')
  $externalHost = $host.gettype().getproperty("ExternalHost",
  [reflection.bindingflags]"NonPublic,Instance").getvalue($host, @())
  try {
    [Void]$externalHost.gettype().getproperty("IsTranscribing",
    [Reflection.BindingFlags]"NonPublic,Instance").getvalue($externalHost, @())
    $true
  } catch {
    $false
  } # try-E
}#*----------------^ END Function Test-TranscriptionSupported ^----------------

#*----------------v Function Test-Transcribing v----------------
function Test-Transcribing {
    <#.SYNOPSIS
    Tests for whether transcript (Start-Transcript) is already running
    .NOTES
    Author: Oisin Grehan
    URL: http://poshcode.org/1500
    requires -version 2.0

    Change Log
    10:13 AM 12/10/2014

    .INPUTS
    .OUTPUTS
    Outputs $TRUE/FALSE reflecting transcribe status
    .EXAMPLE
    if (Test-Transcribing) {Stop-Transcript} 
    #>

    

    # 8:56 AM 2/24/2015 move the test outside of the try
    if($Host.Name -ne "Windows PowerShell ISE Host"){
        $externalHost = $host.gettype().getproperty("ExternalHost",
           [reflection.bindingflags]"NonPublic,Instance").getvalue($host, @())
        try {
            if (Test-TranscriptionSupported) {
                $externalHost.gettype().getproperty("IsTranscribing",
                    [reflection.bindingflags]"NonPublic,Instance").getvalue($externalHost, @())
            } else {
                
            };  # if-E

        } catch {
            Write-Warning "Tested: This host does not support transcription."
            <# 9:19 AM 2/24/2015 the externalhost call is failing and dumping into here
                $Transcript:C:\scripts\logs\set-CSUserPins-20150224-0918-trans.log
                WARNING: Tested: This host does not support transcription.
            #>
        } # try-E
    } else {
        write-host "Test-Transcribing:SKIP PS ISE does not support transcription commands [returning `$true]";
        return $true ; 
    } # if-E
} #*----------------^ END Function Test-Transcribing ^----------------

#*----------------v Function Stop-TranscriptLog v----------------
function Stop-TranscriptLog {
  <#.SYNOPSIS
  Stops & ARCHIVES a transcript file (if no archive needed, just use the stock Stop-Transcript cmdlet)
  .NOTES
  #Written By: Todd Kadrie
  #Website:	http://tinstoys.blogspot.com
  #Twitter:	http://twitter.com/tostka
  Requires test-transcribing() function
  
  Change Log
  # 1:18 PM 1/14/2015 added Lync fs rpt share support
  # 10:54 AM 1/14/2015 added lab support (Server0d\d$)
  # 10:11 AM 12/10/2014 tshot stop-transcriptlog archmove, for existing file clashes
  9:04 AM 12/10/2014 shifted more into the try block
  12:49 PM 12/9/2014

  .INPUTS
  leverages the global $transcript variable (must be set in the root script; not functions)

  .OUTPUTS
  Outputs $TRUE/FALSE reflecting successful archive attempt status

  .EXAMPLE
  Stop-TranscriptLog 
  #>
  
  #can't define $Transcript as a local param/vari, without toasting the main vari!
  if ($bDebug) {"SUB: stop-transcriptlog"}
  
  # 10:48 AM 1/14/2015 adde lab support for archpath
  # 10:56 AM 1/14/2015 adde Lync FS support
  

  if($Host.Name -ne "Windows PowerShell ISE Host"){
        Try {
            if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):`n`$outtransfile:$outtransfile" ;};
                if (Test-Transcribing) {
                    # can't move it if it's locked
                    Stop-Transcript
                    if ($bDebug) {write-host -foregroundcolor green "`$Transcript:$Transcript"} ;
                }  # if-E
        } Catch {
                Write-Error "$(Get-TimeStamp): Failed to move `n$Transcript to `n$Archpath"
                Write-Error "$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName)."
                Write-Error "$(Get-TimeStamp): -- Error information"
                Write-Error "$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)"
                Write-Error "$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)"
                Write-Error "$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)"
                Write-Error "$(Get-TimeStamp): Line: $($_.InvocationInfo.Line)"
                Write-Error "$(Get-TimeStamp): Error Details: $($_)"
        }  # try-E;
  
        if (!(Test-Transcribing)) {  return $true } else {return $false};
    } else {
        write-host "Stop-Transcribing:SKIP PS ISE does not support transcription commands";
        return $true
    }
  

}#*----------------^ END Function Stop-TranscriptLog ^----------------

#*----------------v Function Archive-Log v----------------
function Archive-Log {
  <#.SYNOPSIS
  ARCHIVES a designated file (if no archive needed, just use the stock Stop-Transcript cmdlet). Tests and fails back through restricted subnets to find a working archive locally
  .NOTES
  #Written By: Todd Kadrie
  #Website:	http://tinstoys.blogspot.com
  #Twitter:	http://twitter.com/tostka
  Requires test-transcribing() function
  
  Change Log
  # 7:30 AM 1/28/2015 in use in LineURI script
  # 1:44 PM 1/16/2015 repurposed from Stop-TranscriptLog, focused this on just moving to archive location
  # 1:18 PM 1/14/2015 added Lync fs rpt share support
  # 10:54 AM 1/14/2015 added lab support (Server0d\d$)
  # 10:11 AM 12/10/2014 tshot Archive-Log archmove, for existing file clashes
  9:04 AM 12/10/2014 shifted more into the try block
  12:49 PM 12/9/2014

  .INPUTS
  leverages the global $transcript variable (must be set in the root script; not functions)

  .OUTPUTS
  Outputs $TRUE/FALSE reflecting successful archive attempt status

  .EXAMPLE
  Archive-Log 
  #>
  
	Param(
	[parameter(Mandatory=$true)]
	$FilePath
	)
	
	# 10:37 AM 1/21/2015 moved out of the if\else
	# 10:48 AM 1/14/2015 adde lab support for archpath
	# 10:56 AM 1/14/2015 adde Lync FS support
	#10:34 AM 1/21/2015 fix
	$ArchPathProd ="\\Server0\d$\scripts\rpts\"
	$ArchPathLync = "\\FileServer.domain.com\FileShare\scripts\rpts\";
	$ArchPathLab = "\\Server0D\e$\scripts\rpts\";
	$ArchPathLabLync = "\\FileServer.domainlab.com\FileShare\scripts\rpts\";
	

	if ($bDebug) {"Archive-Log"}
	if(!(Test-Path $FilePath)) {
		write-host -foregroundcolor yellow  "$(Get-TimeStamp):Specified file...`n$Filepath`n NOT FOUND! ARCHIVING FAILED!";
  	} ELSE {


	  #  ip addr test
	  <# prod FE/Edge blocked subnet is: 170.92.9.23
	  Lab FE/Edge blocked subnet is:10.92.9.1
	  #>
	  # lync is blocked for SMB to prod servers; has to use it's own file server for reporting
	  $oIPs=get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -ComputerName . | Select-Object -Property IPAddress ; 
	  $oIPs | foreach {
	    if ($_.ipaddress -match $RgxRestrictedNetwork){
	      #"Y";$_.ipaddress;break;
	      #lync infra server with blocks to prod subnet SMB
	      # lab?
	      
	      #write-host -foregroundcolor yellow  "$(Get-TimeStamp):MESSAGE";
	      if ($bDebug) {write-host -foregroundcolor yellow  "$(Get-TimeStamp):Restricted Subnet Detected. Using Lync ArchPath";};
	      #if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):Restricted Subnet Detected. Using Lync ArchPath"};
	      
	      if($env:USERDOMAIN -eq $DomLab){
	        if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):LAB Server Detected. Using Lync ArchPathLabLync"};
	        # lync Edge or FE server
	        if(test-path $ArchPathLabLync) {
	          $ArchPath = $ArchPathLabLync;
	        } else {
	          write-error "$(get-timestamp):FAILED TO LOCATE Lab Lync `$ArchPath. Exiting.";
	          #Cleanup; # nope, cleanup normally CALLS STOP-TRANSLOG
			  exit
	        };  # if-E
	      } else {
	      	# non-lab prod Lync
	        if(test-path $ArchPathLync) {
	          $ArchPath =$ArchPathLync;
	        } else {
	          write-error "$(get-timestamp):FAILED TO LOCATE Lync `$ArchPath. Exiting.";
	          #Cleanup; # nope, cleanup normally CALLS STOP-TRANSLOG
			  exit
	        }; # if-E non-lab
	      }; # lync lab or prod
	    } else {
	      # non-Lync/restricted subnet
	      if($env:USERDOMAIN -eq $DomLab){
	          if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):LAB Server Detected. Using Lync ArchPathLab"};
	          # lync Edge or FE server
	          if(test-path $ArchPathLab) {
	            $ArchPath = $ArchPathLab;
	          } else {
	            write-error "$(get-timestamp):FAILED TO LOCATE Lab `$ArchPath. Exiting.";
	            #Cleanup; # nope, cleanup normally CALLS STOP-TRANSLOG
			  	exit
	          };  # if-E
	        } else {
	        # non-lab prod 
	          if(test-path $ArchPathProd) {
	            $ArchPath =$ArchPathProd;
	          } else {
	            write-error "$(get-timestamp):FAILED TO LOCATE Prod `$ArchPath. Exiting.";
				# 10:33 AM 1/21/2015 above is last command exec'd/echo'd
				
	            #Cleanup; # nope, cleanup normally CALLS STOP-TRANSLOG
			  	exit
	          };  # if-E
	      }; # if-E non-lab
	    };
	  };  # loop-E

	  if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):`$ArchPath:$ArchPath"};
	  
	  Try {
	    # validate the archpath
	    if ($bDebug) {$ArchPath};
	    if (!(Test-Path $ArchPath)) {
            #$(Read-Host -prompt "Input ArchPath[UNCPath]")
            # $computer = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a computer name", "Computer", "$env:computername") ; 
            $ArchPath = [Microsoft.VisualBasic.Interaction]::InputBox("Input ArchPath[UNCPath]", "Archpath", "") ; 

        }

	    if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):`n`$FilePath:$FilePath `n`$ArchPath:$ArchPath `n" ;};
	   
	        if ($bDebug) {write-host -foregroundcolor green "`$FilePath:$FilePath"}
	        if ((Test-Path $FilePath)) {
	          write-host  ("$(Get-TimeStamp):Moving `n$FilePath `n to:" + $ArchPath) 
	          
	          # 9:59 AM 12/10/2014 pretest for clash
	          
	          $archtarg = (Join-Path $ArchPath (Split-Path $FilePath -leaf));
	          if ($bDebug) {write-host -foregroundcolor green "`$archtarg:$archtarg"}
	          if (Test-Path $archtarg) {
	              $log = Get-ChildItem $FilePath;
	              $archtarg = (Join-Path $ArchPath ($log.BaseName + "-B" + $log.Extension))
	              if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):CLASH DETECTED, RENAMING ON MOVE: `n`$archtarg:$archtarg"};
	              Move-Item $FilePath $archtarg
	          } else {
	            # 8:41 AM 12/10/2014 add error checking
	            $Error.Clear()
	            Move-Item $FilePath $ArchPath 
	          }
	        } else {
	          if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):NO TRANSCRIPT FILE FOUND! SKIPPING MOVE"}
	        }  # if-E
	      
	  } # TRY-E
	  Catch {
	        Write-Error "$(Get-TimeStamp): Failed to move `n$FilePath to `n$Archpath"
	        Write-Error "$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName)."
	        Write-Error "$(Get-TimeStamp): -- Error information"
	        Write-Error "$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)"
	        Write-Error "$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)"
	        Write-Error "$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)"
	        Write-Error "$(Get-TimeStamp): Line: $($_.InvocationInfo.Line)"
	        Write-Error "$(Get-TimeStamp): Error Details: $($_)"
	  } ;
	  
      
	  if (!(Test-Transcribing)) {  return $true } else {return $false};
      
	  
	} # if-E Filepath test
}#*----------------^ END Function Archive-Log ^----------------

#*----------------v Function Start-IseTranscript v----------------
Function Start-IseTranscript {
  <#
   .SYNOPSIS
    This captures output from a script to a created text file
    .NOTES
    NAME:  Start-iseTranscript
    EDITED BY: Todd Kadrie
    AUTHOR: ed wilson, msft
    REVISIONS: 
    * 8:40 AM 3/11/2015 revised to support PSv3's break of the $psise.CurrentPowerShellTab.consolePane.text object
        and replacement with the new...
            $psise.CurrentPowerShellTab.consolePane.text
        (L13 FEs are PSv4, lyn650 is PSv2)
    * 9:22 AM 3/5/2015 tweaked, added autologname generation (from script loc & name)
    * 09/10/2010 17:27:22 - original
    TYPICAL USAGE: 
        Call from Cleanup() (or script-end, only populated post-exec, not realtime)
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        if($Host.Name -eq "Windows PowerShell ISE Host"){
                # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
                $logname= (join-path -path (join-path -path $ScriptDir -childpath "logs") -childpath ($ScriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
                write-host "`$logname: $logname";
                Start-iseTranscript -logname $logname ;
                # optional, normally wouldn't archive ISE debugging passes
                #Archive-Log $logname ;
            } else {
                if($bDebug){ write-host -ForegroundColor Yellow "$(get-timestamp):Stop Transcript" };
                Stop-TranscriptLog ; 
                if($bDebug){ write-host -ForegroundColor Yellow "$(get-timestamp):Archive Transcript" };
                Archive-Log $Transcript ; 
            } # if-E
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


    KEYWORDS: Transcript, Logging, ISE, Debugging
    HSG: WES-09-25-10    
   .DESCRIPTION
   use if($Host.Name -eq "Windows PowerShell ISE Host"){ } to detect and fire this only when in ISE
   .EXAMPLE
    Start-iseTranscript -logname "c:\fso\log.txt"
    Copies output from script to file named xxxxlog.txt in c:\fso folder
   .EXAMPLE
    if($Host.Name -eq "Windows PowerShell ISE Host"){Start-iseTranscript}
    Copies output from script to file named xxxxlog.txt in c:\fso folder (with ISE exec detection & autolog generation)
   .PARAMETER logname
    the name and path of the log file.
   .INPUTS
    [string]
   .OUTPUTS
    [io.file]
   .Link
     Http://www.ScriptingGuys.com
  #Requires -Version 2.0
  #>

  Param(
   [string]$logname
  )
  
   if (!($ScriptDir) -OR !($ScriptNameNoExt)) {
      throw "`$ScriptDir & `$ScriptNameNoExt are REQUIRED values from the main script SUBMAIN. ABORTING!"
  } else {
      if (!($logname)) {
        # build from script if nothing passed in
        <# nope, can't use $MyInvoc from a function, it sees the invoc as the function rather than the parent script filename, 
        # put these in the main script submain
        $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) ;
        $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        #>
        $logname= (join-path -path (join-path -path $ScriptDir -childpath "logs") -childpath ($ScriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
        write-host "`$logname: $logname";
      };
      # 12:11 PM 3/5/2015 strange crash here, maybe it's because of indenting the herestring.
$transcriptHeader = @"
**************************************
Windows PowerShell ISE Transcript Start
Start Time: $(get-date)
UserName: $env:username
UserDomain: $env:USERDNSDOMAIN
ComputerName: $env:COMPUTERNAME
Windows version: $((Get-WmiObject win32_operatingsystem).version)
**************************************
Transcript started. Output file is $logname
"@
        #$transcriptHeader >> $logname
        $transcriptHeader | out-file $logname –append
        #$psISE.CurrentPowerShellTab.Output.Text >> $logname
        <# 8:37 AM 3/11/2015 PSv3 broke/hid the above object, new object is
        $psISE.CurrentPowerShellTab.ConsolePane.text
        Note, it's reportedly not realtime, as the Psv2 .type param was
        #>
        if (($host.version) -lt "3.0") {
            # use legacy obj
            $psISE.CurrentPowerShellTab.Output.Text | out-file $logname –append
        } else {
            # use the new object
            $psISE.CurrentPowerShellTab.ConsolePane.text | out-file $logname –append
        } # if-E
    }  # if-E
} #*----------------^ END Function start-iseTranscript ^---------------- 

#*----------------v Function Start-TranscriptLog v----------------
function start-TranscriptLog {
  <#.SYNOPSIS
  Configures and launches a transcript
  .NOTES
  #Written By: Todd Kadrie
  #Website:	http://tinstoys.blogspot.com
  #Twitter:	http://twitter.com/tostka
  Requires test-transcribing() function
  
  Change Log
  # 10:19 AM 12/10/2014 cleanup
  # 8:48 AM 12/10/2014 cleanup
  12:36 PM 12/9/2014

  .INPUTS
  None

  .OUTPUTS
  Outputs $TRUE/FALSE reflecting transcribe status

  .EXAMPLE
  start-TranscriptLog $Transcript
  #>

  # $outransfile=$LogDir + $ScriptNameNoExt + "-" + $user.ToUpper()  + "-(" + $ticket + ")-" + $TimeStampNow + "-trans.txt" ;
  
  param(
    [parameter(Mandatory=$true,Helpmessage="Transcript location")]
    [ValidateNotNullOrEmpty()]
    [alias("tfile")]
    [string]$Transcript
  )

  # Have to set relative $ScriptDir etc OUTSIDE THE FUNC, build full path to generic core $Transcript vari, and then 
  # start-transcript will auto use it (or can manual spec it with -path)

    if($Host.Name -NE "Windows PowerShell ISE Host"){
        Try {
                if (Test-Transcribing) {Stop-Transcript} 
  
                if($bdebug) {$Transcript}
                # prevaidate specified logging dir is present
                $TransPath=(Split-Path $Transcript).tostring();
                if($bdebug) {$TransPath;}
                if (Test-Path $TransPath ) { } else {mkdir $TransPath};
                #invoke-pause2
                Start-Transcript -path $Transcript
                if (Test-Transcribing) {  return $true } else {return $false};
            } Catch {
                Write-Error "$(Get-TimeStamp): Failed to create $TransPath"
                Write-Error "$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName)."
                Write-Error "$(Get-TimeStamp): -- Error information"
                Write-Error "$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)"
                Write-Error "$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)"
                Write-Error "$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)"
                Write-Error "$(Get-TimeStamp): Line: $($_.InvocationInfo.Line)"
                Write-Error "$(Get-TimeStamp): Error Details: $($_)"
            }  # try-E;

    } else {
        write-host "Test-Transcribing:SKIP PS ISE does not support transcription commands [returning $true]";
        return $true ; 
    };  # if-E


}#*----------------^ END Function Start-TranscriptLog ^----------------

#*----------v Function load-LMS v----------
function load-LMS {
  <#
  .SYNOPSIS
  Checks local machine for registred Lync LMS, and then loads the newest one found
  .NOTES
  Written By: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka

  Additional Credits: [REFERENCE]
  Website:	[URL]
  Twitter:	[URL]

  Change Log
  vers: 10:43 AM 1/14/2015 fixed return & syntax expl to true/false
  vers: 10:20 AM 12/10/2014 moved commentblock into function
  vers: 11:40 AM 11/25/2014 adapted to Lync
  ers: 2:05 PM 7/19/2013 typo fix in 2013 code
  vers: 1:46 PM 7/19/2013
  .INPUTS
  None.
  .OUTPUTS
  Outputs Lync Revision
  .EXAMPLE
  $LMSLoaded = load-LMS ; Write-Debug "`$LMSLoaded: $LMSLoaded" ;
  #>
  # check registred v loaded ;
  $ModsReg=Get-Module -ListAvailable;
  $ModsLoad=Get-Module;
  if ($ModsReg | where {$_.Name -eq "Lync"}) {
    if (!($ModsLoad | where {$_.Name -eq "Lync"})) {
      Import-Module Lync -ErrorAction Stop ;return $TRUE;
    } else {
      return $TRUE;
    }
  } else {
    Write-Error {"$(Get-TimeStamp):($env:computername) does not have Lync Admin Tools installed!";};
    return $FALSE
  }
} #*----------^END Function load-LMS ^----------

#*----------------v Function Send-EmailNotif v----------------
Function Send-EmailNotif {

  # 10:35 AM 8/21/2014 always use a port; tested for $SMTPPort: if not spec'd defaulted to 25.
  # 10:17 AM 8/21/2014 added custom port spec for access to Server0:8111 from my workstation

  <#
  PARAM(
    [parameter(Mandatory=$true)]
    [alias("from")]
    [string] $SMTPFrom,
    [parameter(Mandatory=$true)]
    [alias("to")]
    [string] $SmtpTo,
    [parameter(Mandatory=$true)]
    [parameter(Mandatory=$true)]
    [alias("subj")]
    [string] $SMTPSubj,
    [parameter(Mandatory=$true)]
    [alias("server")]
    [string] $SMTPServer,
    [parameter(Mandatory=$true)]
    [alias("port")]
    [string] $SMTPPort,
    [parameter(Mandatory=$true)]
    [alias("body")]
    [string] $SmtpBody,
    [parameter(Mandatory=$false)]
    [switch] $SmtpBodyHtml,
    [parameter(Mandatory=$false)]
    [alias("attach")]
    [string] $attachment 
  )
  #>
					
    # before you email conv to str & add CrLf:
    $SmtpBody = $SmtpBody | out-string
    # just default the port if missing, and always use it

    # 9:06 AM 12/14/2015 baseline hash, no port no attachment, and add the fields as needed/supported
    $email = @{
        From = $SMTPFrom ; 
        To = $SMTPTo ; 
        Subject = $SMTPSubj ; 
        SMTPServer = $SMTPServer ; 
        Body = $SmtpBody ; 
    } ; 
    # # 9:05 AM 12/14/2015: actually port param isn't supported in PSv2 send-mailmessage; 25 is assumed
    if ($($host.version.major -gt 2) -and ($SMTPPort -eq $null)) {
        $SMTPPort = 25;
        $email.add("Port",$SMTPPort );
    }	else {
        if($SMTPPort) {write-verbose -verbose:$true  "Skipping specified 'SMTPPort' parameter: Powershellv2 does not support a Port specification for send-mailmessage (defaults to 25 only)"} ; 
    } # if-block end ; 

    # define/update variables into $email splat for params
    if ($attachment -AND (test-path $attachment)) {
        $email.add("Attachments",$attachment);
    } ; 

    if($SmtpBodyHtml){
        #$email.BodyAsHtml = $true
        $email.add("BodyAsHtml",$true);
    } # if-E
	
    <# 11:59 AM 8/28/2013 mailing debugging code
    write-host  -ForegroundColor Yellow "Emailing with following parameters:"
    $email
    write-host "body:"
    $SmtpBody
    write-host ("-" * 5)
    write-host "body.length: " $SmtpBody.length
    #>
    write-host "sending mail..."
    $email
    $error.clear() 

    # 8:57 AM 1/29/2015 add a try/catch to it, to echo full errors
    TRY {
        #invoke-pause2
        send-mailmessage @email 
    } Catch {
        Write-Error "$(Get-TimeStamp): Failed send-mailmessage attempt"
        Write-Error "$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName)."
        Write-Error "$(Get-TimeStamp): -- Error information"
        Write-Error "$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)"
        Write-Error "$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)"
        Write-Error "$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)"
        Write-Error "$(Get-TimeStamp): Line: $($_.InvocationInfo.Line)"
        Write-Error "$(Get-TimeStamp): Error Details: $($_)"
    } ; # try/catch-E

    # then pipe just the errors out to console
    #if($error.count -gt 0){write-host $error }

}#*----------------^ END Function Send-EmailNotif ^---------------- ;  

#*----------------v Function Test-Port() v----------------
# attempt to open a port (telnet xxx.yyy.zzz nnn)
# call: Test-Port $server $port
# vers: 8:42 AM 7/24/2014 added proper fail=$false
# vers: 10:25 AM 7/23/2014 disabled feedback, added a return
function Test-Port
{
  PARAM(
  [parameter(Mandatory=$true)]
  [alias("s")]
  [string]$Server,
  [parameter(Mandatory=$true)]
  [alias("p")]
  [int]$port)
  $ErrorActionPreference = “SilentlyContinue”
  $socket = new-object Net.Sockets.TcpClient
  $socket.Connect($Server, $port)
  if ($socket.Connected)
  {
    #write-host "We have successfully connected to the server" -ForegroundColor Yellow -BackgroundColor Black
    $socket.Close()
    # 9:54 AM 7/23/2014 added return true/false
    return $True;
  } # if-block end
  else
  {
   #write-host "The port seems to be closed or you cannot connect to it" -ForegroundColor Red -BackgroundColor Black
   return $False;
  } # if-block end
  $socket = $null
}#*----------------^ END Function Test-Port() ^----------------

#---------------------------v Function Cleanup v---------------------------
function Cleanup {
    # clear all objects and exit
    # Clear-item doesn't seem to work as a variable release 
   
    # 8:39 AM 12/10/2014 shifted to stop-transcriptLog function
    # 7:43 AM 1/24/2014 always stop the running transcript before exiting
    if ($bDebug) {"CLEANUP"}
    #stop-transcript
    # 11:16 AM 1/14/2015 aha! does this return a value!??
    if($Host.Name -eq "Windows PowerShell ISE Host"){
        # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
        $logname= (join-path -path (join-path -path $ScriptDir -childpath "logs") -childpath ($ScriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
        write-host "`$logname: $logname";
        #Start-iseTranscript
        Start-iseTranscript -logname $logname ;
        #Archive-Log $logname ;
    } else {
        if($bDebug){ write-host -ForegroundColor Yellow "$(get-timestamp):Stop Transcript" };
        Stop-TranscriptLog ; 
        if($bDebug){ write-host -ForegroundColor Yellow "$(get-timestamp):Archive Transcript" };
        #Archive-Log $Transcript ; 
    } # if-E
    exit
} #*----------------^ END Function Cleanup ^----------------

# *** ENDREGION SETUP MARKER
#endregion SETUP

#*================^ END FUNCTIONS  ^================

#*----------------v SUB MAIN v----------------

# *** REGION MARKER LOAD
#region LOAD
# *** LOADING 

$LMSLoaded = load-LMS ; write-host -foregroundcolor green "`$LMSLoaded: $LMSLoaded" ;
#$ADMTLoaded = load-ADMS ; write-host -foregroundcolor green "`$ADMTLoaded: $ADMTLoaded" ;

#====== V OUTPUT FILE HANDLING BOILERPLATE (USE IN SUB MAIN) V==================================
# build outtransfile name: 
# try to leverage the global variable, that start-Transcript should leverage
# transcript with userid & ticket # C:\usr\work\lync\scripts\logs\enable-luser-KLUEGJS-(328639)-trans.log
# need explicit timestamp in the $Transcript
$TimeStampNow = Get-Date -uformat "%Y%m%d-%H%M" ;
# user logon & ticket tlog name
#$Transcript = ( (Split-Path -parent $MyInvocation.MyCommand.Definition) + "\logs\" + ([system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName)) + "-" + $logon.ToUpper()  + "-(" + $ticket + ")-" + $TimeStampNow + "-trans.log" ) ;
# name for the script & pass time
$Transcript = ((Split-Path -parent $MyInvocation.MyCommand.Definition) + "\logs\" + ([system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ) + "-" + $TimeStampNow + "-trans.log");
if ($bDebug) {write-host -foregroundcolor green "$(Get-TimeStamp):`n`$Transcript:$Transcript";};
  
start-TranscriptLog $Transcript
#====== ^ OUTPUT FILE HANDLING BOILERPLATE (USE IN SUB MAIN) ^==================================

write-host ("$(Get-TimeStamp):`nPASS STARTED" + ("="*5)) ;

# -----------

# generate a 'ticket' for this batch 20150108-1043AM get-date -format "yyyyMMdd-HHmm"
#20111118-0906

$sTicket=(Get-Date -format "yyyyMMdd-HHmm");

#====== V EMAIL HANDLING BOILERPLATE (USE IN SUB MAIN) V==================================
#$ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
# moved up script to use for include file loads
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
$TimeStampNow = get-date -uformat "%Y%m%d-%H%M" ;
# add explici join 
$outfile=(join-path -Path $ScriptDir -ChildPath "logs")
if (!(test-path $outfile)) {Write-Host "Creating dir: $outfile" ;mkdir $outfile ;} ;
$outfile+="\" + $ScriptNameNoExt + "-" + $TimeStampNow + "-processing.csv" ;
$foutfile = $outfile.replace("-processing","-fails");
$Rptfile=(Join-Path -Path (join-path -path $ScriptDir -childpath "logs") -ChildPath ($ScriptNameNoExt + "-$($env:COMPUTERNAME)-" + $TimeStampNow + "-RPT.txt")) ;

$CrashRptfile=(Join-Path -Path (join-path -path $ScriptDir -childpath "logs") -ChildPath ($ScriptNameNoExt + "-" + $TimeStampNow + "-CRASH-LASTUSER.csv")) ;
#========
$ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName)) ;
$SMTPFrom = (($ScriptBaseName.replace(".","-")) + "@domain.com") ; 
$SMTPSubj= ("POSSIBLE REBOOT?: "+ (Split-Path $transcript -Leaf) + " " + [System.DateTime]::Now) ; 

if ($env:COMPUTERNAME -eq 'LYN-3V6KSY1') {
  $SMTPServer = "Server0";
  $SMTPPort = 8111 ;
} elseif ($env:USERDOMAIN -eq "domain-LAB"){
	# lab
	write-host ("$(get-timestamp):domain-LAB: Using `$SMTPServer:$SMTPLab")
	$SMTPServer = $SMTPLab ;
}else{
  $SMTPServer = "smtp.domain.com" ;
} # if-E

Write-Host ((get-date).ToString("HH:mm:ss") + ":`$SMTPServer:" + $SMTPServer ) ;
$SMTPTo="emailadmin@domain.com" ;
#$attachment=$Rptfile
# setup body as a hash
$SmtpBody = @() ;
# (`n = CrLf in body)
#====== ^ EMAIL HANDLING BOILERPLATE (USE IN SUB MAIN) ^==================================



# pull all pools in SITE
$L13FEs = (get-cspool| ?{($_.Services -like '*Registrar*') -AND ($_.Site -like '*LyncSITE*')} | select computers).computers; 

#Access the array: $fecomps.computers | %{ping $_ -n 1 };
# ping them all
#($fecomps).computers | foreach {write-host $_; ping $_ -n 1;};

write-host ("=" * 15);
write-host -foregroundcolor green "$(Get-TimeStamp):`n`Checking Lync Service Health on $ComputerName `n" ;

write-host -foregroundcolor green "$(Get-TimeStamp):`n`Local Server: $($computername) : " ;
$svcs = (Get-CsWindowsService -computername $env:COMPUTERNAME | select name,status );
$svcs | Format-Table -AutoSize ;
#$DeadSvc = (Get-CsWindowsService -computername $FE | ?{$_.Status -ne 'Running'}  );
$DeadSvc = ($svcs | ?{$_.Status -ne 'Running'}  );
if ($DeadSvc){    write-host -foregroundcolor red "STOPPED SVCS";
    $DeadSvc | ft Name,Status;
} else {
    write-host "[no non-runnicng svcs]";
} ;
write-host ("-" * 3);

# $Rptfile
# build it
"$(Get-TimeStamp):`n`Checking Lync Service Health on $ComputerName" | out-file $Rptfile ;
"`nCurrent Get-CsWindowsService:" | out-file $Rptfile -Append #-Encoding ascii ;
$svcs | Format-Table -AutoSize |Out-File $Rptfile -Append #-Encoding ascii ;
if ($DeadSvc){    
    "STOPPED SVCS" | out-file $Rptfile -Append #-Encoding ascii ;
    $DeadSvc | ft -AutoSize | out-file $Rptfile -Append #-Encoding ascii ;
} else {
    "[no non-runnicng svcs]"| out-file $Rptfile -Append #-Encoding ascii ;
} ;


#---------------------------------------------------------------------------------


#region REPORT ; 
# 12:13 PM 1/26/2015 mail a report
#Load as an attachment into the body text:
$SmtpBody += ("Pass Completed "+ [System.DateTime]::Now) ;
$SmtpBody += (Get-Content $rptfile );

if($bdebug){Write-Host -ForegroundColor Green "Testing SmtpServer ($SmtpServer) Access...";};
if (!($env:USERDOMAIN -eq $DomLab)){
			Send-EmailNotif
}else {
		Write-Host -ForegroundColor Red "skipping email, $SmtpServer unreachable on port 25 in Lab"
		if(!($bDebug)){ break }; 
} # if-E lab

Cleanup; 
# *** REGION REPORT END MARKER
#endregion REPORT 
# ========================================
