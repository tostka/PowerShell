# remove-BadCerts.ps1

  <# 
  .SYNOPSIS
  remove-BadCerts.ps1 - remove bad misplaced IntermedCA  certs being pushed by  domain-level gpo into RootCA! (breaks all  cert-authentication in 2012R2): 
  .NOTES
  Written By: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka
  Change Log
  * 7:38 AM 11/12/2015
  .DESCRIPTION
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  .LINK
  *---^ END Comment-based Help  ^--- #>

# target ANY Intermed CA certs that are in the Trusted Root store, and will break Win2012R2 Cert trust chain
$badcerts = (Get-Childitem cert:\LocalMachine\root -Recurse | Where-Object {$_.Issuer -ne $_.Subject}) ; 

if($badcerts){

    # only log when we're deleting/found something
    #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    # generic code to run a transcript recycled out of the Script FileName-yyyymmdd-hhmm-trans.log
    #region SETUP
    $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
    $TimeStampNow = get-date -uformat "%Y%m%d-%H%M"
    $outtransfile=$ScriptDir + "logs"
    if (!(test-path $outtransfile)) {Write-Host "Creating dir: $outtransfile" ;mkdir $outtransfile ;} ;
    $outtransfile+="\" + $ScriptNameNoExt + "REMOVAL-" + $TimeStampNow + "-trans.log" ;
    Trap {Continue} Stop-Transcript ;
    start-transcript -path $outtransfile ;
    #endregion SETUP
 
    write-host -foregroundcolor red "$((get-date).ToString("HH:mm:ss")):BAD CERTS..." ;
    $badcerts | select Subject,Thumbprint | format-list ; 
    write-host -foregroundcolor red "...are _Intermediate CA_ certs, in the _LM\Trusted Root_ Certificates container!`n";
    write-host -foregroundcolor red "$((get-date).ToString("HH:mm:ss")):PURGING THE PROBLEM CERTS!";

    foreach ($cert in $badcerts){
        "$($cert.subject)" ; 
        write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Removing cert: `n$($cert.Subject)`n ($($cert.Thumbprint))...";
    
        $error.clear() ;
        TRY {
            $cert | Remove-Item –Force #-whatif ; 
        } catch {
            $msg=": Error Details: $($_)";
            Write-Error "$(Get-TimeStamp): FAILURE!" ;
            # opt extended error info...
            Write-Error "$(Get-TimeStamp): Error in $($_.InvocationInfo.ScriptName)." ; 
            Write-Error "$(Get-TimeStamp): -- Error information" ;
            Write-Error "$(Get-TimeStamp): Line Number: $($_.InvocationInfo.ScriptLineNumber)" ;
            Write-Error "$(Get-TimeStamp): Offset: $($_.InvocationInfo.OffsetInLine)" ;
            Write-Error "$(Get-TimeStamp): Command: $($_.InvocationInfo.MyCommand)" ;

            Write-Error "$(Get-TimeStamp): Line: $($_.InvocationInfo.Line)" ;
            #Write-Error "$(Get-TimeStamp): Error Details: $($_)" ;
            $msg=": Error Details: $($_)" ;
            Write-Error  "$(Get-TimeStamp): $($msg)"; 
            # log the processing failure & status, append to aggregator
    
            # 1:00 PM 1/23/2015 autorecover from fail, STOP (debug), EXIT (close), or use Continue to move on in loop cycle
            Continue
            #
        };   # try/catch-E

  }  # loop-E;

  write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Post processing status:";
  Get-Childitem cert:\LocalMachine\root -Recurse | Where-Object {$_.Issuer -ne $_.Subject} | select subject,thumbprint | format-list ;
  
  stop-transcript
} else {
    write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):No problematic certs found. Exiting without changes.";
} # if-E

