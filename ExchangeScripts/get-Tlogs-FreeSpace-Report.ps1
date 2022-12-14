# get-Tlogs-FreeSpace-Report.ps1

<#
.SYNOPSIS
get-Tlogs-FreeSpace-Report-Ex10.ps1 - run daily logparser top-EAS report, and email the results (also usable on-demand by IOC staff, via schedtask)
.NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka
Change Log
vers: 2:59 PM 12/23/2013 COMPANY port completed, basic function in; needs subdir/db id for SITE&SITE; could use tlog counts, for all (since edb & tlogs are comingled)
  for tracking actual tlog accumulation, in isolation from replic & edb footprint
vers: 10:49 AM 12/23/2013 COMPANY port
vers: 12:11 PM 10/25/2013 shifted to include files
vers: 8:00 AM 10/10/2013 did some cleanup, tracked down & disabled debugging code that was reporting falacious errors.
vers: 10:14 AM 9/24/2013 updated send-emailnofication() to latest vers
Vers: 8:35 AM 9/16/2013 added timestamp to console for ref
# vers: 1:01 PM 8/30/2013 replaced win32_perfformatteddata_perfdisk_logicaldisk code with win32_volume calls, to get Capacity (not supported in perfdisk)
# VERS: 11:50 AM 8/30/2013 completed, and added tweak: removed version-specific path code and replaced with generic filter: '*:\ML*'
# VERS: 2:28 PM 8/28/2013 pretty complete
# VERS: 12:55 PM 8/28/2013 fixed output to log, and mailing. Also removed split versions, now all in one. Also added parameters, prmpting and validation
# VERS: 1:03 PM 8/16/2013 - strip back EMS items - to make it run on mailexp, with or without Ex2010 ems avail
# vers: 1:28 PM 8/15/2013 - baseically working, except it's seeing dag servers in susyd & gbmk
# vers: 1:24 PM 8/7/2013 then did coding a daily logparser output, top 50 EAS users on prior 12hrs of logs
# vers: 11:07 AM 8/14/2013 first did a straight port of the tlog reporter for UNI, runs fine remotely
# vers: 2:12 PM 2/10/2012 forgot to change tom's addr on $smtptonotify
# vers: 8:03 AM 2/9/2012, shifted outfile to reports subdir, and made it read targsrvr from local comp name (run on any, not just 52 now), and added an email to TomG
# vers: 2:34 PM 11/18/2011: debugged function under schtask
# vers: 11:11 AM 11/18/2011 added freespace percentage calc
# vers: 11:45 AM 11/17/2011
# pulls tlog status info from the DAG, normally run on ex52
# # TASK ADD CMD:
#schtasks /create /s Server /tn "get-Tlogs-FreeSpace-Report-Ex10" /tr E:\SCRIPTS\get-Tlogs-FreeSpace-Report-Ex10-ps1.cmd /sc DAILY /ST 04:00 /ru na\ea-backup /rp
# TESTFIRE
# schtasks /run /tn "get-Tlogs-FreeSpace-Report-Ex10"
# update pw:
#/RP  [password]    Specifies the password for the "run as" user.
#                    To prompt for the password, the value must be either
#                    "*" or none. This password is ignored for the
#                    system account. Must be combined with either /RU or
#                    /XML switch.
#SCHTASKS /Change /s Servern2 /TN "task-mbx-size-rpt-e2k3" /RP <password>
.DESCRIPTION
Run daily tlog space report, and email the results (also usable on-demand by IOC staff, via schedtask)
.PARAMETER  $TargetServer
Infrastructure subset to be analyzed [NA|SITE|AP|EU]
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
System.Boolean
            True if the current Powershell is elevated, false if not.
[use a | get-member on the script to see exactly what .NET obj TypeName is being returning for the info above]
.EXAMPLE
.\get-Tlogs-FreeSpace-Report.ps1 -SiteName Server$3 -ExTargVers 2007
.LINK
#>

#------------- Declaration and Initialization of Global Variables -------------	

PARAM(
  [alias("s")]
  [string] $SiteName=$(Read-Host -prompt "Specify Site to analyze [SiteName: NA | EU | AU ]"),
  [alias("v")]
  [string] $ExTargVers="2010"
) # PARAM BLOCK END

## immed after param, setup to stop executing the script on the first error
trap { break; }

$bDebug = $TRUE
#$bDebug = $FALSE

If ($bDebug -eq $TRUE) {write-host "*** DEBUGGING MODE ***"}

# ----- V INCLUDES BLOCK V------
$ScriptDir = (Split-Path -parent $MyInvocation.MyCommand.Definition) + "\"
# load standard includes...
write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ":LOADING INCLUDES:")
# constants file
# 10:50 AM 12/23/2013: updated for CO
$sLoad=($ScriptDir + "incl-CO-consts.ps1") ; if (test-path $sLoad) {
  write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad } else {write-warning ((get-date).ToString("HH:mm:ss") + ":MISSING"+ $sLoad + " EXITING...") ; exit}
# functions file
# 10:50 AM 12/23/2013: updated for CO
$sLoad=($ScriptDir + "incl-funcs-base-CO.ps1") ; if (test-path $sLoad) {
  write-host -foregroundcolor darkgray ((get-date).ToString("HH:mm:ss") + ":"+ $sLoad) ; . $sLoad } else {write-warning ((get-date).ToString("HH:mm:ss") + ":MISSING"+ $sLoad + " EXITING...") ; exit}
# ----- ^ INCLUDES BLOCK ^------

$SiteName=$SiteName.ToUpper()
$ExTargVers=$ExTargVers.ToUpper()
# 12:35 PM 8/28/2013 validate the inputs
# NA|EU|AU
if (!($SiteName -match "^(NA|EU|AU)$")) {
  write-warning ("INVALID SiteName SPECIFIED: " + $SiteName + ", EXITING...")
  exit 1
}
if (!($ExTargVers -match "^(2007|2010)$")) {
  write-warning ("INVALID ExTargVers SPECIFIED: " + $ExTargVers + ", EXITING...")
  exit 1
}

# remote task exec server
$runsServer = "Server7"

$ScriptBaseName = [system.io.path]::GetFilename($MyInvocation.InvocationName)
$ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName)  

If ($bDebug -eq $TRUE) {
  write-host ("`$ScriptDir: " + $ScriptDir)
  write-host ("`$ScriptNameNoExt: " + $ScriptNameNoExt)
  write-host ("`$ScriptBaseName: " + $ScriptBaseName)
}

$TimeStampNow = get-date -uformat "%Y%m%d-%H%M" 

write-host " "
# confirm log dir exists
$testPath= $ScriptDir + "logs\" 
if (!(test-path $testPath)) {
  write-host "Creating " $testPath
  New-Item $testPath -type directory    
}

# 10:59 AM 12/23/2013 COMPANY upd
$SMTPFrom = (($ScriptBaseName.replace(".","-")) + "@DOMAIN.com") ; 
$SMTPServer = "Server0.domain.com" ; 
$SmtpToAdmin="todd.kadrie@DOMAIN.com" ; 
$SMTPTo=$SmtpToAdmin ; 
$SmtpToNotify="address@domain.com" ; 
$SMTPTo2=$SmtpToNotify ; 

# setup body as a hash
$SmtpBody = @() ; 


#*================v FUNCTION LISTINGS v================

#---------------------------v Function Cleanup v---------------------------
Function Cleanup {
  # clear all objects and exit
  # Clear-item doesn't seem to work as a variable release 

  exit
}
#*----------------^ END Function Cleanup ^----------------
#*================^ END FUNCTION LISTINGS ^================

#*================v SCRIPT BODY v================
#*---------------------------v Function Get-TLogSpaceRegion v---------------------------
Function Get-TLogSpaceRegion ($ExTargVers,$SiteName) {

	#start-transcript -path $outransfile
	  
  If ($bDebug -eq $TRUE) {
      Write-Host ("Get-TLogSpaceRegion,ExTargVers: " + $ExTargVers)
  }
	# lose site looping - when there's an issue, it's regional at best, use explicit passes per Site, and poll all matching servers.

  Write-Host -ForegroundColor White ("Site: " + $SiteName)
  
  #$ExTargVers = 14 
  # NA | EU | AU
  #if ($ExTargVers -eq 2010 -and $SiteName -ne 'Server$3') {
  if ($SiteName -eq 'NONE') {
  
    Clear-Variable outfile -ErrorAction SilentlyContinue
    
    Write-Host -ForegroundColor Yellow ("No Ex2010 in site: " + $SiteName + " skipping...")
    $SmtpBody +="No Ex2010 in site: " + $SiteName + " skipping..."
    #$attachment = $ScriptDir + "logs\" + $ScriptNameNoExt + "-" + $targServer + "-" + $TimeStampNow + ".txt"
    #$SMTPSubj= ("Daily Rpt: "+ (Split-Path $attachment -Leaf) + " " + [System.DateTime]::Now)
    $SMTPSubj= ("Daily Rpt: Site:" + $SiteName + " " + " ExRev: " + $ExTargVers  + ", " + [System.DateTime]::Now)
    
    Send-EmailNotif
    
  } Else {
    
    # 1:06 PM 8/28/2013 drop role
    $sMsg = "Checking Exch vers " + $ExTargVers 
    if ($ExClstr -ne $null) {$sMsg += " (" + $ExClstr + " replic)" }
    $sMsg += " servers in site " + $SiteName
    Write-Host -ForegroundColor green ($sMsg)
    
    
    # $ExTargVers [8|14],$SiteName as per above
    If ($bDebug -eq $TRUE) {
      #write-host "ServersMbxNAF: " $ServersMbxNAF
      write-host "SiteName: " $SiteName
      write-host "ExTargVers: " $ExTargVers
    } # if-block end
    
    # NA|EU|AU
    switch ($SiteName)
    {
      "NA" {
        If ($bDebug -eq $TRUE) { write-host "US BLOCK"}
        switch ($ExTargVers.ToSTring()){
          "2007" {$targServers=$ServersMbxNA}
          "2010" {$targServers=$ServersMbxNA}
        } # switch block end
      } # switch entry end
      "EU" {$targServers=$ServersMbxEU}
      "AU" {$targServers=$ServersMbxAU}
    } # switch block end
    
    If ($bDebug -eq $TRUE) { write-host "targServers: " $targServers }
    
    if ($targServers.Count -eq $null) {
      Write-Host -ForegroundColor Yellow ("No matching servers in site " + $SiteName)
    } else {
            
      Write-Host ("targServers count: " + $targServers.Count)
      If ($bDebug -eq $TRUE) {$targServers}
      foreach ($targServer in $targServers) {
        # 8:34 AM 9/16/2013 timestamp added
        Write-Host ("Time: " + (get-date).ToString("HH:mm:ss"))
        #$outransfile=$ScriptDir + "logs\" + $ScriptNameNoExt + "-" + $TimeStampNow + "-trans.txt"
        $outransfile=$ScriptDir + "logs\" + $ScriptNameNoExt + "-" + $targServer + $TimeStampNow + "-trans.txt"
        # delete any existing transcript file
        # 11:59 AM 11/18/2011 check first, avoids an error
        if (test-path $outransfile) {remove-item $outransfile -force }

        write-host  -ForegroundColor Yellow ("`$targServer: " + $targServer)
        Clear-Variable outfile -ErrorAction SilentlyContinue
        
        $attachment = $ScriptDir + "logs\" + $ScriptNameNoExt + "-" + $targServer + "-" + $TimeStampNow + ".txt"
        if (test-path $attachment) {write-host -foregroundcolor yellow "removing existing" $attachment ; remove-item $attachment -force }
        write-host ("outfile: " + $attachment)
        
        # dash add to body
        $SmtpBody += ('-' * 50)
        $SmtpBody +=$targServer + " pass started," + (((get-date).ToString("HH:mm:ss")))
        
        $SMTPSubj= ("Daily Rpt: "+ (Split-Path $attachment -Leaf) + " " + [System.DateTime]::Now)

        # exception for PF server (no MPs)
        # 2:26 PM 12/23/2013 EXCEPTION for SITE-site dag; uses vmp's, named
        if (($targServer).ToUpper() -match '(?i:((SITE|BCC)MS64\d))') {

            
               # Volume-mount point; non-PF mailbox servers; those with luns with names with distinct substrings to target
              #$sExcCmd = "Get-Wmiobject -query 'select name,driveletter,capacity,freespace from win32_volume where drivetype=3 AND driveletter=NULL' -computer " + $targServer 
              # 2:20 PM 12/23/2013 COMPANY 
              $sExcCmd = "Get-Wmiobject -query 'select name,driveletter,capacity,freespace from win32_volume where drivetype=3' -computer " + $targServer 
              #$sExcCmd = $sExcCmd + " | where {(`$_.Name -like '*:\ML*')}"
              # F:\Server2
              # matches e:|f: SITE|bcc, mail str and any digit
              $sExcCmd = $sExcCmd + " | where {(`$_.Name -match '(?i:((E|F):\\(SITE|SITE).*MAIL0\d\\))')}"
        
                    # match SITE|SITE servers
        } elseif  (($targServer).ToUpper() -match '(?i:((SITE|SITE)MS64\d))' ) {
          # 12:36 PM 8/30/2013 original win32_perfformatteddata_perfdisk_logicaldisk no-capacity code 
          # pull the drive letter filter (which will come back F:; name comes back F:\
          $sExcCmd = "Get-Wmiobject -query 'select name,driveletter,capacity,freespace from win32_volume where drivetype=3' -computer " + $targServer 
          # regex e-h:
          $sExcCmd = $sExcCmd + " | where {(`$_.Name -match '(?i:((E|F|G|H):\\))')}"

              
        } # if-block end
        
        <#	  test commands...
        #get-wmiobject win32_perfformatteddata_perfdisk_logicaldisk -ComputerName $targServer| where {($_.Name -ne "F:") -AND ($_.Name -ne "E:") -AND ($_.Name -ne "_Total")} | sort Name | select Name, @{ Label ="FreeSpaceGB"; Expression={($_.FreeMegabytes/1024).tostring("F0")}}, @{ Label ="FreeSpace%"; Expression={'{0:P1}' -f ($_.percentfreespace/100)}}| ft -autosize| out-file -filepath $attachment
        # testing command, to retrieve all drives/luns/WMP's on the target box		
        #get-wmiobject win32_perfformatteddata_perfdisk_logicaldisk -ComputerName "Server$3-fedexch2"| select Name  
        
        #>
        
        # 12:00 PM 8/30/2013 sub in win32_volume code, which supports capacity attrib (but doesn't appear to return freespace as % (has to be calcd)
        $sExcCmd = $sExcCmd + " | Sort Name | Select Name,@{Name='VolSize(gb)';Expression={[decimal]('{0:N1}' -f(`$_.capacity/1gb))}},@{Name='Freespace(gb)';Expression={[decimal]('{0:N1}' -f(`$_.freespace/1gb))}},@{Name='Freespace(%)';Expression={'{0:P2}' -f((`$_.freespace/1gb)/(`$_.capacity/1gb))}}"
        
        # 10:57 AM 8/28/2013 fresh tee use, putting tee inside the invoked string
        $sExcCmd=$sExcCmd + " | tee -FilePath $attachment" 
        
        If ($bDebug -eq $TRUE) {
          write-host "sExcCmd at Invoke:"
          $sExcCmd 
        }

        Invoke-Expression $sExcCmd 
        
        If ($bDebug -eq $TRUE) {
          Write-Host  -ForegroundColor Yellow "Pass completed"

          Write-Host "---"
        }
        else {
        
        }

        
        $SmtpBody +=$targServer + " pass completed," + (((get-date).ToString("HH:mm:ss"))) + "`nResults Attached: " + (gci $attachment).Name
        # add dash divider 
        $SmtpBody += ('-' * 50)

        Send-EmailNotif
        Write-Host ('='*50)
        
      } # server for-loop end
    }  # if-block end NoServers test
  }  # if-block end Vers/Site exclusion

	
	#stop-transcript

	Cleanup

}
#*---------------------------^ End Function Get-TLogSpaceRegion ^---------------------------

# 12:13 PM 10/25/2013 shifted to include
#*----------------v Function Send-EmailNotif v---------------- ; 
#Function Send-EmailNotif {

<# 
    .SYNOPSIS
Send-EmailNotif() - Generic send-mailmessage wrapper function

    .NOTES
Written By: Todd Kadrie
Website:	http://tinstoys.blogspot.com
Twitter:	http://twitter.com/tostka

Additional Credits: [REFERENCE]
Website:	[URL]
Twitter:	[URL]

Change Log
  vers: 8:28 AM 9/24/2013 - functional, from the get-db-freespace-report.ps1. Added '*Unable to connect*' error testing

-----------
Code used within main body [declarations etc]
$SMTPFrom = (($ScriptBaseName.replace(".","-")) + "@unisys.com") ; 
#$SMTPSubj= ("Daily Rpt: "+ (Split-Path $Attachment -Leaf) + " " + [System.DateTime]::Now)
# putting a crlf into the msgbody use 'n (equals CrLf)
#$objMailMessage.Body = "Hello `nThis is a second line."
#$SMTPBody="Pass Completed "+ [System.DateTime]::Now + "`nResults Attached: " +$Attachment
$SMTPServer = "NA-MAILRELAY-T3.na.uis.unisys.com" ; 
#$SMTPTo="ExchangeTaskTeam@unisys.com"
$SmtpToAdmin="todd.kadrie@unisys.com" ; 
$SMTPTo=$SmtpToAdmin ; 
$SmtpToNotify="Tom.Gunstad@chrobinson.com" ; 
$SMTPTo2=$SmtpToNotify ; 
# setup body as a hash
$SmtpBody = @()
# to accumulate the body as a log of outputs: 
$SmtpBody += 'Group Share'
# add dash divider 
$SmtpBody += ('-' * 50)
# -----------
  
    .DESCRIPTION
Send-EmailNotif() - Generic send-mailmessage wrapper function
    
    .PARAMETER  SMTPFrom
SMTP From address (alias "from")

    .PARAMETER  SmtpTo
SMTP To address (alias "to")

    .PARAMETER  SMTPSubj
SMTP Subject (alias "subj")

    .PARAMETER  SMTPServer
SMTP Server (alias "server")

    .PARAMETER  SmtpBody
Smtp Body text (alias "body")

    .PARAMETER  Attachment
SMTP File Attachment (alias "attach")

    .INPUTS
None. Does not accepted piped input.

    .OUTPUTS
None. Returns no objects or output.

    .EXAMPLE
Send-EmailNotif 
*----------^ END Comment-based Help  ^---------- #>

  <#
  PARAM(
    [parameter(Mandatory=$true)]
    [alias("from")]
    [string] $SMTPFrom,
    [parameter(Mandatory=$true)]
    [alias("to")]
    [string] $SmtpTo,
    [parameter(Mandatory=$true)]
    [alias("subj")]
    [string] $SMTPSubj,
    [parameter(Mandatory=$true)]
    [alias("server")]
    [string] $SMTPServer = "NA-MAILRELAY-T3.na.uis.unisys.com" ,
    [parameter(Mandatory=$true)]
    [alias("body")]
    [string] $SmtpBody,
    [parameter(Mandatory=$false)]
    [alias("attach")]
    [string] $Attachment 
  )
  #>

  <#
  # 9:49 AM 9/24/2013 save time and don't bother trying to mail from my laptop - 25 is blocked
  if ($bDebug) {
    write-host -foregroundcolor yellow "Funct: Send-EmailNotif()..."
    write-host -foregroundcolor darkgray "`$ComputerName: " $ComputerName
  }
  
  
  if ($ComputerName -ne 'Server$3-KADRIETS2') {
    #write-warning "no match" ; write-host " "
    
    # before you email conv to str & add CrLf:
    $SmtpBody = $SmtpBody | out-string
      
    # define/update variables into $email splat for params
    if ($Attachment -AND (test-path $Attachment)) {
      # attachment send
      $email = @{
      From = $SMTPFrom
      To = $SmtpToAdmin
      Subject = $SMTPSubj
      SMTPServer = $SMTPServer
      Body = $SmtpBody
      Attachments = $Attachment
      } 
    } else {
      # no attachment
      $email = @{
      From = $SMTPFrom
      To = $SmtpToAdmin
      Subject = $SMTPSubj
      SMTPServer = $SMTPServer
      Body = $SmtpBody
      } 
    }
    
    # 11:59 AM 8/28/2013 mailing debugging code
    # write-host  -ForegroundColor Yellow "Emailing with following parameters:"
    # $email
    # write-host "body:"
    # $SmtpBody
    # write-host ("-" * 5)
    # write-host "body.length: " $SmtpBody.length
    write-host ((get-date).ToString("HH:mm:ss") + ": sending mail...")
    $error.clear() 
    send-mailmessage @email 

    # then pipe just the errors out to console
    if($error.count -gt 0){
      write-host -foregroundcolor red ((get-date).ToString("HH:mm:ss") + ": " + $error )
      if ($error -like '*Unable to connect*') {write-host -foregroundcolor yellow ((get-date).ToString("HH:mm:ss") + ": ***" + $SMTPServer + " is unreachable on port 25.`nMake sure " + $ComputerName + " can telnet the host on port 25!***") }
    }
    
  } else {
    write-warning ($ComputerName + " is blocked for SMTP access, skipping email")
  }  # if-block end
  #>
#}#*----------------^ END Function Send-EmailNotif ^---------------- ;
#>

#--------------------------- Invocation of Get-TLogSpaceRegion ---------------------------
	# spec version:8|14) and Sitename:*Server$3*|*AUSYD*|*GBMK* 
	If ($bDebug -eq $TRUE) {write-host Get-TLogSpaceRegion -ExTargVers $ExTargVers -SiteName $SiteName}
	Get-TLogSpaceRegion -ExTargVers $ExTargVers -SiteName $SiteName
	# Turn off Tracing
	#Set-PSDebug -Trace 0
#*================^ END SCRIPT BODY ^================
 


