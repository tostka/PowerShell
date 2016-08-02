# Get-WinEventTail-LyncServerLog.ps1
# Get-WinEventTail.ps1

<# 
    .SYNOPSIS
    Get-WinEventTail-LyncServerLog.ps1 - Tails on a loop, the specified LogName/Provider/Both, for the last -ShowExisting events
    .NOTES
    Written By: Michael Sorens
    Website:	https://stackoverflow.com/questions/15262196/powershell-tail-windows-event-log-is-it-possible
    Change Log
    * 3:18 PM 7/18/2016 added pshelp
    * 3:13 PM 7/18/2016 added splatting, and support for both LogName and Provider params
	* 2:55 PM 7/18/2016 works. switched the -provider spec to -logname, to target all Lync Server log evts.
	Lifted from comment by Michael Sorens
	  .DESCRIPTION
	  Tails on a loop, the specified LogName/Provider/Both, for the last -ShowExisting events
    .PARAMETER  ParaName
    ParaHelpTxt
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    .\Get-WinEventTail-LyncServerLog.ps1
    .LINK
    https://stackoverflow.com/questions/15262196/powershell-tail-windows-event-log-is-it-possible
#>

Set-PSDebug -Strict
# 1:48 PM 7/18/2016 try to splic in -LogName vs -provider
function Get-WinEventTail {
    <# 
    .SYNOPSIS
    Get-WinEventTail-LyncServerLog.ps1 - Tails on a loop, the specified LogName/Provider/Both, for the last -ShowExisting events
    .NOTES
    Written By: Michael Sorens
    Website:	https://stackoverflow.com/questions/15262196/powershell-tail-windows-event-log-is-it-possible
    Change Log
    * 3:13 PM 7/18/2016 added splatting, and support for both LogName and Provider params
    * 2:55 PM 7/18/2016 works. switched the -provider spec to -logname, to target all Lync Server log evts.
    Lifted from comment by Michael Sorens
    .DESCRIPTION
    Tails on a loop, the specified LogName/Provider/Both, for the last -ShowExisting events
    .PARAMETER  LogName
    Applog LogName to be filtered.
    .PARAMETER  Provider
    Applog Provider to be filtered.
    .PARAMETER  ShowExisting
    MaxEvents to be displayed (defaults 10).
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Get-WinEventTail -LogName "Lync Server" ;
    Tail all entries in the Lync Server crimson log
    .LINK
    https://stackoverflow.com/questions/15262196/powershell-tail-windows-event-log-is-it-possible
    #>
    Param(
        [string]$LogName
        ,[string]$Provider
        ,[int]$ShowExisting=10
    )  ; 

    
    if($host.version.major -ge 3){
        $Hash=[ordered]@{
            Dummy = $null ; 
        } ;
    } else {
        # psv2 Ordered obj (can't use with new-object -properites)
        $Hash = New-Object Collections.Specialized.OrderedDictionary ; 
        <# or use an UN-ORDERED psv2 hash:
        $Hash=@{
            Dummy = $null ; 
        } ;
        #>
    } ;
    # then immediately remove the dummy value (blank & null variants too):
    If($Hash.Contains("Dummy")){$Hash.remove("Dummy")} ; 
    # Populate the $hash with fields, post creation 
    if($LogName){$Hash.Add("LogName",$($LogName)) ; } ; 
    if($Provider){$Hash.Add("Provider",$($Provider)) ; } ; 
    #$Hash.Add("MaxEvents",$($ShowExisting)) ; 
    
    if ($ShowExisting -gt 0) {
        #$data = Get-WinEvent -provider $LogName -max $ShowExisting
        # Get-WinEvent -LogName "Lync Server" ;
        #$data = Get-WinEvent -LogName $LogName -max $ShowExisting
        $data = Get-WinEvent @Hash -MaxEvents $($ShowExisting); 
        $data | sort RecordId
        $idx = $data[0].RecordId
    } else {
        #$idx = (Get-WinEvent -provider $LogName -max 1).RecordId
        #$idx = (Get-WinEvent -LogName $LogName -max 1).RecordId
        $idx = (Get-WinEvent Get-WinEvent @Hash -MaxEvents 1).RecordId
    }

    #while ($true) {
    # 1:55 PM 7/18/2016 switch to endless, keeps aborting in mstsc
    for(;;) {
        start-sleep -Seconds 1
        #$idx2  = (Get-WinEvent -LogName $LogName -max 1).RecordId
        $idx2  = (Get-WinEvent @Hash -MaxEvents 1).RecordId
        if ($idx2 -gt $idx) {
            #Get-WinEvent -provider $LogName -max ($idx2 - $idx) | sort RecordId
            #Get-WinEvent -LogName $LogName -max ($idx2 - $idx) | sort RecordId
            #Get-WinEvent -LogName $LogName -MaxEvents ($idx2 - $idx) | sort RecordId
            Get-WinEvent @Hash -MaxEvents ($idx2 - $idx) | sort RecordId
        }
        $idx = $idx2

        # Any key to terminate; does NOT work in PowerShell ISE!
	# 2:54 PM 7/18/2016 rem'd, kept quiting in mstsc
        #if ($Host.UI.RawUI.KeyAvailable) { return; }
    } # loop-E
} #*------^ END Function Get-WinEventTail ^------ ; 

# 2:00 PM 7/18/2016 stick an infinite here - keeps dying
for(;;) {
    Get-WinEventTail -LogName "Lync Server" ;
} ; 
