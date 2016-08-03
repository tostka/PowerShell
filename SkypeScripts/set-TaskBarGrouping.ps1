# set-TaskBarGrouping.ps1

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#<SCIPTFILENAME>.ps1, or #*----------v Function  v----------

<# 
.SYNOPSIS
set-TaskBarGrouping.ps1 - Configure TaskBar grouping settings to CombineWhenTaskbarFull.
.NOTES
Written By: Pat Richard
Website:	http://www.ehloworld.com/2934

Change Log
* 20150218 - posted version

.DESCRIPTION
Configure TaskBar grouping settings to CombineWhenTaskbarFull.
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.\set-TaskBarGrouping.ps1
.LINK
http://www.ehloworld.com/author/pat-richard
*----------^ END Comment-based Help  ^---------- #>

#*------v Function Set-TaskbarGrouping v------
function Set-TaskbarGrouping {

  <# 
  .SYNOPSIS
  Set-TaskbarGrouping() - Configure TaskBar grouping settings.
  .NOTES
  Written By: Pat Richard
  Website:	http://www.ehloworld.com/2934

  Change Log
  * 20150218 - posted version

.NOTES
Note, if not running on Win2012, revise...
  SupportsShouldProcess = $True, SupportsPaging = $True
...to...
SupportsShouldProcess = $True

  .DESCRIPTION
  Configure TaskBar grouping settings.
  .PARAMETER  CombineWhenTaskbarFull
  Switch parameter that causes the TaskBar to begin 'grouping' only when bar is full.
  .PARAMETER  AlwaysCombine
  Switch parameter that causes the TaskBar to always perform task 'grouping'.
  .PARAMETER  NeverCombine
  Switch parameter that causes the TaskBar to never perform 'grouping'.
  .PARAMETER  NoReboot
  Restarts explorer.exe to avoid the need for log-off/restart of system.
  .INPUTS
  None. Does not accepted piped input.
  .OUTPUTS
  None. Returns no objects or output.
  .EXAMPLE
  Set-TaskbarGrouping -NeverCombine
  .LINK
  http://www.ehloworld.com/author/pat-richard
  *----------^ END Comment-based Help  ^---------- #>

  # win2012+-specific version
	#[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True, DefaultParameterSetName = "NeverCombine")]
	
	[CmdletBinding(SupportsShouldProcess = $True, DefaultParameterSetName = "NeverCombine")]
	param(
		# Always combines similar shortcuts into groups
		[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True, ParameterSetName = "AlwaysCombine")]		
		[switch] $AlwaysCombine,
		
		# Combines similar shortcuts into groups only when the taskbar is full
		[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True, ParameterSetName = "CombineWhenTaskbarFull")]
		[switch] $CombineWhenTaskbarFull,
		
		# Never combines similar shortcuts into groups
		[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True, ParameterSetName = "NeverCombine")]
		[switch] $NeverCombine,
		
		# Restarts explorer in order for the grouping setting to immediately take effect. If not specified, the change will take effect after the computer is restarted
		[switch] $NoReboot
	)
	switch ($PsCmdlet.ParameterSetName) {
		"AlwaysCombine" {
			Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name TaskbarGlomLevel -Value 0
		}
		"CombineWhenTaskbarFull" {
			Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name TaskbarGlomLevel -Value 1
		}
		"NeverCombine" {
			Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name TaskbarGlomLevel -Value 2
		}
	}
	if ($NoReboot){
		Stop-Process -ProcessName explorer -force
	}else{
		Write-Verbose "Change will take effect after the computer is restarted"
	}
} #*------^ END Function Set-TaskbarGrouping ^------

Set-TaskbarGrouping -CombineWhenTaskbarFull -NoReboot