# set-AdminBG.ps1

<# 
.SYNOPSIS
set-AdminBG.ps1 - Script that creates desktop wallpaper with specified text overlaid over specified image or background color (PS-native Bginfo.exe alternative)
.NOTES
Reworked by:  Todd Kadrie
Website:  https://tinstoys.blogspot.com/
Based on code By: _Emin_    
Website:	https://p0w3rsh3ll.wordpress.com/2014/08/29/poc-tatoo-the-background-of-your-virtual-machines/
Change Log
# 11:19 AM 6/29/2016 updated docs: have to import the tasks under each UID, or they don't execute on logon
# 10:34 AM 6/29/2016 updated help distribution docs and cmds to add/remove schedtasks
# 1:46 PM 6/28/2016 add explit exit (SchedTask still in Running otherwise)
# 1:34 PM 6/28/2016 pull the %p, %H:%M are in 24h format
# 8:51 AM 6/28/2016 fixed ampm -uformat
# 11:14 AM 6/27/2016: added get-LocalDiskFreeSpace, local-only version (for BGInfo) drops server specs and reporting, and sorts on Name/driveletter
# 1:43 PM 6/27/2016 ln159 psv2 is crapping out here, Primary needs to be tested $primary -eq $true for psv2
# 12:29 PM 6/27/2016 params Psv2 Mandatory requires =$true
# 12:21 PM 6/27/2016 submain: BGInfo: switch font to courier new
# 11:27 AM 6/27/2016  submain: switched AMPM fmt to T
# 11:24 AM 6/27/2016  submain: added | out-string | out-default to the drive info
# 11:23 AM 6/27/2016 submain: added timestamp and drivespace report
* 11:00 AM 6/27/2016 extended to accommodate & detect and redadmin the AdminAcct2 acct as well
* 10:56 AM 6/27/2016 reflects additions (Current theme)from cemaphore's comments & sample @ http://pastebin.com/Fva47UKT
	along with the Red Admin Theme I added, and code to detect AdminAcct 
# 10:48 AM 6/27/2016 tweak the uptime fmt:
* 9:12 AM 6/27/2016 TSK reformatted, added pshelp
* September 5, 2014 - _Emin_'s posted version
# DISTRIBUTION COMMANDS:
This can be configured as a Scheduled Task, to load on all logons. But this would not provide per-admin choice on the topic. 
1) distribute .ps1, .lnk & .xml to all Lync FE's, SSRS & SQL Nodes (to c:\scripts) 1-line:
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$files = (gci -path "\\tsclient\c\usr\work\lync\scripts\set-AdminBG*" | ?{$_.Name -match "(?i:(.*\.(PS1|CMD|LNK|XML)))$" }) ;$L13ProdALL="Server0;Server1;Server2" ; $L13ProdALL=$L13ProdALL.split(";") ; $L13ProdALL | foreach {  write-host -fore yell "copying $($files) to $_" ; copy-item -path $files �destination \\$_\c$\scripts\ -whatif ; } ;
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
2a) For a manual Startup menu launch: copy the shortcut .lnk to startup folder for Account & AdminAcct 1-line
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$tUsers="Account","AdminAcct";$tFldr="C:\Users\XULOGONX\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" ;$files = get-childitem $($tFldr.replace("XULOGONX",$($env:username))) ; $L13ProdALL="Server0;Server1;Server2" ; $L13ProdALL=$L13ProdALL.split(";") ; $L13ProdALL | foreach {  write-host -fore yell "copying $($files) to $_" ;  foreach($usr in $tUsers){    copy-item -path $files �destination "$($tFldr.replace("XULOGONX",$usr))\" -whatif ;  } ; } ;
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
2b) Alternatively, you can use a ScheduledTask to launch on logon (multiple user's trigger Account|AdminAcct:  Export the current SchedTask to xml: (1-line):
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Get-ScheduledTask set-AdminBG.ps1* | foreach {  Export-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath | Out-File -Filepath (Join-Path -path "c:\scripts\" -childpath "$($_.TaskName).xml") �WhatIf ;} ; 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
And then remote import via �CimSession
# note, multi logon triggers only work on the logon that 'created' the entry. So you need to push them under their own sessions!
#-=-=-=-=-=-=Explicit AdminAcct (run from AdminAcct ps)-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$srcTask="c:\scripts\set-AdminBG.ps1 (Admin Informational Wallpaper)-AdminAcct.xml" ;$file =  get-childitem -path $srcTask ;$L13ProdALL="Server0;Server1;Server2".split(";") ;$L13ProdALL | foreach {  write-host -fore yell "importing SchedTask $($file) to $_" ;  copy-item -path $file �destination \\$_\c$\scripts\ -whatif ;   if(test-path "\\$_\c$\scripts\$($file.Name)") {    Register-ScheduledTask -CimSession $_ -Xml (get-content "$file" | out-string) -TaskName $((split-path $srcTask -leaf).replace(".xml","")) -User DOMAIN\AdminAcct �Force ;    } else { write-warning "$((get-date).ToString("HH:mm:ss")):Missing src xml at far end!" } ;} ; 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#-=-=-=-=-=-=Explicit Account (run from Account ps)-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$srcTask="c:\scripts\set-AdminBG.ps1 (Admin Informational Wallpaper)-kadrtiss.xml" ;$file =  get-childitem -path $srcTask ;$L13ProdALL="Server0;Server1;Server2".split(";") ;$L13ProdALL | foreach {  write-host -fore yell "importing SchedTask $($file) to $_" ;  copy-item -path $file �destination \\$_\c$\scripts\ -whatif ;   if(test-path "\\$_\c$\scripts\$($file.Name)") {    Register-ScheduledTask -CimSession $_ -Xml (get-content "$file" | out-string) -TaskName $((split-path $srcTask -leaf).replace(".xml","")) -User DOMAIN\Account �Force ;    } else { write-warning "$((get-date).ToString("HH:mm:ss")):Missing src xml at far end!" } ;} ; 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# to bulk remove the above:
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
$L13ProdALL="Server0;Server1;Server2".split(";") ; $tST="*set-AdminBG.ps1*" ;$L13ProdALL | foreach {write-host -fore yell "Removing SchedTask $tST from $_" ;Get-ScheduledTask -cimsession $($_) |?{$_.taskname -like $tST} | Unregister-ScheduledTask -whatif ;} ;
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
.DESCRIPTION
New-BGinfo - Create desktop wallpaper with specified text overlaid over specified image or background color (PS Bginfo.exe alternative)
Script wrapping _Emin_'s script functions to generate BGInfo style desktop wallpaper, from native Powershell.
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
Creates and assigns a status info wallpaper (solid color) 
.EXAMPLE
Powershell.exe -noprofile -command "& {c:\scripts\set-AdminBG.ps1 }" ; 
To launch on startup: Put above into C:\Users\LOGON\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\AdminBG.lnk file
.LINK
https://p0w3rsh3ll.wordpress.com/2014/08/29/poc-tatoo-the-background-of-your-virtual-machines/
#>

# 11:01 AM 6/27/2016: domain\ admin shared id regex
$rgxAdminAcct="^((DOMAIN|DOMAIN-LAB)\\(AdminAcct2|AdminAcct))$" ; 

#*======v FUNCTIONS v======

#*------v Function New-BGinfo v------
Function New-BGinfo {
    <# 
    .SYNOPSIS
    New-BGinfo - Create desktop wallpaper with specified text overlaid over specified image or background color (PS Bginfo.exe alternative)
    .NOTES
    Written By: _Emin_
    Updated, added psv2 compat, extended, by Todd Kadrie
    Website:	https://p0w3rsh3ll.wordpress.com/2014/08/29/poc-tatoo-the-background-of-your-virtual-machines/
    Change Log
    # # 8:51 AM 6/28/2016 fixed ampm -uformat
    # 11:14 AM 6/27/2016: added get-LocalDiskFreeSpace, local-only version (for BGInfo) drops server specs and reporting, and sorts on Name/driveletter
    # 1:43 PM 6/27/2016 ln159 psv2 is crapping out here, Primary needs to be tested $primary -eq $true for psv2
    # 12:29 PM 6/27/2016 params Psv2 Mandatory requires =$true
    # 12:21 PM 6/27/2016 submain: BGInfo: switch font to courier new
    # 11:27 AM 6/27/2016  submain: switched AMPM fmt to T
    # 11:24 AM 6/27/2016  submain: added | out-string | out-default to the drive info
    # 11:23 AM 6/27/2016 submain: added timestamp and drivespace report
    * 11:00 AM 6/27/2016 extended to accommodate & detect and redadmin the AdminAcct2 acct as well
    * 10:56 AM 6/27/2016 reflects additions (Current theme)from cemaphore's comments & sample @ http://pastebin.com/Fva47UKT
		along with the Red Admin Theme I added, and code to detect AdminAcct 
		# 10:48 AM 6/27/2016 tweak the uptime fmt:
    * 9:12 AM 6/27/2016 TSK reformatted, added pshelp
    * September 5, 2014 - posted version
    .DESCRIPTION
    New-BGinfo - Create desktop wallpaper with specified text overlaid over specified image or background color (PS Bginfo.exe alternative)
    .PARAMETER  Text
    Text to be overlayed over specified background
    .PARAMETER  OutFile
    Output file to be created (and then assigned separately to the desktop). Defaults to c:\temp\BGInfo.bmp
    .PARAMETER  Align
    Text alignment [Left|Center]
    .PARAMETER  Theme
    Desktop Color theme (defaults Blue)
    .PARAMETER  FontName
    Text Font Name (Defaults Arial) [-FontName Arial]
    .PARAMETER  FontSize
    Integer Text Font Size (Defaults 12 point) [9-45]
    .PARAMETER  UseCurrentWallpaperAsSource
    Switch Param that specifies to recycle existing wallpaper [-UseCurrentWallpaperAsSource]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Powershell.exe -noprofile -command "& {c:\scripts\set-AdminBG.ps1 }" ; 
    To launch on startup: Put above into C:\Users\LOGON\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\AdminBG.lnk file
    .LINK
    https://p0w3rsh3ll.wordpress.com/2014/08/29/poc-tatoo-the-background-of-your-virtual-machines/
    #>
  
    # 9:56 AM 6/27/2016 add cmaphore's Current Theme to the Theme list validator & default to Current (v Black)
	# 10:33 AM 6/27/2016 add Red admin theme
	# 11:26 AM 6/27/2016 switched fontsize default from 12 to 10
	# 12:29 PM 6/27/2016 params Psv2 Mandatory requires =$true

    Param(
            [Parameter(Mandatory=$true)]
            [string] $Text
            ,[Parameter()]
            [string] $OutFile= "$($($env:temp))\BGInfo.bmp"
            ,[Parameter()]
            [ValidateSet("Left","Center")]
            [string]$Align="Center"
            ,[Parameter()]
            [ValidateSet("Current","Blue","Grey","Black","Red")]
            [string]$Theme="Current"
            ,[Parameter()]
            [string]$FontName="Arial"
            ,[Parameter()]
            [ValidateRange(9,45)]
            [int32]$FontSize = 8
            ,[Parameter()]
            [switch]$UseCurrentWallpaperAsSource
    ) ; 
    
    Begin {
        # 9:59 AM 6/27/2016 add cmaphore's detection of Current Theme
        # Enumerate current wallpaper now, so we can decide whether it's a solid colour or not
        try {
            $wpath = (Get-ItemProperty 'HKCU:\Control Panel\Desktop' -Name WallPaper -ErrorAction Stop).WallPaper
            if ($wpath.Length -eq 0) {
                # Solid colour used
                $UseCurrentWallpaperAsSource = $false ; 
                $Theme = "Current" ; 
            } ; 
        } catch {
            $UseCurrentWallpaperAsSource = $false ; 
            $Theme = "Current" ; 
        } ; 
    
        Switch ($Theme) {
            # 9:42 AM 6/27/2016 add cmaphore's idea of a 'Current' theme switch case, pulling current background color $RGB, and defaulting if not set
            Current {
                $RGB = (Get-ItemProperty 'HKCU:\Control Panel\Colors' -ErrorAction Stop).BackGround ; 
                if ($RGB.Length -eq 0) {
                    $Theme = "Black" ; # Default to Black and don't break the switch
                } else {
                    $BG = $RGB -split " " ; 
                    $FC1 = $FC2 = @(255,255,255) ; 
                    $FS1=$FS2=$FontSize ; 
                    break ; 
                } ; 
            } ; 
            Blue {
                $BG = @(58,110,165) ; 
                $FC1 = @(254,253,254) ; 
                $FC2 = @(185,190,188) ; 
                $FS1 = $FontSize+1 ; 
                $FS2 = $FontSize-2 ; 
                break ; 
            } ; 
            Grey {
                $BG = @(77,77,77) ; 
                $FC1 = $FC2 = @(255,255,255) ; 
                $FS1=$FS2=$FontSize ; 
                break ; 
            } ; 
            Black {
                $BG = @(0,0,0) ; 
                $FC1 = $FC2 = @(255,255,255) ; 
                $FS1=$FS2=$FontSize ; 
            } ; 
			# 10:30 AM 6/27/2016 add a red theme to mark shared admin accounts
			Red {
                $BG = @(184,40,50) ; 
                $FC1 = $FC2 = @(255,255,255) ; 
                $FS1=$FS2=$FontSize ; 
            } ; 
        } ;  # swtch-E
        
        Try {
            [system.reflection.assembly]::loadWithPartialName('system.drawing.imaging') | out-null ; 
            [system.reflection.assembly]::loadWithPartialName('system.windows.forms') | out-null ; 
            # Draw string > alignement
            $sFormat = new-object system.drawing.stringformat
            Switch ($Align) {
                Center {
                    $sFormat.Alignment = [system.drawing.StringAlignment]::Center ; 
                    $sFormat.LineAlignment = [system.drawing.StringAlignment]::Center ; 
                    break ; 
                } ; 
                Left {
                    $sFormat.Alignment = [system.drawing.StringAlignment]::Center ; 
                    $sFormat.LineAlignment = [system.drawing.StringAlignment]::Near ; 
                } ; 
            } ;  # swtch-E
     
            if ($UseCurrentWallpaperAsSource) {
                # 10:01 AM 6/27/2016 moved $wppath to top of begin
                if (Test-Path -Path $wpath -PathType Leaf) {
                    $bmp = new-object system.drawing.bitmap -ArgumentList $wpath ; 
                    $image = [System.Drawing.Graphics]::FromImage($bmp) ; 
                    $SR = $bmp | Select Width,Height ; 
                } else {
                    Write-Warning -Message "Failed cannot find the current wallpaper $($wpath)" ; 
                    break ; 
                } ; 
            } else {
                # 1:43 PM 6/27/2016 psv2 is crapping out here, Primary needs to be tested $primary -eq $true for psv2
                #$SR = [System.Windows.Forms.Screen]::AllScreens | Where Primary | Select -ExpandProperty Bounds | Select Width,Height ; 
                $SR = [System.Windows.Forms.Screen]::AllScreens |?{$_.Primary} | Select -ExpandProperty Bounds | Select Width,Height ; 
                #}
                Write-Verbose -Message "Screen resolution is set to $($SR.Width)x$($SR.Height)" -Verbose ; 
     
                # Create Bitmap
                $bmp = new-object system.drawing.bitmap($SR.Width,$SR.Height) ; 
                $image = [System.Drawing.Graphics]::FromImage($bmp) ; 
         
                $image.FillRectangle(
                    (New-Object Drawing.SolidBrush (
                        [System.Drawing.Color]::FromArgb($BG[0],$BG[1],$BG[2]) 
                    )),
                    (new-object system.drawing.rectanglef(0,0,($SR.Width),($SR.Height))) 
                ) ; 
            } ; 
            
        } Catch {
            Write-Warning -Message "Failed to $($_.Exception.Message)" ; 
            break ; 
        } ; 
    } ;  # BEG-E
    
    Process {
        # Split our string as it can be multiline
        $artext = ($text -split "\r\n") ; 
        $i = 1 ; 
        Try {
            for ($i ; $i -le $artext.Count ; $i++) {
                if ($i -eq 1) {
                    $font1 = New-Object System.Drawing.Font($FontName,$FS1,[System.Drawing.FontStyle]::Bold) ; 
                    $Brush1 = New-Object Drawing.SolidBrush (
                        [System.Drawing.Color]::FromArgb($FC1[0],$FC1[1],$FC1[2]) 
                    ) ; 
                    $sz1 = [system.windows.forms.textrenderer]::MeasureText($artext[$i-1], $font1) ; 
                    $rect1 = New-Object System.Drawing.RectangleF (0,($sz1.Height),$SR.Width,$SR.Height) ; 
                    $image.DrawString($artext[$i-1], $font1, $brush1, $rect1, $sFormat) ; 
                } else {
                    $font2 = New-Object System.Drawing.Font($FontName,$FS2,[System.Drawing.FontStyle]::Bold) ; 
                    $Brush2 = New-Object Drawing.SolidBrush (
                        [System.Drawing.Color]::FromArgb($FC2[0],$FC2[1],$FC2[2]) 
                    ) ; 
                    $sz2 = [system.windows.forms.textrenderer]::MeasureText($artext[$i-1], $font2) ; 
                    $rect2 = New-Object System.Drawing.RectangleF (0,($i*$FontSize*2 + $sz2.Height),$SR.Width,$SR.Height) ; 
                    $image.DrawString($artext[$i-1], $font2, $brush2, $rect2, $sFormat) ; 
                } ; 
            } ;  # loop-E
            
        } Catch {
            Write-Warning -Message "Failed to $($_.Exception.Message)" ; 
            break ; 
        } ; 
        
    } ;  # PROC-E
    
    End {  
        Try {
            # Close Graphics
            $image.Dispose(); ; 
     
            # Save and close Bitmap
            $bmp.Save($OutFile, [system.drawing.imaging.imageformat]::Bmp); ; 
            $bmp.Dispose() ;      
            # Output our file path into the pipeline
            Get-Item -Path $OutFile ; 
        } Catch {
            Write-Warning -Message "Failed to $($_.Exception.Message)" ; 
            break ; 
        } ; 
    } ;  # END-E
} #*------^ END Function New-BGinfo ^------

#*------v Function Set-Wallpaper v------
Function Set-Wallpaper {
    <# 
    .SYNOPSIS
    Set-Wallpaper - Set specified file as desktop wallpaper
    .NOTES
    Written By: _Emin_
    Tweaked/Updated by: Todd Kadrie
    Website:	https://p0w3rsh3ll.wordpress.com/2014/08/29/poc-tatoo-the-background-of-your-virtual-machines/
    Change Log
    * 9:12 AM 6/27/2016 TSK reformatted & added pshelp
    * September 5, 2014 - posted version
    .DESCRIPTION
    .PARAMETER  Path
    Path to image to be set as desktop background
    .PARAMETER  Style
    Style to apply to wallpaper [Center|Stretch|Fill|Tile|Fit]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE

    .LINK
    https://p0w3rsh3ll.wordpress.com/2014/08/29/poc-tatoo-the-background-of-your-virtual-machines/
    #>
    Param(
        [Parameter(Mandatory=$true)]
        $Path
        ,[ValidateSet('Center','Stretch','Fill','Tile','Fit')]
        $Style = 'Stretch' 
    ) ; 
    
    Try {
        if (-not ([System.Management.Automation.PSTypeName]'Wallpaper.Setter').Type) {
            Add-Type -TypeDefinition @"
           using System;
            using System.Runtime.InteropServices;
            using Microsoft.Win32;
            namespace Wallpaper {
                public enum Style : int {
                Center, Stretch, Fill, Fit, Tile
                }
                public class Setter {
                    public const int SetDesktopWallpaper = 20;
                    public const int UpdateIniFile = 0x01;
                    public const int SendWinIniChange = 0x02;
                    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
                    private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
                    public static void SetWallpaper ( string path, Wallpaper.Style style ) {
                        SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
                        RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
                        switch( style ) {
                            case Style.Tile :
                                key.SetValue(@"WallpaperStyle", "0") ;
                                key.SetValue(@"TileWallpaper", "1") ;
                                break;
                            case Style.Center :
                                key.SetValue(@"WallpaperStyle", "0") ;
                                key.SetValue(@"TileWallpaper", "0") ;
                                break;
                            case Style.Stretch :
                                key.SetValue(@"WallpaperStyle", "2") ;
                                key.SetValue(@"TileWallpaper", "0") ;
                                break;
                            case Style.Fill :
                                key.SetValue(@"WallpaperStyle", "10") ;
                                key.SetValue(@"TileWallpaper", "0") ;
                                break;
                            case Style.Fit :
                                key.SetValue(@"WallpaperStyle", "6") ;
                                key.SetValue(@"TileWallpaper", "0") ;
                                break;
}
                        key.Close();
                    }
                }
            }
"@ -ErrorAction Stop ; 
            } else {
                Write-Verbose -Message "Type already loaded" -Verbose ; 
            } ; 
        # } Catch TYPE_ALREADY_EXISTS
        } Catch {
            Write-Warning -Message "Failed because $($_.Exception.Message)" ; 
        } ; 
     
    [Wallpaper.Setter]::SetWallpaper( $Path, $Style ) ; 
} ; #*------^ END Function Set-Wallpaper ^------

#*------v Function get-LocalDiskFreeSpace v------
function get-LocalDiskFreeSpace {
	# 11:14 AM 6/27/2016 This local-only version (for BGInfo) drops server specs and reporting, and sorts on Name/driveletter
	#Get-WmiObject Win32_Volume -filter "drivetype = 3" | Select-Object Name, @{Name="Size(GB)";Expression={"{0:N1}" -f($_.Capacity/1gb)}},@{Name="FreeSpace(GB)";Expression={"{0:N1}" -f($_.freespace/1gb)}},@{Name="FreeSpacePerCent";Expression={"{0:P0}" -f($_.freespace/$_.capacity)}} |Sort-Object -property "Name" | Format-Table -auto ; 
	# 12:06 PM 6/27/2016try to elipses the namespace @18 chars
	# 12:24 PM 6/27/2016 shortened col titles
	Get-WmiObject Win32_Volume -filter "drivetype = 3" | 
        Select-Object @{Name="Vol";Expression={if($_.Name.tostring().length -gt 18){"$($_.Name.tostring().substring(0,18))..."} else {"$($_.Name.tostring())" }} },@{Name="Size(gb)";Expression={"{0:N1}" -f($_.Capacity/1gb)}},@{Name="Free(gb)";Expression={"{0:N1}" -f($_.freespace/1gb)}},@{Name="Free`%";Expression={"{0:P0}" -f($_.freespace/$_.capacity)}} | 
            Sort-Object -property "Name" | Format-Table -auto ; 
	
} ; #*------^ END Function get-DiskFreeSpace ^------
#*======^ END FUNCTIONS ^======

#*------v Function Sub Main v------

# Typical Call
# Gather data & build the Overlay Text into a text here-string variable
# 12:38 PM 6/27/2016 get-ciminstance was only introduced with Psv3, doesn't work with psv2
<#
if($host.version.major -gt 2){
	$os = Get-CimInstance Win32_OperatingSystem ;
} else { 
    $os = Get-WmiObject �class Win32_OperatingSystem �computername $($env:COMPUTERNAME) �erroraction Stop ; 
} ; 
#>
$os = Get-WmiObject �class Win32_OperatingSystem �computername $($env:COMPUTERNAME) �erroraction Stop
<#($o = [pscustomobject]@{
		HostName =  $env:COMPUTERNAME ;
		UserName = '{0}\{1}' -f  $env:USERDOMAIN,$env:USERNAME ;
		'Operating System' = '{0} Service Pack {1} (build {2})' -f  $os.Caption,
		$os.ServicePackMajorVersion,$os.BuildNumber ; 
	}) | ft -AutoSize ;
#>
<# 1:27 PM 6/27/2016 psv2 is throwing: 
WARNING: Failed to Cannot bind parameter 'FilterScript'. Cannot convert the "Primary" value of type "System.String" to
type "System.Management.Automation.ScriptBlock".
#>
# try rewriting into 2-stage cobj
$oHash=@{
    HostName =  $env:COMPUTERNAME ;
	UserName = '{0}\{1}' -f  $env:USERDOMAIN,$env:USERNAME ;
	'Operating System' = '{0} Service Pack {1} (build {2})' -f  $os.Caption,
	$os.ServicePackMajorVersion,$os.BuildNumber ; 
} ; 
$o=New-Object -TypeName PSObject -Property $oHash ; 
$o | format-table -autosize ; 	
# 10:48 AM 6/27/2016 tweak the uptime fmt: $ts = (New-TimeSpan -Start $os.LastBootUpTime -End (Get-Date)) ; 
#$ts = (New-TimeSpan -Start $($os.LastBootUpTime) -End (Get-Date)) ;
# use WMI converttodatetime() to make psv2 compat
$ts = (New-TimeSpan -Start $($os.ConvertToDateTime($os.LastBootUpTime)) -End (Get-Date)) ; 

$BootTime = "$($ts.days)d:$($ts.hours)h:$($ts.minutes)m:$($ts.seconds)s" ; 

# $t is the multiline text here-string
# 11:23 AM 6/27/2016 added timestamp and drivespace report
# 11:24 AM 6/27/2016 added | out-string | out-default to the drive info
# 11:27 AM 6/27/2016 switched AMPM fmt to T
# 8:51 AM 6/28/2016 fixed ampm -uformat
# 1:34 PM 6/28/2016 pull the %p, %H:%M are in 24h format
$t = @"
Status as of: $(get-date -uformat "%m/%d/%Y %H:%M")
$($o.HostName) ;
Logged on user: $($o.UserName)
$($o.'Operating System')
Uptime: $BootTime
$(get-LocalDiskFreeSpace | out-string )
"@ ; 

<#Example 1: ala Backinfo - Build a color-only wallpaper with the specified $t overlayed, and assign to desiktop
$WallPaper = New-BGinfo -text $t ;
Set-Wallpaper -Path $WallPaper.FullName -Style Center ; 
#>

#Example 2: ala Bginfo:  using a splat, and overlaying on the current wallpaper

# default to Black Theme
# 12:21 PM 6/27/2016 switch font to courier new
$BGInfo = @{
	 Text  = $t ;
	 Theme = "Black" ;
	 FontName = "courier new" ;
	 UseCurrentWallpaperAsSource = $false ;
} ; 

# test for and give AdminAcct a distinctive red background color
#if("$($env:USERDOMAIN)\$($env:USERNAME)" -eq 'DOMAIN\AdminAcct'){
if("$($env:USERDOMAIN)\$($env:USERNAME)" -match $rgxAdminAcct){
	$BGInfo.Theme = "Red" ; 
} ; 

$WallPaper = New-BGinfo @BGInfo ;
Set-Wallpaper -Path $WallPaper.FullName -Style Fill ;
# To Restore the default VM wallpaper
#Set-Wallpaper -Path "C:\Windows\Web\Wallpaper\Windows\img0.jpg" -Style Fill ; 

# 1:46 PM 6/28/2016 add explit exit (SchedTask still in Running otherwise)
exit 0 ; 
