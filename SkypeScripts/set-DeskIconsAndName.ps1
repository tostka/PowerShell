# set-desktopicons.ps1
# this enables the Mycomputer & Network icons on the desktop and renames the Mycomputer to the $env:COMPUTERNAME

# vers: 10:47 AM 6/23/2014
# code bits from: http://blogs.technet.com/b/heyscriptingguy/archive/2012/06/09/weekend-scripter-use-powershell-to-change-computer-icon-caption-to-computer-name.aspx
#   and from:
#  and ref from http://blog.danovich.com.au/2010/02/18/add-my-computer-to-desktop-and-change-to-computer-name/ 
# and http://hlouwers.wordpress.com/2010/07/24/show-hide-desktop-items-windows-2008-r2-windows-7-by-means-of-registry-and-microsoft-group-policy-preferences/
 
# on the setValue, 0=show the icon, 1=hide it
$ComputerName = ($env:computername) ;
$Hive = "CurrentUser" ;
$Key = "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" ;
$MyComputer = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" ;
$UserFiles = "{59031a47-3f72-44a7-89c5-5595fe6b30ee}" ;
$Network = "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}";
$RecycleBin="{645FF040-5081-101B-9F08-00AA002F954E}";
$Kind = [Microsoft.Win32.RegistryValueKind] ;
$RegHive = [Microsoft.Win32.RegistryHive]$hive ;
$RegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive,$ComputerName) ;

$MyComputerValue = ($RegKey.OpenSubKey($Key)).GetValue($MyComputer) ;
$UserFilesValue = ($RegKey.OpenSubKey($Key)).GetValue($UserFiles) ;
$NetworkValue  = ($RegKey.OpenSubKey($Key)).GetValue($Network) ;

if ($MyComputerValue -eq $null -or $MyComputerValue -eq 1) 
{
	$Computer = $regKey.OpenSubKey($Key,$true) ;
	$Computer.SetValue($MyComputer, 0,$Kind::DWord) ;
} ;

if ($NetworkValue -eq $null -or $UserFilesValue -eq 1) 
{
	$User = $regKey.OpenSubKey($Key,$true) ;
	$User.SetValue($Network, 0,$Kind::DWord) ;
} ;

# finally update the Mycomputer title:  computer namespace is 17
$My_Computer = 17

$Shell = new-object -comobject shell.application

$NSComputer = $Shell.Namespace($My_Computer)

$NSComputer.self.name = $env:COMPUTERNAME
