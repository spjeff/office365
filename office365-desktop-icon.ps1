<#
.SYNOPSIS
	SharePoint Central Admin - View active services across entire farm. No more select machine drop down dance!
.DESCRIPTION
	Create Desktop Icon to launch Office 365 with securely saved credentials

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Name		: office365-desktop-icon.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.05
	Last Modified	: 06-28-2016
.LINK
	Source Code
		http://www.github.com/spjeff/o365/office365-desktop-icon.ps1
	
	Download PowerShell Plugins
		* SPO - SharePoint Online
		https://www.microsoft.com/en-us/download/details.aspx?id=35588
		
		* PNP - Patterns and Practices
		https://github.com/officedev/pnp-powershell
#>

Write-Host "=== Make Office 365 PowerShell desktop icon  ==="

# input
$url = Read-Host "Tenant - Admin URL:"
$user = Read-Host "Tenant - Username:"
$pw = Read-Host "Tenant - Password:" -AsSecureString

# save to registry
$hash = $pw | ConvertFrom-SecureString

# shortcut
"`$h = ""$hash""`n`$secpw = ConvertTo-SecureString -String `$h`n`$c = New-Object System.Management.Automation.PSCredential (""$user"", `$secpw)`nImport-Module -WarningAction SilentlyContinue Microsoft.Online.SharePoint.PowerShell -Prefix MS`nImport-Module -WarningAction SilentlyContinue SharePointPnPPowerShellOnline -Prefix PNP`nConnect-SPOService -URL $url -Credential `$c`nGet-SPOSite" | Out-File "$home\o365-icon.ps1"

# create shortcut
$folder = [Environment]::GetFolderPath("Desktop")
$TargetFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$ShortcutFile = "$folder\Office365.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.Arguments = " -NoExit ""$home\o365-icon.ps1"""
$Shortcut.IconLocation = "powershell.exe, 0";
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()