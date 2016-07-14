<#
.SYNOPSIS
	SharePoint Central Admin - View active services across entire farm. No more select machine drop down dance!
.DESCRIPTION
	Create Desktop Icon to launch Office 365 with securely saved credentials

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Name		: office365-desktop-icon.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.09
	Last Modified	: 07-14-2016
.LINK
	Source Code
		http://www.github.com/spjeff/o365/office365-desktop-icon.ps1
	
	Download PowerShell Plugins
		* SPO - SharePoint Online
		https://www.microsoft.com/en-us/download/details.aspx?id=35588
		
		* PNP - Patterns and Practices
		https://github.com/officedev/pnp-powershell
#>

# input
Write-Host "=== Make Office 365 PowerShell desktop icon  ==="
$url = Read-Host "Tenant - Admin URL"
$user = Read-Host "Tenant - Username"
$pw = Read-Host "Tenant - Password" -AsSecureString

# save to registry
$hash = $pw | ConvertFrom-SecureString

# command
"`$h = ""$hash""`n`$secpw = ConvertTo-SecureString -String `$h`n`$c = New-Object System.Management.Automation.PSCredential (""$user"", `$secpw)`nImport-Module -WarningAction SilentlyContinue Microsoft.Online.SharePoint.PowerShell -Prefix MS -ErrorAction SilentlyContinue`nImport-Module -WarningAction SilentlyContinue SharePointPnPPowerShellOnline -Prefix PNP -ErrorAction SilentlyContinue`nConnect-MSSPOService -URL $url -Credential `$c`n`$firstUrl = (Get-MSSPOSite)[0].Url`n`$pnp = gcm Connect-PNPSPOnline -ErrorAction SilentlyContinue`n`$pnpurl = ""https://github.com/OfficeDev/PnP-PowerShell""`nif (`$pnp) {`nConnect-PNPSPOnline -URL `$firstUrl -Credential `$c`n} else {`nWrite-Warning ""Missing PNP cmds. Download at $pnpurl""`nstart $pnpurl`n}`nGet-MSSPOSite`n" | Out-File "$home\o365-icon.ps1"

# create desktop shortcut
$folder = [Environment]::GetFolderPath("Desktop")
$TargetFile = "c:\Windows\System32\cmd.exe"
$ShortcutFile = "$folder\Office365.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.Arguments = "/c ""start powershell -noexit """"$home\o365-icon.ps1"""""""
$Shortcut.IconLocation = "powershell.exe, 0";
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()