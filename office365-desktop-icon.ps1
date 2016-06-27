<#
Office365 - Desktop Icon

 * make desktop shortcut icon
 * save Tenant URL
 * auto login User & Password
 * secure string hash
 * enumerate all tenant sites
#>

Write-Host "=== Make Office 365 PowerShell desktop icon  ==="

# input
$url = Read-Host "Tenant - Admin URL:"
$user = Read-Host "Tenant - Username:"
$pw = Read-Host "Tenant - Password:" -AsSecureString

# save to registry
$hash = $pw | ConvertFrom-SecureString

# shortcut
"`$h = ""$hash""`n`$secpw = ConvertTo-SecureString -String `$h`n`$c = New-Object System.Management.Automation.PSCredential (""$user"", `$secpw)`nImport-Module -WarningAction SilentlyContinue Microsoft.Online.SharePoint.PowerShell`nConnect-SPOService -URL $url -Credential `$c`nGet-SPOSite" | Out-File "$home\o365-icon.ps1"

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