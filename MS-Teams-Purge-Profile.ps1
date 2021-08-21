<#
.DESCRIPTION
	Purge inactive local MS Teams profile data.  Comments and suggestions always welcome.
.EXAMPLE
	.\MS-Teams-Purge-Profile.ps1
	
.NOTES  
	File Name:  MS-Teams-Purge-Profile.ps1
	Author   :  Jeff Jones
	Version  :  1.0
	Modified :  2021-08-21
#>

function Main() {
    # Discover all local users https://stackoverflow.com/questions/9725521/how-to-get-the-parents-parent-directory-in-powershell
    $profile    = $env:USERPROFILE
    $root       = (Get-ChildItem $profile)[0].Parent.Parent.FullName
    $allUsers   = $root | Get-ChildItem -Directory -Exclude "Public"

    # Compare to Active Directory https://techibee.com/active-directory/powershell-search-for-a-user-without-using-ad-module/2872
    foreach ($u in $allUsers) {
        # Define scope
        $name       = $u.Name
        $ldap       = "(&(ObjectCategory=Person)(ObjectClass=User)(SamAccountName=" + $name + "))"
        $search     = [adsisearcher]$ldap
        $results    = $search.FindAll()

        # Search Active Directory
        if ($results.Count -eq 0) {
            Write-Host "NOT FOUND [" + $name + "] IN AD.  DELETE FOLDER." -ForegroundColor "Yellow"
            Remove-Item "$root\$name" -Force
        } else {
            Write-Host "FOUND [" + $name + "] IN AD" -ForegroundColor "Green"
        }
    }
}

# Open Log
$prefix = $MyInvocation.MyCommand.Name
$host.UI.RawUI.WindowTitle = $prefix
$stamp = Get-Date -UFormat "%Y-%m-%d-%H-%M-%S"
Start-Transcript "$PSScriptRoot\log\$prefix-$stamp.log"
$start = Get-Date

Main

# Close Log
$end = Get-Date
$totaltime = $end - $start
Write-Host "`nTime Elapsed: $($totaltime.tostring("hh\:mm\:ss"))"
Stop-Transcript