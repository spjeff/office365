# Export Office 365 user profiles to CSV (Azure AD & SharePoint UPS)
# Download from ... Microsoft Online Services Sign-in Assistant for IT Professionals RTW http://go.microsoft.com/fwlink/p/?LinkId=286152
# Download from ... Windows Azure Active Directory Module for Windows PowerShell (64-bit version) http://go.microsoft.com/fwlink/p/?linkid=236297
Import-Module MSOnline

# Tenant
$tenantUrl = "https://tenant-admin.sharepoint.com/"
$userName = "admin@tenant.onmicrosoft.com" 
$password = "pass@word1"

# Azure AD (AAD) - All Profiles
Write-Host "Azure AD (AAD)"
$secPassword = ConvertTo-SecureString $password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($userName, $secPassword)
Connect-MsolService -Credential $cred
$allUsers = Get-MsolUser -All
$allUsers | Export-Csv "azure-ad-profiles.csv"

# SharePoint Online (SPO) - All Profiles
Write-Host "SharePoint Online (SPO)"
$start = Get-Date
$i = 0
$total = $allUsers.Count
$coll = @()
$csv = ""
Connect-SPOnline -Url $tenantUrl -Credentials $cred
foreach ($u in $allUsers) {
    # Progress Tracking
    $i++
    $prct = [Math]::Round((($i / $total) * 100.0), 2)
    $elapsed = (Get-Date) - $start
    $remain = ($elapsed.TotalSeconds) / ($prct / 100.0)
    $eta = (Get-Date).AddSeconds($remain)
	
    # Display
    Write-Progress -Activity "Download SharePoint Online user profiles - ETA $eta" -Status "$prct" -PercentComplete $prct
	
    # Append CSV sign in name
    $csv += $u.SignInName + ","
    if ($i -eq 200) {
        # Download
        $csv = $csv.TrimEnd(",")
        $obj = Get-SPOUserProfileProperty -Account $csv.Split(",") -ErrorAction SilentlyContinue
        $coll += $obj
		
        # clear
        $csv = ""
        $i = 0
        Write-Host "." -NoNewLine
    }
    $i++
}
$coll | Export-Csv "sharepoint-profiles.csv"