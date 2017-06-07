# Export Office 365 user profiles to CSV (Azure AD & SharePoint UPS)
# --------------------
# Download -- Microsoft Online Services Sign-in Assistant for IT Professionals RTW http://go.microsoft.com/fwlink/p/?LinkId=286152
# Download -- Windows Azure Active Directory Module for Windows PowerShell (64-bit version) http://go.microsoft.com/fwlink/p/?linkid=236297
# Download -- SharePoint Server 2013 Client Components CSOM SDK https://www.microsoft.com/en-us/download/details.aspx?id=35585

# Plugins
Import-Module MSOnline
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles") | Out-Null

# Tenant
$tenantAdminUrl = "https://tenant-admin.sharepoint.com/"
$username = "admin@tenant.onmicrosoft.com"
$password = "pass@word1"

# Date stamp
$dateStamp = Get-Date -Format yyyyMMdd-hhmm

# Azure AD (AAD) - All Profiles
# --------------------
function aadProfiles() {
    Write-Host "Azure AD (AAD)"
    $secPassword = ConvertTo-SecureString $password -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential ($username, $secPassword)
    Connect-MsolService -Credential $cred
    $msolUsers = Get-MsolUser -All
    $msolUsers | Export-Csv "get-msoluser-$dateStamp.csv" -NoTypeInformation -Force
}

# SharePoint Online (SPO) - User Profile Service (UPS) ASMX
# --------------------
function spoProfiles() {
    # Take the AdminAccount and the AdminAccount password, and create a credential
    $secpw = ConvertTo-SecureString $password -AsPlainText -Force
    $creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $secpw)
    
    # Add the path of the User Profile Service to the SPO admin URL, then create a new webservice proxy to access it
    $proxyaddr = "$tenantAdminUrl/_vti_bin/UserProfileService.asmx?wsdl"
    $UserProfileService = New-WebServiceProxy -Uri $proxyaddr -UseDefaultCredential $false
    $UserProfileService.Credentials = $creds
    
    # Set variables for authentication cookies
    $strAuthCookie = $creds.GetAuthenticationCookie($tenantAdminUrl)
    $uri = New-Object System.Uri($tenantAdminUrl)
    $container = New-Object System.Net.CookieContainer
    $container.SetCookies($uri, $strAuthCookie)
    $UserProfileService.CookieContainer = $container
    
    # Sets the first User profile, at index -1
    $UserProfileResult = $UserProfileService.GetUserProfileByIndex(-1)
    $NumProfiles = $UserProfileService.GetUserProfileCount()
    $i = 1
    
    # As long as the next User profile is NOT the one we started with (at -1)...
    $ups = @()
    While ($UserProfileResult.NextValue -ne -1) {
        Write-Host "Examining profile $i of $NumProfiles"
        $props = $UserProfileResult.UserProfile
        $row = New-Object PSObject
        foreach ($p in $props) {
            if ($p.values.value) {
                $row | Add-Member -MemberType NoteProperty -Name $p.Name -Value $p.values.value.ToString()
            }
        }
        $UserProfileResult = $UserProfileService.GetUserProfileByIndex($UserProfileResult.NextValue)
        $ups += $row
        $ups.Count
        $i++
    }

    # Save CSV
    Write-Host "Writing CSV file..."
    $ups | Export-Csv "userprofileservice-asmx-$dateStamp.csv" -NoTypeInformation -Force
}
aadProfiles
spoProfiles