# SPJeff-Inventory-SPKG-Available.ps1
# Scan all site collection app catalogs for available SPPKG files and write report to CSV with full SPPKG details

# Load Modules
Import-Module PnP.PowerShell

# Memory collection
$coll = @()

# Load CSV with SiteURLs and loop through each site
$tenantAdminUrl = "https://spjeffdev-admin.sharepoint.com"

# Connect to the site
Connect-PnPOnline -Url $tenantAdminUrl -UseWebLogin -WarningAction SilentlyContinue
# $token = Get-PnPAccessToken

# Open the Tenant App Catalog  
$catalogUrl = Get-PnPTenantAppCatalogUrl

# Open the Site Collection App Catalogs
$tenantAppCatalogSite = Get-PnPSiteCollectionAppCatalog
$tenantAppCatalogSite.Count

# Append array with Tenant App Catalog new PSObject with property AbsoluteUrl
$tenantAppCatalogSite += [PSCustomObject]@{
    AbsoluteUrl = $catalogUrl
}

# Loop through each site collection app catalog
foreach ($site in $tenantAppCatalogSite) {
    $global:siteUrl = $site.AbsoluteUrl
    $global:siteUrl

    # Connect to the site
    Connect-PnPOnline -Url $global:siteUrl -UseWebLogin -WarningAction SilentlyContinue

    # Get all SPPKG files in the Tenant App Catalog
    $files = Get-PnPListItem -List "Apps for SharePoint" -Fields "FileLeafRef", "FileRef", "Title", "ID"

    # Loop through each file
    foreach ($file in $files) {
        $fileUrl = $file["FileRef"]
        $fileName = $file["FileLeafRef"]
        $fileTitle = $file["Title"]
        $fileID = $file["ID"]

        # Match with PNP App
        $app = Get-PnPApp -Scope "Site" | Where-Object { $_.Title -eq $fileTitle }

        # Write to CSV
        $coll += [PSCustomObject]@{
            SiteUrl              = $global:siteUrl
            FileUrl              = $fileUrl
            FileName             = $fileName
            FileTitle            = $fileTitle
            FileID               = $fileID
            AppCatalogVersion    = $app.AppCatalogVersion
            Deployed             = $app.Deployed
            AppId                = $app.Id
            IsClientSideSolution = $app.IsClientSideSolution
        }
    }
}

# Write to CSV
$coll | Export-Csv -Path "SPJeff-SPKG-Available.csv" -NoTypeInformation -Force