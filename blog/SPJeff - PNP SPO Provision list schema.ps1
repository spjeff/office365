# Provision list schema with higher number of fields

# Configuration
$siteUrl = "https://spjeffdev.sharepoint.com/sites/demo/"
$listTitle = "Customer Tracking"

# Connect to SharePoint Online with App ID and App Secret
Connect-PnPOnline -Url $siteUrl -ClientID "TBD" -ClientSecret "TBD"

# MFA popup support

# Connect-PnPOnline -Url $siteUrl -UseWebLogin
Get-PnPWeb

# Open SPList
$list = Get-PnPList -Identity $listTitle

# Open CSV
$csv = Import-Csv "SPJeff - PNP SPO Provision list schema.csv"
# Loop through CSV and add fields to SPList with data type
foreach ($row in $csv) {
    $row |Ft -a
    Add-PnPField -List $list -DisplayName $row.Name -InternalName $row.Name -Type $row.Type -AddToDefaultView
}