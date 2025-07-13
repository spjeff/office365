# from https://pnp.github.io/script-samples/spo-get-sharinglinks/README.html?tabs=pnpps
# from https://reshmeeauckloo.com/posts/powershell-get-sharing-links-sharepoint/

# Configuration
$tenantUrl = "https://spjeffdev-admin.sharepoint.com"
$clientId = "0cd6579c-b893-4041-8ec0-dc029b4f34a0"
$clientPFx = "SPJeff-Clean-Sharing-Links.ps1.pfx"
$dateTime = (Get-Date).toString("dd-MM-yyyy-hh-ss")
$fileName = "SharedLinks-" + $dateTime + ".csv"
$ReportOutput = $fileName

# Register the PnP PowerShell module if not already registered
#REM Register-PnPAzureADApp 

# Connect to PnP Online with PFX
Connect-PnPOnline -Url $tenantUrl -ClientId $clientId -CertificatePath $clientPFx -Tenant "spjeffdev.onmicrosoft.com" -ErrorAction Stop
Get-PnPContext

$global:Results = @();

function getSharingLink($_object,$_type,$_siteUrl,$_listUrl)
{
    $relativeUrl = $_object.FieldValues["FileRef"]
    $SharingLinks = if ($_type -eq "File" -or $_type -eq "Item") {
        Get-PnPFileSharingLink -Identity $relativeUrl
    } elseif ($_type -eq "Folder") {
        Get-PnPFolderSharingLink -Folder $relativeUrl
    }
    
    ForEach($ShareLink in $SharingLinks)
    {
        $result = New-Object PSObject -property $([ordered]@{
            SiteUrl = $_SiteURL
            listUrl = $_listUrl
            Name = $_type -eq 'Item' ? $_object.FieldValues["Title"] : $_object.FieldValues["FileLeafRef"]          
            RelativeURL = $_object.FieldValues["FileRef"] 
            ObjectType = $_Type
            ShareId = $ShareLink.Id
            RoleList = $ShareLink.Roles -join "|"
            Users = $ShareLink.GrantedToIdentitiesV2.User.Email -join "|"
            ShareLinkUrl  = $ShareLink.Link.WebUrl
            ShareLinkType  = $ShareLink.Link.Type
            ShareLinkScope  = $ShareLink.Link.Scope
            Expiration = $ShareLink.ExpirationDateTime
            BlocksDownload = $ShareLink.Link.PreventsDowload
            RequiresPassword = $ShareLink.HasPassword
                        
        })
        $global:Results +=$result;

        # Remove and cleanup
        Remove-PnPFileSharingLink -FileUrl $relativeUrl -Force
        Write-Host "Removed sharing link for $($_object.FieldValues["FileLeafRef"]) in $($_siteUrl) from list $($_listUrl)" -ForegroundColor Yellow
    }     
}

# Exclude certain libraries
$ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
    "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images"
    , "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", "Preservation Hold Library",
    "Site Pages", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", "Workflow History", "Workflow Tasks", "Pages")

$m365Sites = Get-PnPTenantSite | Where-Object { ( $_.Url -like '*/sites/test') -and $_.Template -ne 'RedirectSite#0' } 
$m365Sites | ForEach-Object {
$siteUrl = $_.Url;    

# Connect to each site
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -CertificatePath $clientPFx -Tenant "spjeffdev.onmicrosoft.com" -ErrorAction Stop
Write-Host "Processing site $siteUrl"  -Foregroundcolor "Red"; 

# Get Sharing Links for all lists in the site
$ll = Get-PnPList -Includes BaseType, Hidden, Title,HasUniqueRoleAssignments,RootFolder | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLists } #$_.BaseType -eq "DocumentLibrary" 
  Write-Host "Number of lists $($ll.Count)";

  # Loop through each list
  foreach($list in $ll)
  {
    $listUrl = $list.RootFolder.ServerRelativeUrl;       

    #Get all list items in batches
    $ListItems = Get-PnPListItem -List $list -PageSize 2000 

        ForEach($item in $ListItems)
        {
            #Check if the Item has unique permissions
            $HasUniquePermissions = Get-PnPProperty -ClientObject $Item -Property "HasUniqueRoleAssignments"
            If($HasUniquePermissions)
            {       
                #Get Shared Links
                if($list.BaseType -eq "DocumentLibrary")
                {
                    $type= $item.FileSystemObjectType;
                }
                else
                {
                    $type= "Item";
                }
                getSharingLink $item $type $siteUrl $listUrl;
            }
        }
    }
 }
 
 # Save the results to a CSV file
 $global:Results | Export-CSV $ReportOutput -NoTypeInformation

# Display the results
$global:Results | Format-Table -AutoSize

# Disconnect from PnP Online
Disconnect-PnPOnline

# Output the report file path
Write-Host "Report saved to: $ReportOutput" -ForegroundColor Green
Write-host -f Green "Sharing Links Report Generated Successfully!"