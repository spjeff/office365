<#
.DESCRIPTION
	Downloads full listing of SharePoint Online (SPO) site collection into JSON. Expands columns with Azure AD properties (listed below).  Finally uploads to SPO document library. 

        * Site Owner UPN
        * Site Owner Display Name
        * Site Owner Email
        * Site Owner Department
        * Site Owner Manager UPN
        * Site Owner Manager Display Name
        * Site Owner Manager Email
	
    Leveage JSON output for future use cases:

        * Site inventory
        * Site history report (version history, time machine, restore by URL).
        * Site ownership for training
        * Site ownership for governance.  If owner left company then locate new owner.
        * Site directory to find by keyword

	Comments and suggestions always welcome!  Please, use the issues panel at the project page.

.EXAMPLE
	.\PnP-Site-Directory-JSON.ps1
	
.NOTES  
	File Name:  PnP-Site-Directory-JSON.ps1
	Author   :  Jeff Jones - @spjeff
	Version  :  1.0
	Modified :  2020-01-23
.LINK
	https://github.com/spjeff/office365
#>


# Import
Import-Module -Name "PNP.PowerShell" -ErrorAction SilentlyContinue | Out-Null
Import-Module -Name "AzureAD" -ErrorAction SilentlyContinue | Out-Null

# Config
$tenant        = "spjeff"
$pfxClientFile = "PnP-PowerShell-$tenant.txt"
$pfxPassword   = "password"
$username      = "spjeff@spjeff.com"

$jsonOutput    = "PnP-Site-Directory.json"
$jsonFolder    = "PnPSiteDirectory"
$dtPeople      = New-Object System.Data.DataTable("people")

# Cache table
function createTable() {
    # Schema column
    @("mail","upn","displayname","managerupn","managerdisplayname","department") |% { $dtPeople.Columns.Add($_) | Out-Null}
}

# Connect both AAD and PNP
function connectCloud() {
    # Connect AAD
    "Connect AAD"
    $secpassword = ConvertTo-SecureString -String $password -AsPlainText -Force
    $cred        = New-Object -Typename "System.Management.Automation.PSCredential" -ArgumentList $username, $secpassword
    $out         = Connect-AzureAD -Credential $cred

    # Connect PNP
    "Connect PNP"
    $pfxClientId    = Get-Content $pfxClientFile
    $pfxSecPassword = $pfxPassword | ConvertTo-SecureString -AsPlainText -Force
    $out            = Connect-PnPOnline -ClientId $pfxClientId -Url "https://$tenant.sharepoint.com" -Tenant "$tenant.onmicrosoft.com" -CertificatePath "PnP-PowerShell-$tenant.pfx" -CertificatePassword $pfxSecPassword
}

# Do we have this user?
function findUser($mail) {
    # Input validation
    if (!$mail) {
        "Not found"
        return
    }

    # DataView rapid filter
    $dvPeople	= New-Object System.Data.DataView($dtPeople)
    $dvPeople.RowFilter = "Mail = '$mail'"

    # Result not found
    if ($dvPeople.Count -eq 0) {
        # Lookup user
        $user = $null
        $user = Get-AzureADUser -Filter "mail eq '$mail'"

        # Lookup manager
        $mgr = $null
        $mgr = Get-AzureADUserManager -ObjectId $user.ObjectId

        # Hash
        $hash = @{
            "mail" = $user.Mail
            "upn" = $user.UserPrincipalName
            "displayname" =$user.DisplayName
            "department"  =$user.UsageLocation
            "managerupn" = $mgr.UserPrincipalName
            "managerdisplayname" =  $mgr.DisplayName
        }

        # Add
        $row = $dtPeople.NewRow()
        $row['mail']               = $hash['mail']
        $row['upn']                = $hash['upn']   
        $row['displayname']        = $hash['displayname'] 
        $row['department']         = $hash['department']
        $row['managerupn']         = $hash['managerupn']  
        $row['managerdisplayname'] = $hash['managerdisplayname']
        $dtpeople.Rows.Add($row)
        Write-Host "Add $mail" -ForegroundColor Green
    }
	
    Write-Host "Found $mail" -ForegroundColor Yellow
	return $dvPeople
}

# Collect input PNP sites
function collectSites() {
    # Download original
    $sites = Get-PnPTenantSite
    "Found sites = $($sites.count)"

    # Convert CSV
    $sites | Export-Csv "temp.csv" -Force -NoTypeInformation
    $rows = Import-csv "temp.csv"

    # Expand columns
    foreach ($row in $rows) {
        $hash = findUser $s.Owner
        $row| Add-Member Noteproperty 'mail'                $hash.mail
        $row| Add-Member Noteproperty 'upn'                 $hash.upn
        $row| Add-Member Noteproperty 'displayname'         $hash.displayname
        $row| Add-Member Noteproperty 'department'          $hash.department
        $row| Add-Member Noteproperty 'managerupn'          $hash.managerupn
        $row| Add-Member Noteproperty 'managerdisplayname'  $hash.managerdisplayname
    }

    # Write JSON local
    "Write JSON"
    $json = $rows | ConvertTo-Json -Depth 9
    $json | Out-File $jsonOutput -Force
}

# Upload JSON to SPO
function uploadJSON() {
    "Upload JSON"
    $out = Add-PnPFile -Path $jsonOutput -Folder $jsonFolder
    $out.ServerRelativeUrl
}

# main
function main() {
    createTable
    connectCloud
    collectSites
    uploadJSON
    "Done"
}
main