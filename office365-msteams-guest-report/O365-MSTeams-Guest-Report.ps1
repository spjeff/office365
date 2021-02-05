# from https://techcommunity.microsoft.com/t5/microsoft-teams/microsoft-teams-tenant-wide-csv-report/m-p/151875
# from https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps

## Created by SAMCOS @ MSFT, Collaboration with others!
## You must first connect to Microsoft Teams Powershell & Exchange Online Powershell for this to work.
## Links:
## Teams: https://www.powershellgallery.com/packages/MicrosoftTeams/1.0.0
## Exchange: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps
## Have fun! Let me know if you have any comments or asks! 

# Transcript
Start-Transcript

# Install
Install-Module "ExchangeOnlineManagement"
Install-Module -Name "MicrosoftTeams"

# Import
Import-Module "ExchangeOnlineManagement"
Import-Module -Name "MicrosoftTeams"

# Connect
Connect-ExchangeOnline
Connect-MicrosoftTeams 

# Default
$AllTeamsInOrg = (Get-Team).GroupID
$TeamList = @()

# Loop
Write-Output "This may take a little bit of time... Please sit back, relax and enjoy some GIFs inside of Teams!"
Foreach ($Team in $AllTeamsInOrg) {       
    # Parse Inputs
    $TeamGUID           = $Team.ToString()
    $TeamGroup          = Get-UnifiedGroup -identity $Team.ToString()
    $TeamName           = (Get-Team | ? { $_.GroupID -eq $Team }).DisplayName
    $TeamOwner          = (Get-TeamUser -GroupId $Team | ? { $_.Role -eq 'Owner' }).User
    $TeamUserCount      = ((Get-TeamUser -GroupId $Team).UserID).Count
    $TeamCreationDate   = Get-unifiedGroup -identity $team.ToString() | Select -expandproperty WhenCreatedUTC
    $TeamGuest          = (Get-UnifiedGroupLinks -LinkType Members -identity $Team | ? { $_.Name -match "#EXT#" }).Name

    # Zero Guests
    if ($TeamGuest -eq $null) {
        $TeamGuest = "No Guests in Team"
    }

    # Append for CSV
    $TeamList = $TeamList + [PSCustomObject]@{
        TeamName = $TeamName; 
        TeamObjectID = $TeamGUID; 
        TeamCreationDate = $TeamCreationDate; 
        TeamOwners = $TeamOwner -join ', '; 
        TeamMemberCount = $TeamUserCount; 
        TeamSite = $TeamGroup.SharePointSiteURL; 
        AccessType = $TeamGroup.AccessType; 
        TeamGuests = $TeamGuest -join ',' 
    }
}

# Write CSV
$TempFolder = "c:\temp"
New-Item -ItemType "Directory" -Path $TempFolder -ErrorAction "SilentlyContinue" | Out-Null
$TeamList | Export-Csv "$TempFolder\TeamsDatav2.csv" -NoTypeInformation

# Transcript
Stop-Transcript