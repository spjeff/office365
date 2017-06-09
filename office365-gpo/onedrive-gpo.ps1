<#
.SYNOPSIS
	Disables allow Allow members to share setting on OneDrive sites and webs

.DESCRIPTION
	Disables Members Can Share setting
#>

$datestamp = (Get-Date).tostring("yyyy-MM-dd-hh-mm-ss")
Start-Transcript "onedrive-gpo-$datestamp.csv"

# Plugins
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
Import-Module SharePointPnPPowerShellOnline -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

# Settings
[xml]$settings = Get-Content "OneDriveGPO.xml"
$admins = @()
$adminUsers = $settings.settings.adminUsers.split(",")
foreach ($adminUser in $adminUsers) {
    $admins += $adminUser.Trim()
}
$tenant = $settings.settings.tenant
$adminURL = "https://$tenant-admin.sharepoint.com"

# Grant Site Collection Admin with SPO
Function GrantSCA ($url) {
    $site = Get-SPOSite $url
    foreach ($admin in $admins) {
        $user = Get-SPOUser -Site $site -LoginName $admin
        if (!$user.IsSiteAdmin) {
            Set-SPOUser -Site $site -LoginName $admin -IsSiteCollectionAdmin $true
        }
    }
}

# Main
Function Main() {

    # Credentials
   	$secpw = ConvertTo-SecureString -String $settings.settings.adminPass -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential ($settings.settings.adminUserName, $secpw)
    Connect-PnPOnline $adminURL
    Connect-SPOService -URL $adminURL -Credential $cred

    # Azure AD (AAD) - All Profiles
    Write-Host "Azure AD (AAD)"
    $tracker = @()
    Connect-MsolService -Credential $cred
    $msolUsers = Get-MsolUser -All
    foreach ($msolUser in $msolUsers) {
        $profile = Get-PNPUserProfileProperty -Account $msolUser.UserPrincipalName
        $personalURLs += New-Object -TypeName PSObject -Props {@('UPN'=$msolUser.UserPrincipalName; 'Url'=$profile.PersonalURL; 'State'='New')}
    }

    # Grant Site Collection Admin
    foreach ($url in $csv.SiteUrl) {
        GrantSCA $url
    }

    # ScriptBlock
    $sb = { param($url)

        # Connect
        Connect-PnPOnline -Url $url

        # Enable Versioning
        $DocLib = Get-PnPList -Identity "Documents"
        if (!$DocLib.EnableVersioning) {
            Set-PnPList -Identity "Documents" -EnableVersioning $true -MajorVersions 99
            Write-Host "$url - Versioning Enabled" -Fore Green
        }
        else {
            Write-Host "$url - already enabled" -Fore Magenta
        }
    
        # Disable AllowMembersEditMembership for SubWeb Member Groups
        $RootWeb = Get-PnPWeb -Includes MembersCanShare
        $SubWebs = Get-PnPSubWebs -Recurse
        $AllWebs = @()
        $AllWebs += $RootWeb
 
        # Member Invites Disabled
        foreach ($SubWeb in $SubWebs) {
            $Sub = Get-PnPWeb -Identity $SubWeb -Includes MembersCanShare, AssociatedMemberGroup
            $AllWebs += $Sub
            $MemberGroupTitle = $Sub.AssociatedMemberGroup.Title
            if ($MemberGroupTitle) {
                $MemberGroup = Get-PnPGroup -Identity $MemberGroupTitle -Includes AllowMembersEditMembership
                if ($MemberGroup.AllowMembersEditMembership) {
                    $GroupTitle = $MemberGroup.Title
                    Set-PnPGroup -Identity $GroupTitle -AllowMembersEditMembership $False
                    Write-Host "$GroupTitle - Member Invites Disabled" -Fore Yellow
                }
                else {
                    Write-Host "$GroupTitle - Member Invites Already Disabled" -Fore DarkCyan
                }
            }
            else {
                Write-Host "No Member Group Found" -Fore Red
            }
        }

        # Disable MembersCanShare in All Webs
        foreach ($Web in $AllWebs) {
            $WebTitle = $Web.ServerRelativeUrl
            if ($Web.MembersCanShare) {
                $Web.MembersCanShare = $false
                $Web.Update()
                $Web.Context.ExecuteQuery()
                Write-Host "$WebTitle - MembersCanShare Disabled" -Fore Yellow
            }
            else {
                Write-Host "$WebTitle - MembersCanShare Already Disabled" -Fore DarkCyan
            }
        }

        # Enable Audit
        $audit = Get-PnPAuditing
        if ($audit.AuditFlags -eq "All") {
            Write-Host "Audit OK" -Fore Green
        }
        else {
            Write-Host "Enabling Audit" -Fore Yellow
            Set-PnPAuditing -RetentionTime 30 -TrimAuditLog -EnableAll
        }

        # Add Custom JS
        $scriptName = $settings.settings.scriptName
        $scriptUrl = $settings.settings.scriptUrl
        Write-Host "Found $scriptName" -Fore Green
        Write-Host "Found $scriptUrl" -Fore Green

        # Detect Custom JS
        $found = Get-PnPJavaScriptLink -Scope All |? {$_.Name -eq $scriptName}
        if ($found) {
            # Found
            Write-Host "Found $scriptName" -Fore Green
            $found | ft -a
        }
        else {
            # Add JS Script Link missing
            Write-Host "Adding $scriptName" -Fore Yellow
            Add-PnPJavaScriptLink -Name $scriptName -Url $scriptUrl -Sequence 99 -Scope Site
        }
    }

    # Parallel PNP to apply standard settings to destiation sites
    $numWorkers = $settings.settings.workers
    do {
        #Pull active workers
        $pending = $tracker |? {$_.State -ne "Complete"}
        $activeWorkers = Get-Job |? {$_.State -eq "Running" -or $_.State -eq "NotStarted"}

        if ($activeWorkers -lt $numWorkers) {
            #Create new worker
            $nextSite = $pending[0].Url
            $found =  $tracker |? {$_.UPN -eq $pending[0].UPN}
            $found.State = "Complete"
            Write-Host "Next site..." -ForegroundColor Yellow
            Invoke-Command -ScriptBlock $sb -ArgumentList $nextSite -AsJob
        }
        
        #Clean up
        $idleWorkers = Get-Job |? {$_.State -ne "Running" -and $_.State -ne "NotStarted"}
        $idleWorkers | Remove-Job
    }
    while ($pending)
}
Main
Stop-Transcript