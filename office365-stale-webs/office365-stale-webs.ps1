<# === Office365 - Stale Webs === 
 * leverages 3 libraries (SPO, PNP, CSOM)
 * leverages parallel PowerShell
 * grant Site Collection Admin for support staff
 * apply Site Collection quota 5GB  (if none)
 * enable Site Collection Auditing
 * enable Site Collection Custom Action JS ("office365-gpo.js")
#>

#Config
$AdminUrl = "https://tenant-admin.sharepoint.com"
$UserName = "admin@tenant.onmicrosoft.com"
$Password = "pass@word1"
$ThresholdDays = 180
$ReminderDays = 6
$MaxReminders = 4
$EmailFrom = "sharepoint_support@tenant.com"

Function defineMetrics() {
    # Global cache
    $global:dtWebs = New-Object "System.Data.DataTable" -ArgumentList "Webs"
		
    # Global counters
    $global:countScan = 0
    $global:countStale = 0
    $global:countDelete = 0

    # Schema
    $col = $global:dtWebs.Columns.Add("URL", [String])
    $col = $global:dtWebs.Columns.Add("Title", [String])
    $col = $global:dtWebs.Columns.Add("Last Modified Time Web", [String])
    $col = $global:dtWebs.Columns.Add("Owner", [String])
    $col = $global:dtWebs.Columns.Add("IsRootWeb", [Boolean])
    $col = $global:dtWebs.Columns.Add("StaleWebs_WarningCount", [Int])
    $col = $global:dtWebs.Columns.Add("StaleWebs_EmailSentTime", [DateTime])
    $col = $global:dtWebs.Columns.Add("StaleWebs_EmailSent", [Boolean])
    $col = $global:dtWebs.Columns.Add("Deleted", [Boolean])
    $col = $global:dtWebs.Columns.Add("Script Run Date", [DateTime])
}

Function emailReminder($final, $title, $url, $to) {
    # Final notification
    if ($final) {
        $file = "Stale_Webs_Email_Site_Final.htm"
    }
    else {
        $file = "Stale_Webs_Email_Site_Owner.htm"
    }

    # Site Owner
    $html = Get-Content $file
    $subject = $html[0]
    $body = ($html | Select -Skip 1 | Out-String) -f $title, $url

    # Send Email
    emailCloud $to, $subject, $body  
}
Function emailSummary() {
    # SharePoint Admin team
    
    # Pivot table and count
    #TODO

    # Summary
    $file = "Stale_Webs_Email_Summary.htm"
    $html = Get-Content $file
    $subject = $html[0]
    $body = ($html | Select -Skip 1 | Out-String) -f $title, $url

    # Send Email
    emailCloud $EmailSupport, $subject, $body
}

Function emailCloud($to, $subject, $body) {
    # Get the PowerShell credential and prints its properties 
    $MyCredential = "O365SMTP"
    $cred = Get-AutomationPSCredential -Name $MyCredential 
    if ($cred -eq $null) {return}

    Send-MailMessage -To $to -Subject $subject -Body $body -UseSsl -Port 587 -SmtpServer 'smtp.office365.com' -From $EmailFrom -BodyAsHtml -Credential $cred 
}

Function processWeb($web) {
    Write-Host "Processing web $($web.Url)"

    # Current site - Is stale?
    $url = $web.Url
    $stale = $false
    $lists = Get-PnPList -Web $web -Includes "LastModified" | % {New-Object PSObject -Property @{LastModified = $_.LastModified; }}
    $mostRecentList = ($lists | sort LastModified -Desc)[0]
    $age = (Get-Date) - $mostRecentList.LastModified
    if ($age.Days -gt $global:ThresholdDays) {
        $stale = $true
    }

    # Current property bag
    $StaleWebs_WarningCount = Get-SPOPropertyBag -Key "StaleWebs_WarningCount"
    $StaleWebs_EmailSentTime = Get-SPOPropertyBag -Key "StaleWebs_EmailSentTime"
    
    if ($stale) {
        if (!$StaleWebs_WarningCount) {
            # First notification
            Set-SPOPropertyBag -Key "StaleWebs_WarningCount" -Value 1
            Set-SPOPropertyBag -Key "StaleWebs_EmailSentTime" -Value (Get-Date)

            # Add row to table
            $newRow = $global:dtWebs.NewRow()
            $newRow["URL"] = $url
            $newRow["Title"] = $web.Tile
            $newRow["Last Modified Time Web"] = $mostRecentList.LastModified
            $newRow["Owner"] = $web.RequestAccessEmail
            $newRow["IsRootWeb"] = $web.IsRootWeb
            $newRow["StaleWebs_WarningCount"] = $StaleWebs_WarningCount 
            $newRow["StaleWebs_EmailSentTime"] = $StaleWebs_EmailSentTime
            $newRow["Deleted"] = 0
            $global:dtWebs.Rows.Add($newRow)
            
            # Email notify site owner
            emailReminder $true, $web.Title, $url
        }
        elseif ($StaleWebs_WarningCount -gt $MaxReminders) {
            # Delete
            Write-Host "Deleting web $url"
            Remove-PnPWeb $url -Force
            
            # Add row to table
            $newRow = $global:dtWebs.NewRow()
            $newRow["URL"] = $url
            $newRow["Title"] = $web.Tile
            $newRow["Last Modified Time Web"] = $mostRecentList.LastModified
            $newRow["Owner"] = $web.RequestAccessEmail
            $newRow["IsRootWeb"] = $web.IsRootWeb
            $newRow["StaleWebs_WarningCount"] = $StaleWebs_WarningCount 
            $newRow["StaleWebs_EmailSentTime"] = $StaleWebs_EmailSentTime
            $newRow["Deleted"] = 1
            $global:dtWebs.Rows.Add($newRow)

            # Email notify site owner
            emailReminder $true, $web.Title, $url
        }
        else {
            # Reminder
            $timeSinceLastReminder = (Get-Date) - $StaleWebs_EmailSentTime
            if ($timeSinceLastReminder.Hours -gt $ReminderDays) {
                $StaleWebs_WarningCount++
                $StaleWebs_EmailSentTime = Get-Date
                Set-SPOPropertyBag -Key "StaleWebs_WarningCount" -Value $StaleWebs_WarningCount 
                Set-SPOPropertyBag -Key "StaleWebs_EmailSentTime" -Value $StaleWebs_EmailSentTime

                # Add row to table
                $newRow = $global:dtWebs.NewRow()
                $newRow["URL"] = $url
                $newRow["Title"] = $web.Tile
                $newRow["Last Modified Time Web"] = $mostRecentList.LastModified
                $newRow["Owner"] = $web.RequestAccessEmail
                $newRow["IsRootWeb"] = $web.IsRootWeb
                $newRow["StaleWebs_WarningCount"] = $StaleWebs_WarningCount 
                $newRow["StaleWebs_EmailSentTime"] = $StaleWebs_EmailSentTime
                $newRow["Deleted"] = 0
                $global:dtWebs.Rows.Add($newRow)

                # Email notify site owner
                emailReminder $true, $web.Title, $url
            }
        }
    }
}

Function Main {
    # Log
    Start-Transcript
    $start = Get-Date

    # Metrics
    defineMetrics

    # SPO and PNP modules
    Import-Module -WarningAction SilentlyContinue Microsoft.Online.SharePoint.PowerShell
    Import-Module -WarningAction SilentlyContinue SharePointPnPPowerShellOnline
		
    # Credential
    $secpw = ConvertTo-SecureString -String $Password -AsPlainText -Force
    $c = New-Object System.Management.Automation.PSCredential ($UserName, $secpw)

    # Connect Office 365
    Connect-SPOService -URL $AdminUrl -Credential $c
	
    # Scope
    Write-Host "Opening list of sites ..." -Fore Green
    $sites = Get-SPOSite
    Write-Host $sites.Count

    # Serial loop
    Write-Host "Loop sites "
    ForEach ($s in $sites) {
        Write-Host "." -NoNewLine

        # PNP
        Connect-PnPOnline -Url $s.Url -Credentials $c
        
        # Root
        $root = Get-PnPWeb
        processWeb $root

        # Child webs
        $webs = Get-PnPSubWeb -Recurse
        $webs | % {processWeb $_}
    }

    # Email Summary
    emailSummary

    # Duration
    $min = [Math]::Round(((Get-Date) - $start).TotalMinutes, 2)
    Write-Host "Duration Min : $min"
    Stop-Transcript
}
Main