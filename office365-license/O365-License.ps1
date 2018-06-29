<#
.SYNOPSIS
	Insane Move - Copy sites to Office 365 in parallel.  ShareGate Insane Mode times ten!
.DESCRIPTION
    Copy SharePoint site collections to Office 365 in parallel.  CSV input list of source/destination URLs.  XML with general preferences.
#>

[CmdletBinding()]
param (	
    [Parameter(Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'Verify all Office 365 site collections.  Prep step before real migration.')]
    [Alias("ro")]
    [switch]$readonly = $false,

    [Parameter(Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'Verify all Office 365 site collections.  Prep step before real migration.')]
    [Alias("r")]
    [switch]$report = $false
)

# Config
#REM $srvlic = "srvlic@fnma.onmicrosoft.com"
$srvlic = "acptlictest@acptfanniemae.com"
$domain = "@acptfanniemae.com"
$tenant = "fmacpt"

# Services in License Plan
$EPServicesToAdd = "OFFICESUBSCRIPTION", "EXCHANGE_S_ENTERPRISE", "YAMMER_ENTERPRISE", "SHAREPOINTWAC", "SHAREPOINTENTERPRISE", "TEAMS1", "PROJECTWORKMANAGEMENT"
$EMSServicesToAdd = "RMS_S_PREMIUM", "INTUNE_A", "RMS_S_ENTERPRISE", "AAD_PREMIUM", "MFA_PREMIUM"
$KioskServicesToAdd = "OFFICESUBSCRIPTION"
$VisioServicesToAdd = "VISIO_CLIENT_SUBSCRIPTION"
$ProjectServicesToAdd = "PROJECT_CLIENT_SUBSCRIPTION"

# Plugin
Import-Module ActiveDirectory -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module MSOnline -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module MSOnlineExtended -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module CredentialManager -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

# Log
$datestamp = (Get-Date).tostring("yyyy-MM-dd-hh-mm-ss")
$root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
mkdir "$root\log\" -ErrorAction SilentlyContinue | Out-Null

function Report {
    Write-Host "> PrepareReport" -Fore Yellow
    # Create table with Schema (columns)
    $global:dtReport = New-Object System.Data.DataTable("Report")
    @("UserPrincipalName", "DisplayName", "IsLicensed", "AccountSkuId", "ServicePlan", "ProvisioningStatus") | % {
        $global:dtReport.Columns.Add($_) | Out-Null
    }

    # All users
    $users = Get-MsolUser -All

    # Gather data
    foreach ($u in $users) {
        foreach ($l in $u.Licenses) {
            foreach ($ss in $l.ServiceStatus) {
                $row = $global:dtReport.NewRow()
                $row["UserPrincipalName"] = $u.UserPrincipalName
                $row["DisplayName"] = $u.DisplayName
                $row["IsLicensed"] = $u.IsLicensed
                $row["AccountSkuId"] = $l.AccountSkuId
                $row["ServicePlan"] = $ss.ServicePlan.ServiceName
                $row["ProvisioningStatus"] = $ss.ProvisioningStatus
                $global:dtReport.Rows.Add($row)
            }
        }
    }

    # Save CSV
    $global:dtReport | Export-Csv "d:\O365-report.csv" -NoTypeInformation
}
function ConnectO365 {
    # Read from Windows O/S Credential Manager
    $cred = Get-StoredCredential -Target $srvlic
    if (!$cred) {
        # Prompt and save
        Write-Host $srvlic -Fore Green
        $secpw = Read-Host -AsSecureString -Prompt "Enter Password: "
        New-StoredCredential -Target $srvlic -Username $srvlic -SecurePassword $secpw
        $cred = Get-StoredCredential -Target $srvlic
    }
    # Connect to Office 365
    Connect-MsolService -Credential $cred

    # Display SKU summary
    Get-MsolAccountSku | ft -AutoSize
}
function PrepareTable() {
    Write-Host "> PrepareTable" -Fore Yellow
    # Create table with Schema (columns)
    $global:dtLicenseO365 = New-Object System.Data.DataTable("O365")
    $global:dtLicenseNeed = New-Object System.Data.DataTable("Need")
    @("login", "SKU") | % {
        $global:dtLicenseO365.Columns.Add($_) | Out-Null
        $global:dtLicenseNeed.Columns.Add($_) | Out-Null
    }
}
function RecordLicense($need, $msg, $users, $skuActiveDirectory) {
    # Append to tracking table
    Write-Host "> RecordLicense $msg - $($users.count) $skuActiveDirectory" -Fore Green
    foreach ($u in $users) {
        # Add to table
        if ($need) {
            # AD License Need
            # AD Need with PACK
            $rowNeed = $global:dtLicenseNeed.NewRow()
            $login = $u.SamAccountName
            $rowNeed["login"] = $login
            $rowNeed["SKU"] = $skuActiveDirectory
            $global:dtLicenseNeed.Rows.Add($rowNeed) | Out-Null
        }
        else {
            # O365 License Have
            # Loop SKU
            $sku = ""
            foreach ($l in $u.Licenses) {
                # User Login
                $login = $u.UserPrincipalName.split("@")[0]
                $sku = $l.AccountSkuId
                $rowHave = $global:dtLicenseO365.NewRow()
                $rowHave["login"] = $login
                $rowHave["SKU"] = $sku
                $global:dtLicenseO365.Rows.Add($rowHave)
            }
            
        }
    }
}
function DetectLicenceO365() {
    Write-Host "> DetectLicenceO365" -Fore Yellow
    $users = Get-MsolUser -All
    #REM $users = $users |? {$_.userprincipalname -eq "s6uayox@acptfanniemae.com"}
    RecordLicense $false "DetectLicenceO365" $users "DetectLicenceO365"
}
function DetectLicenseNeed() {
    Write-Host "> DetectLicenseNeed" -Fore Yellow

    
    # DefaultUsers
    $users = Get-ADUser -Properties extensionAttribute9,extensionAttribute15,msExchRecipientTypeDetails,employeeType -ResultSetSize $null -Filter {UserprincipalName -like "*$domain" -and enabled -eq $True} | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and ($_.msExchRecipientTypeDetails -eq 1 -or $_.msExchRecipientTypeDetails -eq 2147483648)} 
    $users = Get-ADUser s6ujikx -Properties extensionAttribute9, extensionAttribute15, msExchRecipientTypeDetails, employeeType 
    $users.Count
    RecordLicense $true "DefaultUsers" $users "$($tenant):ENTERPRISEPACK"

    # EMSUsers
    $users.Count
    RecordLicense $true "EMSUsers" $users "$($tenant):EMS"

    # KIOSKUsers
    $users = Get-ADUser -Properties extensionAttribute9,extensionAttribute15,msExchRecipientTypeDetails,employeeType -ResultSetSize $null -Filter {UserprincipalName -like "*$domain" -and enabled -eq $True} | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and $_.msExchRecipientTypeDetails -eq $null} 
    $users.Count
    RecordLicense $true "KIOSKUsers" $users "$($tenant):KIOSK"

    # VisioUsers
    $group = Get-ADGroup "SG-FM-ICOM-MSFT-VISIO-365-C2R-X86"
    $users = Get-ADGroupMember -Identity $group | Get-ADUser -Properties extensionAttribute15,employeeType,userAccountControl | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and $_.userAccountControl -ne 514 -and $_.userAccountControl -ne 546 -and $_.userAccountControl -ne 66050} 
    $users.Count
    RecordLicense $true "VisioUsers" $users "$($tenant):VISIOCLIENT"

    # ProjectUsers
    $group = Get-ADGroup "SG-FM-ICOM-MSFT-PROJECT-365-C2R-X86"
    $users = Get-ADGroupMember -Identity $group | Get-ADUser -Properties extensionAttribute15,employeeType,userAccountControl | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and $_.userAccountControl -ne 514 -and $_.userAccountControl -ne 546 -and $_.userAccountControl -ne 66050} 
    $users.Count
    RecordLicense $true "ProjectUsers" $users "$($tenant):PROJECTCLIENT"

    # Summary
    Write-Host "dtLicenseNeed rows = $($global:dtLicenseNeed.Rows.Count)"
}
function GrantRevoke() {
    $global:dtLicenseO365.WriteXML("d:\365.xml", $true);
    $global:dtLicenseNeed.WriteXML("d:\need.xml", $true);
    Write-Host "> GrantRevoke" -Fore Yellow

    # Grant - Need and Missing in O365
    $grant = 0
    $dv = New-Object System.Data.DataView $global:dtLicenseO365
    $dv.Sort = "login"
    foreach ($row in $global:dtLicenseNeed.Rows) {
        $login = $row["login"]
        $sku = $row["SKU"]

        if ($sku -like "*KIOSK") {
            # KIOSK Sublicense
            $sku = $sku.Replace("KIOSK","ENTERPRISEPACK")
            $EPSubLicense = $KioskServicesToAdd
        } else {
            # ENTERPRISEPACK Pack Sublicense
            $EPSubLicense = $EPServicesToAdd
        }

        $dv.RowFilter = "login='" + $login + "' AND SKU='" + $sku + "'"
        if ($dv.Count -eq 0) {
            # Grant Display
            Write-Host "GRANT >> $sku - $login" -Fore Green
            $grant++

            # Prepare Sublicense
            if ($sku -like "*ENTERPRISEPACK") {
                $sub = $EPSubLicense
            }
            if ($sku -like "*EMS") {
                $sub = $EMSServicesToAdd
            }
            if ($sku -like "*VISIOCLIENT") {
                $sub = $VisioServicesToAdd
            }
            if ($sku -like "*PROJECTCLIENT") {
                $sub = $ProjectServicesToAdd
                    
            }

            # Grant
            Modify-SubLicense -upn ($login + $domain) -PrimaryLicense $sku -SublicensesToAdd $sub
        }
    }

    # Revoke - Don't need and Have in O365
    $revoke = 0
    $dv = New-Object System.Data.DataView $global:dtLicenseNeed
    $dv.Sort = "login"
    foreach ($row in $global:dtLicenseO365.Rows) {
        $login = $row["login"]
        $sku = $row["SKU"]
        $dv.RowFilter = "login='" + $login + "' AND SKU='" + $sku + "'"
        if ($dv.Count -eq 0) {
            # Revoke Display
            if ($login -eq "s6ujikx") {
                Write-Host "REVOKE >> $sku - $login" -Fore Red
                $revoke++

                # Revoke Permission
                Modify-SubLicense -upn ($login + $domain) -PrimaryRevoke $sku 
            }
        }
    }

    # Summary
    Write-Host "Grant : $grant" -Fore Green
    Write-Host "Revoke: $revoke" -Fore Red
}

# http://sharepointjack.com/2016/modify-sublicense-powershell-function-for-modifying-office-365-sublicenses/
function Modify-SubLicense($upn, $PrimaryLicense, $SublicensesToAdd, $SublicensesToRemove, $PrimaryRevoke) {
    Write-Host "Modify-SubLicense"

    # Revoke primary
    if ($PrimaryRevoke) {
        Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $PrimaryRevoke
        return
    }

    #assemble a list of sub-licenses types the user has that are currently disabled, minus the one we're trying to add 
    $spouser = Get-MsolUser -UserPrincipalName $upn
    $disabledServices = $($spouser.Licenses |? {$_.AccountSkuID -eq $PrimaryLicense}).servicestatus | where {$_.ProvisioningStatus -eq "Disabled"}  | select -expand serviceplan | select ServiceName 
	
    #disabled items need to be in an array form, next 2 lines build that...
    $disabled = @()
    foreach ($item in $disabledServices.servicename) {$disabled += $item}
    Write-Host "   DisabledList before changes: $disabled" -Foregroundcolor yellow
	
    # If there are other sublicenses to be removed (Passed in via -SublicensesToRemove) then lets add those to the disabled list.
    foreach ($license in $SublicensesToRemove) {$disabled += $license }
	
    # Cleanup duplicates in case the license to remove was already missing
    $disabled = $disabled | select -unique
	
    # If there are licenses to ADD, we need to REMOVE them from the list of disabled licenses
    # http://stackoverflow.com/questions/8609204/union-and-intersection-in-powershell
    $disabled = $disabled | ? {$SublicensesToAdd -NotContains $_}
    Write-Host "    DisabledList after changes: $Disabled" -ForegroundColor green
    
    # Apply
    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $PrimaryLicense -DisabledPlans $disabled
    $LicenseOptions    
    Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $PrimaryLicense
    Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions  $LicenseOptions
    Write-Host "OK"
}


function Main() {
    # Log
    $log = "$root\log\O365-License-$datestamp.log"
    Start-Transcript $log
    $start = Get-Date

    if ($report) {
        # Report
        ConnectO365
        Report
    }
    else {
        # Core
        ConnectO365
        PrepareTable
        DetectLicenceO365
        DetectLicenseNeed
        GrantRevoke
    }

    # Cleanup
    Write-Host "--- Run Duration"
    $now = Get-Date
    ($now - $start)
    Stop-Transcript
}
Main