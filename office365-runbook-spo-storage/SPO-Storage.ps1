"START"

# Modules
Import-Module "SharePointPnPPowerShellOnline"
Import-Module "Microsoft.PowerShell.Utility"

# Config
$url        = "https://spjeff-admin.sharepoint.com/"

# App ID and Secret from Azure Automation secure string "Credential" storage
# from https://stackoverflow.com/questions/28352141/convert-a-secure-string-to-plain-text
# from https://sharepointyankee.com/2018/02/23/azure-automation-credentials/
$cred = Get-AutomationPSCredential "SPO-Storage-SPOApp"
$cred |ft -a

$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password)
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
$clientId      = $cred.UserName
$clientSecret  = $UnsecurePassword
$clientId
$clientSecret

# PNP Get Sites
Connect-PnPOnline -Url $url -ClientId $clientId -ClientSecret $clientSecret
$sites = Get-PnPTenantSite
$table = $sites | Select-Object Url,Template,StorageUsage,StorageMaximumLevel,LockState,Owner,OwnerEmail,OwnerLoginName,OwnerName
$table | Format-Table -AutoSize

# Format HTML
$CSS = @'
<style>
table {
    border: 1px solid black;
}
th {
    background: #DAE0E6;
    padding-right: 10px;
}
tr:nth-child(even) {
    background: #F1F1F1;
}
</style>
'@
$csv = "SPO-Storage.csv"
$table | Export-CSV $csv -Force -NoTypeInformation
$html = ($table | ConvertTo-HTML -Property * -Head $CSS) -Join ""
$totalStorageUsage = $table | Measure-Object "StorageUsage" -Sum
$html += "<p>Count Sites        = $($sites.Count)</p>"
$html += "<p>Total Storage (MB) = $($totalStorageUsage.Sum)</p>"

# Send Email
$cred = Get-AutomationPSCredential "SPO-Storage-EXOUser"
#REM $cred = Get-Credential
$cred |ft -a
$cred.UserName
$cred.Password
$recip  = "spjeff@spjeff.com"
$subj   = "SPO-Storage"
Send-MailMessage -To $recip -from $recip -Subject $subj -Body $html -BodyAsHtml -smtpserver "smtp.office365.com" -UseSSL -Credential $cred -Port "587" -Attachments $csv

"FINISH"