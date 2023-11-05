# PnP-PowerShell-spjeff-Connect-PNPOnline.ps1

# PNP Connect
# https://pnp.github.io/powershell/articles/connecting.html
# https://pnp.github.io/powershell/articles/authentication.html
# https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/register-pnpazureadapp?view=sharepoint-ps
# https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
# https://mmsharepoint.wordpress.com/2018/12/19/modern-sharepoint-authentication-in-azure-automation-runbook-with-pnp-powershell/

# Scope
$tenant = "spjeff"

# Azure Certificate
$password = "password"
$secPassword = $password | ConvertTo-SecureString -AsPlainText -Force
$cert = Get-AutomationCertificate -Name 'PNP-PowerShell'
$pfxCert = $cert.Export("pfx" , $password ) # 3=Pfx
$certPath = "PNP-PowerShell.pfx"
Set-Content -Value $pfxCert -Path $certPath -Force -Encoding Byte 

# Connect
$clientId = "client-id-guid"
Connect-PnPOnline -ClientId $clientId -Url "https://$tenant.sharepoint.com" -Tenant "$tenant.onmicrosoft.com" -CertificatePath $certPath -CertificatePassword $secPassword
Get-PnPTenantSite | Format-Table -AutoSize
