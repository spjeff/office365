# PNP Register
# https://pnp.github.io/powershell/articles/connecting.html
# https://pnp.github.io/powershell/articles/authentication.html
# https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/register-pnpazureadapp?view=sharepoint-ps
# https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
# https://mmsharepoint.wordpress.com/2018/12/19/modern-sharepoint-authentication-in-azure-automation-runbook-with-pnp-powershell/
# Scope
$tenant = "spjeff"
$clientFile = "PnP-PowerShell-$tenant.txt"
# Register
$password = ConvertTo-SecureString -String "password" -AsPlainText -Force
$reg = Register-PnPAzureADApp -ApplicationName "PnP-PowerShell-$tenant" -Tenant "$tenant.onmicrosoft.com" -CertificatePassword $password -Interactive
$reg."AzureAppId/ClientId" | Out-File $clientFile -Force