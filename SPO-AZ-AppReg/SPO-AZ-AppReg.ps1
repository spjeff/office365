# BLOG AT 
# 
# REFERENCES
# ************************************************************
# https://pnp.github.io/powershell/articles/connecting.html
# https://pnp.github.io/powershell/articles/authentication.html
# https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/register-pnpazureadapp?view=sharepoint-ps
# https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
# https://mmsharepoint.wordpress.com/2018/12/19/modern-sharepoint-authentication-in-azure-automation-runbook-with-pnp-powershell/

# STEP 1 - PNP Register AZ AppReg
# ************************************************************

# Modules
Install-Module "PNP.PowerShell"
Import-Module "PNP.PowerShell"
Install-Module "Azure"
Import-Module "Azure"

# Scope
$tenant = "spjeffdev"
$password = "password"
$certname = "SPO-AZ-AppReg-$tenant"

# Register into Azure Application Registration Portal
$secPassword = ConvertTo-SecureString -String $password -AsPlainText -Force
$reg = Register-PnPAzureADApp -ApplicationName "SPO-AZ-AppReg-$tenant" -Tenant "$tenant.onmicrosoft.com" -CertificatePassword $secPassword -Interactive
$reg."AzureAppId/ClientId" | Out-File "$certname-ClientID.txt" -Force

# STEP 2 - PNP Connect
# ************************************************************

# Connect PNP
$clientId = Get-Content "$certname-ClientID.txt"
$certFilename = "$certname.pfx"
Connect-PnPOnline -Url "https://$tenant.sharepoint.com" -ClientId $clientId -Tenant "$tenant.onmicrosoft.com" -CertificatePath $certFilename -CertificatePassword $secPassword

# PNP query to verify.  Pass unit test.
Get-PnPTenantSite | Format-Table -AutoSize

# Add item to SPList
$siteURL = "https://spjeffdev.sharepoint.com/"
$listTitle = "Test"
Connect-PnPOnline -Url $siteURL -ClientId $clientId -Tenant "$tenant.onmicrosoft.com" -CertificatePath $certFilename -CertificatePassword $secPassword
Add-PnPListItem -List $listTitle -Values @{"Title" = "Test"} | Out-Null