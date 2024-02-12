# from https://pnp.github.io/pnpassessment/using-the-assessment-tool/setupauth.html

# Sample for the Microsoft Syntex adoption module. Remove the application/delegated permissions depending on your needs
# and update the Tenant and Username properties to match your environment.
#
# If you prefer to have a password set to secure the created PFX file then add below parameter
# -CertificatePassword (ConvertTo-SecureString -String "password" -AsPlainText -Force)
#
# See https://pnp.github.io/powershell/cmdlets/Register-PnPAzureADApp.html for more options
#

Install-Module PnP.PowerShell
Import-Module PnP.PowerShell
Get-Command -Module PnP.PowerShell

Register-PnPAzureADApp -ApplicationName Microsoft365AssessmentToolForSyntex `
                       -Tenant spjeffdev.onmicrosoft.com `
                       -Store CurrentUser `
                       -GraphApplicationPermissions "Sites.Read.All" `
                       -SharePointApplicationPermissions "Sites.FullControl.All" `
                       -GraphDelegatePermissions "Sites.Read.All", "User.Read" `
                       -SharePointDelegatePermissions "AllSites.Manage" `
                       -Username "spjeffdev@spjeffdev.onmicrosoft.com" `
                       -Interactive

<#

#>