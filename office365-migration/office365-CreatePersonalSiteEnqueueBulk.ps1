# CSOM method to provision MySite /personal/ sites in Office 365

# tenant
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
$webUrl = "https://tenant.sharepoint.com"
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl)

# admin user
$web = $ctx.Web
$username = "admin@tenant.onmicrosoft.com"
$password = read-host -AsSecureString

# context
$ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username,$password )
$ctx.Load($web)
$ctx.ExecuteQuery()

# assembly
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles")
$loader =[Microsoft.SharePoint.Client.UserProfiles.ProfileLoader]::GetProfileLoader($ctx)

#To Get Profile
$profile = $loader.GetUserProfile()
$ctx.Load($profile)
$ctx.ExecuteQuery()
$profile 

#To enqueue Profile
$loader.CreatePersonalSiteEnqueueBulk(@("user@tenant.onmicrosoft.com")) 
$loader.Context.ExecuteQuery()