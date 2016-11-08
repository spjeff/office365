# from https://blogs.msdn.microsoft.com/frank_marasco/2014/03/25/so-you-want-to-programmatically-provision-personal-sites-one-drive-for-business-in-office-365/

# tenant
$webUrl = "https://tenant-admin.SharePoint.com"
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
$loader.CreatePersonalSiteEnqueueBulk(@("admin@tenant.onmicrosoft.com")) 
$loader.Context.ExecuteQuery()