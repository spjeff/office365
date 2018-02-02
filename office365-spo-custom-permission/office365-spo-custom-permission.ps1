# Credentials to connect to office 365 site collection url 
$url = "https://tenant.sharepoint.com/sites/team"
$username = "spadmin@tenant.onmicrosoft.com"
$password = "pass@word1"
$secPassword = $password | ConvertTo-SecureString -AsPlainText -Force

# Load CSOM
Write-Host "Load CSOM libraries" -Foregroundcolor Black -Backgroundcolor Yellow
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
Write-Host "CSOM libraries loaded successfully" -Foregroundcolor black -Backgroundcolor Green 

# Connect
Write-Host "Authenticate to SharePoint Online site collection $url and get ClientContext object" -Foregroundcolor black -Backgroundcolor yellow  
$context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $secPassword) 
$Context.Credentials = $cred
$context.RequestTimeOut = 1000 * 60 * 10
$web = $context.Web
$site = $context.Site 
$context.Load($web)
$context.Load($site)
try {
    $context.ExecuteQuery()
    Write-Host "Authenticated to SharePoint Online $url" -Foregroundcolor black -Backgroundcolor Green
}
catch {
    Write-Host "Not able to authenticate to SharePoint Online $url - $($_.Exception.Message)" -Foregroundcolor black -Backgroundcolor Red
    return
}

# Microsoft custom permission levels
# from https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.permissionkind.aspx

function CreateRoleDefinitions($permName, $permDescription, $clone, $addPermissionString, $removePermissionString) {
    $roleDefinitionCol = $web.RoleDefinitions
    $Context.Load($roleDefinitionCol)
    $Context.ExecuteQuery()

    # Check if the permission level is exists or not
    $permExists = $roleDefinitionCol |? {$_.Name -eq $permName}
    $clonePerm = $roleDefinitionCol |? {$_.Name -eq $clone}
    
    Write-Host Creating Pemission level with the name $permName  -Foregroundcolor black -Backgroundcolor Yellow
    if (!$permExists) {
        try {
            $spRoleDef = New-Object Microsoft.SharePoint.Client.RoleDefinitionCreationInformation
            $spBasePerm = New-Object Microsoft.SharePoint.Client.BasePermissions
			
            if ($clonePerm) {
                $spBasePerm = $clonePerm.BasePermissions
            }
            if ($addPermissionString) {
                $addPermissionString.split(",") | % { $spBasePerm.Set($_) }
            }
            if ($removePermissionString) {
                $removePermissionString.split(",") | % { $spBasePerm.Clear($_) }
            }
            $spRoleDef.Name = $permName
            $spRoleDef.Description = $permDescription
            $spRoleDef.BasePermissions = $spBasePerm    
            $web.RoleDefinitions.Add($spRoleDef)

            $Context.ExecuteQuery()
            Write-Host "Pemission level with the name $permName created" -Foregroundcolor black -Backgroundcolor Green
        }
        catch {
            Write-Host "There was an error creating Permission Level $permName : Error details $($_.Exception.Message)" -Foregroundcolor black -backgroundcolor Red
        }
    }
    else {
        Write-Host "Pemission level with the name $permName already exists" -Foregroundcolor black -Backgroundcolor Red
    }
}
 
#calling role definition function
# CreateRoleDefinitions -permName "Test" -permDescription "Test - custom level" -addPermissionString "addListItems, editListItems, viewListItems"
CreateRoleDefinitions -permName "NoDelete" -permDescription "Contribute - without Delete" -clone "Contribute" -removePermissionString "DeleteListItems"
CreateRoleDefinitions -permName "AddOnly" -permDescription "Contribute - without Edit or Delete" -clone "Contribute" -removePermissionString "DeleteListItems,EditListItems"