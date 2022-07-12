# Download Azure AD users granted Admin Role Assignments

# Connect
Connect-AzureAD

# Loop Role Definition
$coll = @()
foreach ($rd in (Get-AzureADMSRoleDefinition)) {
	$roleAssignment = Get-AzureADMSRoleAssignment -Filter "roleDefinitionId eq '$($rd.Id)'" -ErrorAction "SilentlyContinue"
	if ($roleAssignment) {
		foreach ($ra in $roleAssignment) {
			$users = Get-AzureADObjectByObjectId -ObjectIds $ra.PrincipalId
			foreach ($u in $users) {
				if ($u.ObjectType -eq "User") {
					$obj = [PSCustomObject]@{
						'Id'                = $ra.Id
						'RoleDefinitionId'  = $ra.RoleDefinitionId
						'PrincipalId'       = $ra.PrincipalId
						'RoleDisplayName'   = $rd.DisplayName
						'RoleIsBuiltIn'     = $rd.IsBuiltIn
						'RoleDescription'   = $rd.Description
						'RoleIsEnabled'     = $rd.IsEnabled
						'UserDisplayName'   = $u.DisplayName
						'UserPrincipalName' = $u.UserPrincipalName
						'UserObjectType'    = $u.UserType
					}
					$coll += $obj
				}
			}
		}
	}
}

# Write CSV
Write-Host "Found $($coll.count) Azure admins" -ForegroundColor "Green"
$stamp = Get-Date -UFormat "%Y-%m-%d-%H-%M-%S"
$file = "AzureAD-Admin-Roles-$stamp.csv"
$coll | Export-Csv $file -NoTypeInformation
Start-Process $file