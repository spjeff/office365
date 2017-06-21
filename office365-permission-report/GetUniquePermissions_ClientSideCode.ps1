# from http://ericjalexander.com/blog/2016/11/20/Full-Permission-Report-PowerShell
Clear-Host
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Add-Type -Path ".\SharePoint Assemblies\Microsoft.SharePoint.Client.dll"
Add-Type -Path ".\SharePoint Assemblies\Microsoft.SharePoint.Client.Runtime.dll"

# from Gary Lapointe https://www.itunity.com/article/loading-specific-values-lambda-expressions-sharepoint-csom-api-windows-powershell-1249
$pLoadCSOMProperties = (Get-Location).ToString() + "\Load-CSOMProperties.ps1"
. $pLoadCSOMProperties

 # Defaults
$properties = @{SiteUrl = ''; SiteTitle = ''; ListTitle = ''; Type = ''; RelativeUrl = ''; ParentGroup = ''; MemberType = ''; MemberName = ''; MemberLoginName = ''; Roles = ''; }; 
$RootWeb = ""; 
$RootSiteTitle = "";
$ExportFileDirectory = (Get-Location).ToString();

# Prompt Input
$SiteCollectionUrl = Read-Host -Prompt "Enter site collection URL: ";
$Username = Read-Host -Prompt "Enter userName: ";
$password = Read-Host -Prompt "Enter password: " -AsSecureString ;

Function PermissionObject($_object, $_type, $_relativeUrl, $_siteUrl, $_siteTitle, $_listTitle, $_memberType, $_parentGroup, $_memberName, $_memberLoginName, $_roleDefinitionBindings) {
    $permission = New-Object -TypeName PSObject -Property $properties; 
    $permission.SiteUrl = $_siteUrl; 
    $permission.SiteTitle = $_siteTitle; 
    $permission.ListTitle = $_listTitle; 
    $permission.Type = $_type; 
    $permission.RelativeUrl = $_relativeUrl; 
    $permission.MemberType = $_memberType; 
    $permission.ParentGroup = $_parentGroup; 
    $permission.MemberName = $_memberName; 
    $permission.MemberLoginName = $_memberLoginName; 
    $permission.Roles = $_roleDefinitionBindings -join ","; 

    ## Write-Host  "Site URL: " $_siteUrl  "Site Title"  $_siteTitle  "List Title" $_istTitle "Member Type" $_memberType "Relative URL" $_RelativeUrl "Member Name" $_memberName "Role Definition" $_roleDefinitionBindings -Foregroundcolor "Green";
    return $permission;
}


Function QueryUniquePermissionsByObject($_web, $_object, $_Type, $_RelativeUrl, $_siteUrl, $_siteTitle, $_listTitle) {
    $_permissions = @();
  
    Load-CSOMProperties -object $_object -propertyNames @("RoleAssignments") ;

    $ctx.ExecuteQuery() ;
  
    foreach ($roleAssign in $_object.RoleAssignments) {
        $RoleDefinitionBindings = @(); 
        Load-CSOMProperties -object $roleAssign -propertyNames @("RoleDefinitionBindings", "Member");
        $ctx.ExecuteQuery() ;
        $roleAssign.RoleDefinitionBindings | ForEach-Object { 
            Load-CSOMProperties -object $_ -propertyNames @("Name");
            $ctx.ExecuteQuery() ;
            $RoleDefinitionBindings += $_.Name; 
        }
 
        $MemberType = $roleAssign.Member.GetType().Name; 

        $collGroups = "";
        if ($_Type -eq "Site") {
            $collGroups = $_web.SiteGroups;
            $ctx.Load($collGroups);
            $ctx.ExecuteQuery() ;
        }

        if ($MemberType -eq "Group" -or $MemberType -eq "User") { 
 
            Load-CSOMProperties -object $roleAssign.Member -propertyNames @("LoginName", "Title");
            $ctx.ExecuteQuery() ;    
   
            $MemberName = $roleAssign.Member.Title; 
 
            $MemberLoginName = $roleAssign.Member.LoginName;    

            if ($MemberType -eq "User") {
                $ParentGroup = "NA";
            }
            else {
                $ParentGroup = $MemberName;
            }
 
            $_permissions += (PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle $_listTitle $MemberType $ParentGroup $MemberName $MemberLoginName $RoleDefinitionBindings); 
     
            if ($_Type -eq "Site" -and $MemberType -eq "Group") {
                foreach ($group in $collGroups) {
                    if ($group.Title -eq $MemberName) {
                        $ctx.Load($group.Users);
                        $ctx.ExecuteQuery() ;  
                        ##Write-Host "Number of users" $group.Users.Count;
                        $group.Users| ForEach-Object { 
                            Load-CSOMProperties -object $_ -propertyNames @("LoginName");
                            $ctx.ExecuteQuery() ; 
            
                            $_permissions += (PermissionObject $_object "Site" $_RelativeUrl $_siteUrl $_siteTitle "" "GroupMember" $group.Title $_.Title $_.LoginName $RoleDefinitionBindings); 
                            ##Write-Host  $permissions.Count
                        }
                    }
                }
            } 
        }
      
    }
    return $_permissions;

}

Function QueryUniquePermissions($_web) {
    ##query list, files and items unique permissions
    $permissions = @();
    Write-Host "Querying web " +  $_web.Title  ;
    $siteUrl = $_web.Url; 
 
    $siteRelativeUrl = $_web.ServerRelativeUrl; 
 
    Write-Host $siteUrl -Foregroundcolor "Red"; 
 
    $siteTitle = $_web.Title; 

    Load-CSOMProperties -object $_web -propertyNames @("HasUniqueRoleAssignments");
    $ctx.ExecuteQuery()
    ## See more at: https://www.itunity.com/article/loading-specific-values-lambda-expressions-sharepoint-csom-api-windows-powershell-1249#sthash.2ncW42CM.dpuf
    #Get Site Level Permissions if it's unique  
 
    if ($_web.HasUniqueRoleAssignments -eq $True) { 
        $permissions += (QueryUniquePermissionsByObject $_web $_web "Site" $siteRelativeUrl $siteUrl $siteTitle "");
    }
   
    #Get all lists in web
    $ll = $_web.Lists
    $ctx.Load($ll);
    $ctx.ExecuteQuery()

    Write-Host "Number of lists" + $ll.Count

    foreach ($list in $ll) {      
        Load-CSOMProperties -object $list -propertyNames @("RootFolder", "Hidden", "HasUniqueRoleAssignments");
        $ctx.ExecuteQuery()
 
        $listUrl = $list.RootFolder.ServerRelativeUrl; 
  
        #Exclude internal system lists and check if it has unique permissions 
 
        if ($list.Hidden -ne $True) { 
            Write-Host $list.Title  -Foregroundcolor "Yellow"; 
            $listTitle = $list.Title; 
            #Check List Permissions 

            if ($list.HasUniqueRoleAssignments -eq $True) { 
                $Type = $list.BaseType.ToString(); 
                $permissions += (QueryUniquePermissionsByObject $_web $list $Type $listUrl $siteUrl $siteTitle  $listTitle);
 
                if ($list.BaseType -eq "DocumentLibrary") { 
                    #TODO Get permissions on folders 
                    $rootFolder = $list.RootFolder;
                    $listFolders = $rootFolder.Folders;
                    $ctx.Load($rootFolder);
                    $ctx.Load( $listFolders);
       
                    $ctx.ExecuteQuery() ;
   
                    #get all items 

                    $spQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
                    $spQuery.ViewXml = "<View>
                    <RowLimit>2000</RowLimit>
                </View>"
                    ## array of items
                    $collListItem = @();

                    do {
                        $listItems = $list.GetItems($spQuery);
                        $ctx.Load($listItems);
                        $ctx.ExecuteQuery() ;
                        $spQuery.ListItemCollectionPosition = $listItems.ListItemCollectionPosition
                        foreach ($item in $listItems) {
                            $collListItem += $item 
                        }
                    }
                    while ($spQuery.ListItemCollectionPosition -ne $null)

                    Write-Host  $collListItem.Count 

                    foreach ($item in $collListItem) {
                        Load-CSOMProperties -object $item -propertyNames @("File", "HasUniqueRoleAssignments");
                        $ctx.ExecuteQuery() ;  
        
                        Load-CSOMProperties -object $item.File -propertyNames @("ServerRelativeUrl");
                        $ctx.ExecuteQuery() ;  

                        $fileUrl = $item.File.ServerRelativeUrl; 
 
                        $file = $item.File; 
 
                        if ($item.HasUniqueRoleAssignments -eq $True) { 
                            $Type = $file.GetType().Name; 

                            $permissions += (QueryUniquePermissionsByObject $_web $item $Type $fileUrl $siteUrl $siteTitle $listTitle);
                        } 
                    } 
                } 
            } 
        }
    }
    return  $permissions;
}

if (Test-Path $ExportFileDirectory) {
    Write-Host $Username
    Write-Host $password
 
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteCollectionUrl);
    $ctx.Credentials = New-Object System.Net.NetworkCredential($Username, $password);
  

    $rootWeb = $ctx.Web
    $ctx.Load($rootWeb)
    $ctx.Load($rootWeb.Webs)
    $ctx.ExecuteQuery()

    #Root Web of the Site Collection 

    $RootSiteTitle = $rootWeb.Title; 
 
    $RootWeb = $rootWeb;  
    #array storing permissions
    $Permissions = @(); 

    #root web , i.e. site collection level
    $Permissions += QueryUniquePermissions($RootWeb);
    Write-Host $Permissions.Count;

    Write-Host "Querying Number of webs " $rootWeb.Webs.Count ;  
    foreach ($web in $rootWeb.Webs) {
        $Permissions += (QueryUniquePermissions $web);
        Write-Host "Web :  "  $web.Title "Count" $Permissions.Count
    }

    $exportFilePath = Join-Path -Path $ExportFileDirectory -ChildPath $([string]::Concat($RootSiteTitle, "-Permissions.csv"));
  
    Write-Host $Permissions.Count
 
    $Permissions | Select-Object SiteUrl, SiteTitle, Type, RelativeUrl, ListTitle, MemberType, MemberName, MemberLoginName, ParentGroup, Roles|Export-CSV -Path $exportFilePath -NoTypeInformation;
}
else {
 
    Write-Host "Invalid directory path:" $ExportFileDirectory -ForegroundColor "Red";
 
}
