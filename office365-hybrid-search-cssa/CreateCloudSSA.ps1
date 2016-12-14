## Gather mandatory parameters ##     
## Note: SearchServiceAccount needs to already exist in Windows Active Directory as per Technet Guidelines https://technet.microsoft.com/library/gg502597.aspx ##     
Param(     
[Parameter(Mandatory=$true)][string] $SearchServerName,       
[Parameter(Mandatory=$true)][string] $SearchServiceAccount,      
[Parameter(Mandatory=$true)][string] $SearchServiceAppName,      
[Parameter(Mandatory=$true)][string] $DatabasePrefix,
[Parameter(Mandatory=$true)][string] $DatabaseServerName      
)
Start-Transcript
Add-PSSnapin Microsoft.SharePoint.Powershell -ea 0      
## Validate if the supplied account exists in Active Directory and whether supplied as domain\username     
    if ($SearchServiceAccount.Contains("\")) # if True then domain\username was used      
    {      
    $Account = $SearchServiceAccount.Split("\")      
$Account = $Account[1]      
    }      
    else # no domain was specified at account entry      
    {      
    $Account = $SearchServiceAccount      
    }      
    
    $domainRoot = New-Object System.DirectoryServices.DirectoryEntry
    $dirSearcher = New-Object System.DirectoryServices.DirectorySearcher($domainRoot)      
    $dirSearcher.filter = "(&(objectClass=user)(sAMAccountName=$Account))"
    $results = $dirSearcher.findall()      
if ($results.Count -gt 0) # Test for user not found      
    {       
 
    Write-Output "Active Directory account $Account exists. Proceeding with configuration"
## Validate whether the supplied SearchServiceAccount is a managed account. If not make it one.     
if(Get-SPManagedAccount | ?{$_.username -eq $SearchServiceAccount})       
    {      
        Write-Output "Managed account $SearchServiceAccount already exists!"      
    }      
    else      
{      
        Write-Output "Managed account does not exists – creating it"      
$ManagedCred = Get-Credential -Message "Please provide the password for $SearchServiceAccount" -UserName $SearchServiceAccount      
        try      
        {      
        New-SPManagedAccount -Credential $ManagedCred      
        }      
        catch      
        {      
         Write-Output "Unable to create managed account for $SearchServiceAccount. Please validate user and domain details"      
         break      
         }  
    }      
Write-Output "Creating Application Pool"       
$appPoolName=$SearchServiceAppName+"_AppPool"     
$appPool = Get-SPServiceApplicationPool $appPoolName -ErrorAction SilentlyContinue
if (!$appPool) {
    $appPool = New-SPServiceApplicationPool -name $appPoolName -account $SearchServiceAccount      
	Write-Output "APPPOOL created - $appPoolName"
} else {
	Write-Output "APPPOOL found - $appPoolName"
}
Write-Output "Starting Search Service Instance"      
Start-SPEnterpriseSearchServiceInstance $SearchServerName      
 
Write-Output "Creating Cloud Search Service Application" 
 
$searchApp = New-SPEnterpriseSearchServiceApplication -Name $SearchServiceAppName -ApplicationPool $appPool -DatabaseName "$DatabasePrefix$SearchServiceAppName" -DatabaseServer $DatabaseServerName -CloudIndex $true      
Write-Output "Configuring Admin Component"      
$searchInstance = Get-SPEnterpriseSearchServiceInstance $SearchServerName      
$searchApp | get-SPEnterpriseSearchAdministrationComponent | set-SPEnterpriseSearchAdministrationComponent -SearchServiceInstance $searchInstance      
$admin = ($searchApp | get-SPEnterpriseSearchAdministrationComponent)      
Write-Output "Waiting for the admin component to be initialized"      
$timeoutTime=(Get-Date).AddMinutes(20)      
do {Write-Output .;Start-Sleep 10;} while ((-not $admin.Initialized) -and ($timeoutTime -ge (Get-Date)))      
if (-not $admin.Initialized) { throw "Admin Component could not be initialized"}      
 
Write-Output "Inspecting Cloud Search Service Application" 
  
$searchApp = Get-SPEnterpriseSearchServiceApplication $SearchServiceAppName      
 
Write-Output "Setting IsHybrid Property to 1"      
$searchApp.SetProperty("IsHybrid",1)      
#Output some key properties of the Search Service Application     
  
Write-Host "Search Service Properties"       
Write-Host "Hybrid Cloud SSA Name    : " $searchapp.Name      
Write-Host "Hybrid Cloud SSA Status  : " $searchapp.Status      
Write-Host "Cloud Index Enabled      : " $searchApp.CloudIndex      
Write-Output "Configuring Search Topology"      
$searchApp = Get-SPEnterpriseSearchServiceApplication $SearchServiceAppName      
$topology = $searchApp.ActiveTopology.Clone()      
 
$oldComponents = @($topology.GetComponents()) 
 
if (@($oldComponents | ? { $_.GetType().Name -eq "AdminComponent" }).Length -eq 0) 
 
{ 
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.AdminComponent $SearchServerName))      
}     
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.CrawlComponent $SearchServerName))      
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.ContentProcessingComponent $SearchServerName))      
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.AnalyticsProcessingComponent $SearchServerName))      
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.QueryProcessingComponent $SearchServerName))      
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.IndexComponent $SearchServerName,0))      
 
$oldComponents  | ? { $_.GetType().Name -ne "AdminComponent" } | foreach { $topology.RemoveComponent($_) } 
Write-Output "Activating topology"      
$topology.Activate()     
$timeoutTime=(Get-Date).AddMinutes(20)      
do {Write-Output .;Start-Sleep 10;} while (($searchApp.GetTopology($topology.TopologyId).State -ne "Active") -and ($timeoutTime -ge (Get-Date)))      
if ($searchApp.GetTopology($topology.TopologyId).State -ne "Active")  { throw 'Could not activate the search topology'}      
Write-Output "Creating Proxy"      
$searchAppProxy = new-spenterprisesearchserviceapplicationproxy -name ($SearchServiceAppName+"_proxy") -SearchApplication $searchApp      
Write-Output " Cloud hybrid search service application provisioning completed successfully."      
    }      
    else # The Account Must Exist so we can proceed with the script      
    {      
    Write-Output "Account supplied for Search Service does not exist in Active Directory."      
    Write-Output "Script is quitting. Please create the account and run again."      
 
    Break 
} # End Else   
Stop-Transcript