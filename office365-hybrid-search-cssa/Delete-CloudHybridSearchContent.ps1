<#
.SYNOPSIS
    Issue a call to SharePoint Online to delete all metadata from on-premises content that was
    indexed through cloud hybrid search. This operation is asynchronous.
.PARAMETER PortalUrl
    SharePoint Online portal URL, for example 'https://contoso.sharepoint.com'.
.PARAMETER Credential
    Logon credential for tenant admin. Will prompt for credential if not specified.
#>
param(
    [Parameter(Mandatory=$true, HelpMessage="SharePoint Online portal URL, for example 'https://contoso.sharepoint.com'.")]
    [ValidateNotNullOrEmpty()]
    [String] $PortalUrl,

    [Parameter(Mandatory=$false, HelpMessage="Logon credential for tenant admin. Will be prompted if not specified.")]
    [PSCredential] $Credential
)

$SP_VERSION = "15"
$regKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office Server\15.0\Search" -ErrorAction SilentlyContinue
if ($regKey -eq $null) {
    $regKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office Server\16.0\Search" -ErrorAction SilentlyContinue
    if ($regKey -eq $null) {
        throw "Unable to detect SharePoint Server installation."
    }
    $SP_VERSION = "16"
}

Add-Type -AssemblyName ("Microsoft.SharePoint.Client, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
Add-Type -AssemblyName ("Microsoft.SharePoint.Client.Search, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
Add-Type -AssemblyName ("Microsoft.SharePoint.Client.Runtime, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")

if ($Credential -eq $null)
{
    $Credential = Get-Credential -Message "SharePoint Online tenant admin credential"
}

$context = New-Object Microsoft.SharePoint.Client.ClientContext($PortalUrl)
$spocred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)
$context.Credentials = $spocred

$manager = New-Object Microsoft.SharePoint.Client.Search.ContentPush.PushTenantManager $context
$task = $manager.DeleteAllCloudHybridSearchContent()
$context.ExecuteQuery()

Write-Host "Started delete task (id=$($task.Value))"
