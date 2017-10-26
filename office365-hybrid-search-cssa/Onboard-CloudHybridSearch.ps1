<#
.SYNOPSIS
    When you run this script you onboard your SharePoint Online (SPO) tenant and your SharePoint server cloud SSA to cloud hybrid search.
    This includes setting up server to server authentication between SharePoint Online and SharePoint Server
.PARAMETER PortalUrl
    SharePoint Online portal URL, for example 'https://tenant.sharepoint.com'.
.PARAMETER CloudSsaId
    Name or id (Guid) of the cloud Search service application, created with the CreateCloudSSA script.
.PARAMETER Credential
    Logon credential for tenant admin. Will prompt for credential if not specified.
#>
Param(
    [Parameter(Mandatory=$true, HelpMessage="SharePoint Online portal URL, for example 'https://tenant.sharepoint.com'.")]
    [ValidateNotNullOrEmpty()]
    [string] $PortalUrl,

    [Parameter(Mandatory=$false, HelpMessage="Name or id (Guid) of the cloud Search service application, created with the CreateCloudSSA script.")]
    [ValidateNotNullOrEmpty()]
    [string] $CloudSsaId,
    
    [Parameter(Mandatory=$false, HelpMessage="Logon credential for tenant admin. Will be prompted if not specified.")]
    [PSCredential] $Credential
)
Start-Transcript
(Get-Date)
if ($ACS_APPPRINCIPALID -eq $null) {
    New-Variable -Option Constant -Name ACS_APPPRINCIPALID -Value '00000001-0000-0000-c000-000000000000'
    New-Variable -Option Constant -Name ACS_HOST -Value "accounts.accesscontrol.windows.net"
    New-Variable -Option Constant -Name PROVISIONINGAPI_WEBSERVICEURL -Value "https://provisioningapi.microsoftonline.com/provisioningwebservice.svc"
    New-Variable -Option Constant -Name SCS_AUTHORITIES -Value @(
        "*.search.msit.us.trafficmanager.net",
        "*.search.production.us.trafficmanager.net",
        "*.search.production.emea.trafficmanager.net",
        "*.search.production.apac.trafficmanager.net"
    )
}

New-Variable -Option Constant -Name SCS_APPPRINCIPALID -Value '8f0dc9ad-0d19-4fec-a421-6d0279080014'
New-Variable -Option Constant -Name SCS_APPPRINCIPALDISPLAYNAME -Value 'Search Content Service'
New-Variable -Option Constant -Name SP_APPPRINCIPALID -Value '00000003-0000-0ff1-ce00-000000000000'
New-Variable -Option Constant -Name SPO_MANAGEMENT_APPPROXY_NAME -Value 'SPO App Management Proxy'
New-Variable -Option Constant -Name ACS_APPPROXY_NAME -Value 'ACS'
New-Variable -Option Constant -Name ACS_STS_NAME -Value 'ACS-STS'
New-Variable -Option Constant -Name AAD_METADATAEP_FSTRING -Value 'https://{0}/{1}/metadata/json/1'

$SP_VERSION = "15"
$regKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office Server\15.0\Search" -ErrorAction SilentlyContinue
if ($regKey -eq $null) {
    $regKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office Server\16.0\Search" -ErrorAction SilentlyContinue
    if ($regKey -eq $null) {
        throw "Unable to detect SharePoint installation."
    }
    $SP_VERSION = "16"
}

Write-Host "Configuring for SharePoint Server version $SP_VERSION."

function Configure-LocalSharePointFarm {
    Param(
        [Parameter(Mandatory=$true)][string] $Realm
    )

    # Set up to authenticate as AAD realm
    Set-SPAuthenticationRealm -Realm $Realm

    $acsMetadataEndpoint = $AAD_METADATAEP_FSTRING -f $ACS_HOST,$Realm
    $acsMetadataEndpointUri = [System.Uri] $acsMetadataEndpoint
    $acsMetadataEndpointUriSlash = [System.Uri] "$($acsMetadataEndpoint)/"
    Write-Host "ACS metatada endpoint: $acsMetadataEndpoint"

    # ACS Proxy
    $acsProxy = Get-SPServiceApplicationProxy | ? {
        $_.TypeName -eq "Azure Access Control Service Application Proxy" -and
        (($_.MetadataEndpointUri -eq $acsMetadataEndpointUri) -or ($_.MetadataEndpointUri -eq $acsMetadataEndpointUriSlash))
    }
    if ($acsProxy -eq $null) {
        Write-Host "Setting up ACS proxy..." -Foreground Yellow
        $acsProxy = Get-SPServiceApplicationProxy | ? {$_.DisplayName -eq $ACS_APPPROXY_NAME}
        if ($acsProxy -ne $null) {
            throw "There is already a service application proxy registered with name '$($acsProxy.DisplayName)'. Remove manually and retry."
        }
        $acsProxy = New-SPAzureAccessControlServiceApplicationProxy -Name $ACS_APPPROXY_NAME -MetadataServiceEndpointUri $acsMetadataEndpointUri -DefaultProxyGroup
    } elseif ($acsProxy.Count > 1) {
        throw "Found multiple existing ACS proxies for this metadata endpoint."
    } else {
        Write-Host "Found existing ACS proxy '$($acsProxy.DisplayName)'." -Foreground Green
    }

    # The proxy must be in default group and set as default for authentication to work
    if (((Get-SPServiceApplicationProxyGroup -Default).DefaultProxies | select Id).Id -notcontains $acsProxy.Id) {
        throw "ACS proxy '$($acsProxy.DisplayName)' is not set as the default. Configure manually through Service Application Associations admin UI and retry."
    }

    # Register ACS token issuer
    $acsTokenIssuer = Get-SPTrustedSecurityTokenIssuer | ? {
        (($_.MetadataEndPoint -eq $acsMetadataEndpointUri) -or ($_.MetadataEndPoint -eq $acsMetadataEndpointUriSlash))
    }
    if ($acsTokenIssuer -eq $null) {
        Write-Host "Registering ACS as trusted token issuer..." -Foreground Yellow
        $acsTokenIssuer = Get-SPTrustedSecurityTokenIssuer | ? {$_.DisplayName -eq $ACS_STS_NAME}
        if ($acsTokenIssuer -ne $null) {
            throw "There is already a token issuer registered with name '$($acsTokenIssuer.DisplayName)'. Remove manually and retry."
        }
        try {
            $acsTokenIssuer = New-SPTrustedSecurityTokenIssuer -Name $ACS_STS_NAME -IsTrustBroker -MetadataEndPoint $acsMetadataEndpointUri -ErrorAction Stop
        } catch [System.ArgumentException] {
            Write-Warning "$($_)"
        }
    } elseif ($acsTokenIssuer.Count > 1) {
        throw "Found multiple existing token issuers for this metadata endpoint."
    } else {
        if ($acsTokenIssuer.IsSelfIssuer -eq $true) {
            Write-Warning "Existing trusted token issuer '$($acsTokenIssuer.DisplayName)' is configured as SelfIssuer."
        } else {
            Write-Host "Found existing token issuer '$($acsTokenIssuer.DisplayName)'." -Foreground Green
        }
    }

    # SPO proxy
    $spoProxy = Get-SPServiceApplicationProxy | ? {$_.TypeName -eq "SharePoint Online Application Principal Management Service Application Proxy" -and $_.OnlineTenantUri -eq [System.Uri] $PortalUrl}
    if ($spoProxy -eq $null) {
        Write-Host "Setting up SPO Proxy..." -Foreground Yellow
        $spoProxy = Get-SPServiceApplicationProxy | ? {$_.DisplayName -eq $SPO_MANAGEMENT_APPPROXY_NAME}
        if ($spoProxy -ne $null) {
            throw "There is already a service application proxy registered with name '$($spoProxy.DisplayName)'. Remove manually and retry."
        }
        $spoProxy = New-SPOnlineApplicationPrincipalManagementServiceApplicationProxy -Name $SPO_MANAGEMENT_APPPROXY_NAME -OnlineTenantUri $PortalUrl -DefaultProxyGroup
    } elseif ($spoProxy.Count > 1) {
        throw "Found multiple existing SPO proxies for this tenant URI."
    } else {
        Write-Host "Found existing SPO proxy '$($spoProxy.DisplayName)'." -Foreground Green
    }

    # The proxy should be in default group and set to default
    if (((Get-SPServiceApplicationProxyGroup -Default).DefaultProxies | select Id).Id -notcontains $spoProxy.Id) {
        throw "SPO proxy '$($spoProxy.DisplayName)' is not set as the default. Configure manually through Service Application Associations admin UI and retry."
    }

    return (Get-SPSecurityTokenServiceConfig).LocalLoginProvider.SigningCertificate
}

function Upload-SigningCredentialToSharePointPrincipal {
    Param(
        [Parameter(Mandatory=$true)][System.Security.Cryptography.X509Certificates.X509Certificate2] $Cert
    )

    $exported = $Cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $certValue = [System.Convert]::ToBase64String($exported)

    $principal = Get-MsolServicePrincipal -AppPrincipalId $SP_APPPRINCIPALID
    $keys = Get-MsolServicePrincipalCredential -ObjectId $principal.ObjectId -ReturnKeyValues $true | ? Value -eq $certValue
    if ($keys -eq $null) {
        New-MsolServicePrincipalCredential -AppPrincipalId $SP_APPPRINCIPALID -Type Asymmetric -Value $certValue -Usage Verify
    } else {
        Write-Host "Signing credential already exists in SharePoint principal."
    }
}

function Add-ScsServicePrincipal {
    $spns = $SCS_AUTHORITIES | foreach { "$SCS_APPPRINCIPALID/$_" }
    $principal = Get-MsolServicePrincipal -AppPrincipalId $SCS_APPPRINCIPALID -ErrorAction SilentlyContinue

    if ($principal -eq $null) {
        Write-Host "Creating new service principal for $SCS_APPPRINCIPALDISPLAYNAME with the following SPNs:"
        $spns | foreach { Write-Host $_ }
        $scspn = New-MsolServicePrincipal -AppPrincipalId $SCS_APPPRINCIPALID -DisplayName $SCS_APPPRINCIPALDISPLAYNAME -ServicePrincipalNames $spns
    } else {
        $update = $false
        $spns | foreach {
            if ($principal.ServicePrincipalNames -notcontains $_) {
                $principal.ServicePrincipalNames.Add($_)
                Write-Host "Adding new SPN to existing service principal: $_."
                $update = $true
            }
        }
        if ($update -eq $true) {
            Set-MsolServicePrincipal -AppPrincipalId $principal.AppPrincipalId -ServicePrincipalNames $principal.ServicePrincipalNames
        } else {
            Write-Host "Service Principal already registered, containing the correct SPNs."
        }
    }
}

function Prepare-Environment {
    $MSOIdCRLRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL" -ErrorAction SilentlyContinue
    if ($MSOIdCRLRegKey -eq $null) {
        Write-Host "Online Services Sign-In Assistant required, install from http://www.microsoft.com/en-us/download/details.aspx?id=39267." -Foreground Red
    } else {
        Write-Host "Found Online Services Sign-In Assistant!" -Foreground Green
    }

    $MSOLPSRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOnlinePowershell" -ErrorAction SilentlyContinue
    if ($MSOLPSRegKey -eq $null) {
        Write-Host "AAD PowerShell required, install from http://go.microsoft.com/fwlink/p/?linkid=236297." -Foreground Red
    } else {
        Write-Host "Found AAD PowerShell!" -Foreground Green
    }

    if ($MSOIdCRLRegKey -eq $null -or $MSOLPSRegKey -eq $null) {
        throw "Manual installation of prerequisites required."
    }

    Write-Host "Configuring Azure AD settings..." -Foreground Yellow

    $regkey = "HKLM:\SOFTWARE\Microsoft\MSOnlinePowerShell\Path"
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL" -Name "ServiceEnvironment" -Value "Production"
    Set-ItemProperty -Path $regkey -Name "WebServiceUrl" -Value $PROVISIONINGAPI_WEBSERVICEURL
    Set-ItemProperty -Path $regkey -Name "FederationProviderIdentifier" -Value "microsoftonline.com"

    Write-Host "Restarting MSO IDCRL Service..." -Foreground Yellow

    # Service takes time to get provisioned, retry restart.
    for ($i = 1; $i -le 10; $i++) {
        try {
            Stop-Service -Name msoidsvc -Force -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            $svc = Get-Service msoidsvc
            $svc.WaitForStatus("Stopped")
            Start-Service -Name msoidsvc
        } catch {
            Write-Host "Failed to start msoidsvc service, retrying..."
            Start-Sleep -seconds 2
            continue
        }
        Write-Host "Service Restarted!" -Foreground Green
        break
    }
}

function Get-CloudSsa {
    $ssa = $null

    if (-not $CloudSsaId) {
        $ssa = Get-SPEnterpriseSearchServiceApplication
        if ($ssa.Count -ne 1) {
            throw "Multiple SSAs found, specify which cloud SSA to on-board."
        }
    } else {
        $ssa = Get-SPEnterpriseSearchServiceApplication -Identity $CloudSsaId
    }

    if ($ssa -eq $null) {
        throw "Cloud SSA not found."
    }

    # Make sure SSA is created with CreateCloudSSA.ps1
    if ($ssa.CloudIndex -ne $true) {
        throw "The provided SSA is not set up for cloud hybrid search, please create a cloud SSA before proceeding with onboarding."
    }

    Write-Host "Using SSA with id $($ssa.Id)."
    $ssa.SetProperty("IsHybrid", 1)
    $ssa.Update()

    return $ssa
}

$code = @"
using System;
using System.Net;
using System.Security;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.IdentityModel;
using Microsoft.SharePoint.IdentityModel.OAuth2;

static public class ClientContextHelper
{
    public static ClientContext GetAppClientContext(string siteUrl)
    {
        SPServiceContext serviceContext = SPServiceContext.GetContext(SPServiceApplicationProxyGroup.Default, SPSiteSubscriptionIdentifier.Default);
        using (SPServiceContextScope serviceContextScope = new SPServiceContextScope(serviceContext))
        {
            ClientContext clientContext = new ClientContext(siteUrl);
            ICredentials credentials = null;
            clientContext.ExecutingWebRequest += (sndr, request) =>
            {
                    request.WebRequestExecutor.RequestHeaders.Add(HttpRequestHeader.Authorization, "Bearer");
                    request.WebRequestExecutor.WebRequest.PreAuthenticate = true;
            };

            // Run elevated to get app credentials
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
               credentials = SPOAuth2BearerCredentials.Create();
            });

            clientContext.Credentials = credentials;

            return clientContext;
        }
    }
}
"@

$assemblies = @(
    "System.Core.dll",
    "System.Web.dll",
    "Microsoft.SharePoint, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
    "Microsoft.SharePoint.Client, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
    "Microsoft.SharePoint.Client.Runtime, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
)

Add-Type -AssemblyName ("Microsoft.SharePoint.Client, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
Add-Type -AssemblyName ("Microsoft.SharePoint.Client.Search, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
Add-Type -AssemblyName ("Microsoft.SharePoint.Client.Runtime, Version=$SP_VERSION.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
Add-Type -TypeDefinition $code -ReferencedAssemblies $assemblies

Add-PSSnapin Microsoft.SharePoint.PowerShell

try {
    Write-Host "Accessing Cloud SSA..." -Foreground Yellow
    $ssa = Get-CloudSsa

    Write-Host "Preparing environment..." -Foreground Yellow
    Prepare-Environment

    Import-Module MSOnline
    Import-Module MSOnlineExtended -force

    Write-Host "Connecting to O365..." -Foreground Yellow
    if ($Credential -eq $null) {
        $Credential = Get-Credential -Message "Tenant Admin credential"
    }
    Connect-MsolService -Credential $Credential -ErrorAction Stop
    $tenantInfo = Get-MsolCompanyInformation
    $AADRealm = $tenantInfo.ObjectId.Guid

    Write-Host "AAD tenant realm is $AADRealm."

    Write-Host "Configuring on-prem SharePoint farm..." -Foreground Yellow
    $signingCert = Configure-LocalSharePointFarm -Realm $AADRealm
    
    Write-Host "Adding local signing credential to SharePoint principal..." -Foreground Yellow
    Upload-SigningCredentialToSharePointPrincipal -Cert $signingCert

    Write-Host "Configuring service principal for the cloud search service..." -Foreground Yellow
    Add-ScsServicePrincipal

    Write-Host "Connecting to content farm in SPO..." -foreground Yellow
    $cctx = [ClientContextHelper]::GetAppClientContext($PortalUrl)
    Write-Host "Context:"
    $cctx 
    $pushTenantManager = new-object Microsoft.SharePoint.Client.Search.ContentPush.PushTenantManager $cctx
    Write-Host "PushTenantManager:"
    $pushTenantManager 
	
    # Retry up to 4 minutes, mitigate 401 Unauthorized from CSOM
    Write-Host "Preparing tenant for cloud hybrid search (this can take a couple of minutes)..." -foreground Yellow
    for ($i = 1; $i -le 12; $i++) {
        try {
            $pushTenantManager.PreparePushTenant()
            $cctx.ExecuteQuery()
            Write-Host "PreparePushTenant was successfully invoked!" -Foreground Green
            break
        } catch {
            if ($i -ge 12) {
                throw "Failed to call PreparePushTenant, error was $($_.Exception.Message)"
            }
            Start-Sleep -seconds 20
        }
    }

    Write-Host "Getting service info..." -foreground Yellow
    $info = $pushTenantManager.GetPushServiceInfo()
    $info.Retrieve("EndpointAddress")
    $info.Retrieve("TenantId")
    $info.Retrieve("AuthenticationRealm")
    $info.Retrieve("ValidContentEncryptionCertificates")
    $cctx.ExecuteQuery()

    Write-Host "Registered cloud hybrid search configuration:"
    $info | select TenantId,AuthenticationRealm,EndpointAddress | format-list

    if ([string]::IsNullOrEmpty($info.EndpointAddress)) {
        throw "No indexing service endpoint found!"
    }

    if ($info.ValidContentEncryptionCertificates -eq $null) {
        Write-Warning "No valid encryption certificate found."
    }

    if ($AADRealm -ne $info.AuthenticationRealm) {
        throw "Unexpected mismatch between realm ids read from Get-MsolCompanyInformation ($AADRealm) and GetPushServiceInfo ($($info.AuthenticationRealm))."
    }

    Write-Host "Configuring Cloud SSA..." -foreground Yellow
    $ssa.SetProperty("CertServerURL", $PortalUrl)
    $ssa.SetProperty("HybridTenantID", $info.TenantId)
    $ssa.SetProperty("AuthRealm", $info.AuthenticationRealm)
    $ssa.Update()

    Write-Host "Restarting SharePoint Timer Service..." -foreground Yellow
    Stop-Service SPTimerV4
    Write-Host "Restarting SharePoint Server Search..." -foreground Yellow
    if ($SP_VERSION -eq "15") {
        Restart-Service OSearch15
    } else {
        Restart-Service OSearch16
    }
    Start-Service SPTimerV4

    Write-Host "All done!" -foreground Green
}
catch {
    Write-Error -ErrorRecord $_
    Write-Host "It is safe to re-run onboarding if you believe this error is transient." -Foreground Yellow
    return
}
Stop-Transcript