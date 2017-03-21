[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $false, HelpMessage = 'Optional URL to configure one site only')]
    [Alias("url")]
    [string]$matchURL
)
<#
Office365 - Group Policy

 * leverages 3 libraries (SPO, PNP, CSOM)
 * leverages parallel PowerShell
 * grant Site Collection Admin for support staff
 * apply Site Collection quota 5GB  (if none)
 * enable Site Collection Auditing
 * enable Site Collection Custom Action JS ("office365-gpo.js")
#>

#Core
workflow GPOWorkflow {
    param ($sites, $UserName, $Password)

    Function VerifySite([string]$SiteUrl, $UserName, $Password) {
        Function Get-SPOCredentials([string]$UserName, [string]$Password) {
            if ([string]::IsNullOrEmpty($Password)) {
                $SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString 
            }
            else {
                $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
            }
            return New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
        }
        Function Get-ActionBySequence([Microsoft.SharePoint.Client.ClientContext]$context, [int]$Sequence) {
            $customActions = $context.Site.UserCustomActions
            $context.Load($customActions)
            $context.ExecuteQuery()
            $customActions | where { $_.Sequence -eq $Sequence }
        }
        Function Delete-Action($UserCustomAction) {
            $context = $UserCustomAction.Context
            $UserCustomAction.DeleteObject()
            $context.ExecuteQuery()
            "DELETED"
        }
        Function Verify-ScriptLinkAction([Microsoft.SharePoint.Client.ClientContext]$context, [string]$ScriptSrc, [string]$ScriptBlock, [int]$Sequence) {
            $actions = Get-ActionBySequence -Context $context -Sequence $Sequence
			
            if (!$actions) {
                $action = $context.Site.UserCustomActions.Add();
                $action.Location = "ScriptLink"
                if ($ScriptSrc) {
                    $action.ScriptSrc = $ScriptSrc
                }
                if ($ScriptBlock) {
                    $action.ScriptBlock = $ScriptBlock
                }
                $action.Sequence = $Sequence
                $action.Update()
                $context.ExecuteQuery()
            }
        }
        Function Verify-Features([Microsoft.SharePoint.Client.ClientContext]$context) {	
            #list of Site features
            $feat = $context.Site.Features
            $context.Load($feat)
            $context.ExecuteQuery()
			
            #SPSite - Enable Workflow
            $id = New-Object System.Guid "0af5989a-3aea-4519-8ab0-85d91abe39ff"
            $found = $feat |? {$_.DefinitionId -eq $id}
            if (!$found) {
                $feat.Add($id, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::Farm)
            }
            $context.ExecuteQuery()

            #SPWeb - Disable Minimal Download Strategy (MDS)
            Loop-WebFeature $context $context.Site.RootWeb $false "87294c72-f260-42f3-a41b-981a2ffce37a"
        }
        Function Loop-WebFeature ($context, $currWeb, $wantActive, $featureId) {
            #get parent
            $context.Load($currWeb)
            $context.ExecuteQuery()
			
            #ensure parent
            Ensure-WebFeature $context $currWeb $wantActive $featureId

            #get child
            $webs = $currWeb.Webs
            $context.Load($webs)
            $context.ExecuteQuery()
			
            #loop child subwebs
            foreach ($web in $webs) {
                Write-Host "ensure feature on " + $web.url
                #ensure child
                Ensure-WebFeature $context $web $wantActive $featureId

                #Recurse
                $subWebs = $web.Webs
                $context.Load($subWebs)
                $context.ExecuteQuery()
                $subWebs | ForEach-Object { Loop-WebFeature $context $_ $wantActive $featureId }
            }
        }
        Function Ensure-WebFeature ($context, $web, $wantActive, $featureId) {
            #list of Web features
            if ($web.Url) {
                Write-Host " - $($web.Url)"
                $feat = $web.Features
                $context.Load($feat)
                $context.ExecuteQuery()

                #Disable/Enable Web features
                $id = New-Object System.Guid $featureId
                $found = $feat |? {$_.DefinitionId -eq $id}
                if ($wantActive) {
                    Write-Host "ADD FEAT" -Fore Yellow
                    if (!$found) {
                        $feat.Add($id, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::Farm)
                        $context.ExecuteQuery()
                    }
                }
                else {
                    Write-Host "REMOVE FEAT" -Fore Yellow
                    if ($found) {
                        $feat.Remove($id, $true)
                        $context.ExecuteQuery()
                    }
                }
                #no changes. already OK
            }
        }
        Function Verify-General([Microsoft.SharePoint.Client.ClientContext]$context) {			
            #SPSite
            $site = $context.Site
            $context.Load($site)
            $context.ExecuteQuery()
			
            #RootWeb
            $rootWeb = $site.RootWeb
            $context.Load($rootWeb)
            $context.ExecuteQuery()

            #RootLists
            $rootLists = $rootWeb.Lists
            $context.Load($rootLists)
            $context.ExecuteQuery()
			
            #Access Request List
            $arList = $rootLists.GetByTitle("Access Requests");
            $context.Load($arList)
            $context.ExecuteQuery()
            if ($arList) {
                if ($arList.Hidden) {
                    $arList.Hidden = $false
                    $arList.Update()
                    $context.ExecuteQuery()
                }
            }

            #Trim Audit Log
            if (!$site.TrimAuditLog) {
                $site.TrimAuditLog = $true
                $site.AuditLogTrimmingRetention = 180
                $site.Update()
                $context.ExecuteQuery()
            }
			
            #Audit Log Storage
            if (!$rootWeb.AllProperties["_auditlogreportstoragelocation"]) {
                $url = $context.Site.ServerRelativeUrl
                if ($url -eq "/") {
                    $url = ""
                }
                $rootWeb.AllProperties["_auditlogreportstoragelocation"] = "$url/SiteAssets"
                $rootWeb.Update()
                $context.ExecuteQuery()
            }
        }
		
        #Assembly CSOM
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
				
        try {
            $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
            $cred = Get-SPOCredentials -UserName $UserName -Password $Password
            $context.Credentials = $cred

            #JS CUSTOM ACTION
            $scriptUrl = "https://tenant.sharepoint.com/SiteAssets/Office365-GPO/office365-gpo.js"
            Verify-ScriptLinkAction -Context $context -ScriptSrc $scriptUrl -Sequence 2001

            #SITE COLLECTION CONFIG
            Verify-General -Context $context
			
            #FEATURES
            Verify-Features -Context $context

            $context.Dispose()
        }
        catch {
            Write-Error "ERROR -- $SiteUrl -- $($_.Exception.Message)"
        }
    }


    #Parallel Loop - CSOM
    ForEach -Parallel -ThrottleLimit 100 ($s in $sites) {
        Write-Output "Start thread >> $($s.Url)"
        VerifySite $s.Url $UserName $Password
    }
    "DONE"
}

Function Main {
    #Start
    $start = Get-Date

    #SPO and PNP modules
    Import-Module -WarningAction SilentlyContinue Microsoft.Online.SharePoint.PowerShell -Prefix M
    Import-Module -WarningAction SilentlyContinue SharePointPnPPowerShellOnline -Prefix P
	
    #Config
    $AdminUrl = "https://tenant-admin.sharepoint.com"
    $UserName = "admin@tenant.onmicrosoft.com"
    $Password = "pass@word1"
	
    #Credential
    $secpw = ConvertTo-SecureString -String $Password -AsPlainText -Force
    $c = New-Object System.Management.Automation.PSCredential ($UserName, $secpw)

    #Connect Office 365
    Connect-MSPOService -URL $AdminUrl -Credential $c
	
    #Scope
    Write-Host "Opening list of sites ... $matchURL" -Fore Green
    if ($matchURL) {
        $sites = Get-MSPOSite -Filter "url -like ""$matchURL"""
    }
    else {
        $sites = Get-MSPOSite
    }
    $sites.Count

    #Serial loop
    Write-Host "Serial loop"
    ForEach ($s in $sites) {
        Write-Host "." -NoNewLine
        #SPO
        #Storage quota
        if (!$s.StorageQuota) {
            Set-MSPOSite -Identity $s.Url -StorageQuota 5000 -StorageQuotaWarningLevel 4000
            Write-Output "set 2GB quota on $($s.Url)"
        }

        #Site collection admin
        $scaUser = "SharePoint Service Administrator"
        try {
            $user = Get-MSPOUser -Site $s.Url -Loginname $scaUser -ErrorAction SilentlyContinue
            if (!$user.IsSiteAdmin) {
                Set-MSPOUser -Site $s.Url -Loginname $scaUser -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue | Out-Null
            }
        }
        catch {
        }

        #PNP
        Connect-PSPOnline -Url $s.Url -Credentials $c

        # SharePoint Designer force OFF unless we see "keepspd=1" manual exclude flag ON
        $keepspd = Get-PSPOPropertyBag -Key "keepspd"
        if ($keepspd -ne 1) {
            Set-PSPOPropertyBagValue -Key "allowdesigner" -Value 0
        }

        # Verify Auditing
        $audit = Get-PSPOAuditing
        if ($audit.AuditFlags -ne 7099) {
            Set-PSPOAuditing -RetentionTime 180 -TrimAuditLog -EditItems -CheckOutCheckInItems -MoveCopyItems -DeleteRestoreItems -EditContentTypesColumns -EditUsersPermissions -ErrorAction SilentlyContinue
            Write-Output "set audit flags on $($s.Url)"
        }
    }

   	#Parallel loop
    #CSOM
    Write-Host "Parallel loop"
    GPOWorkflow $sites $UserName $Password

    #Duration
    $min = [Math]::Round(((Get-Date) - $start).TotalMinutes, 2)
    Write-Host "Duration Min : $min"
}
Main