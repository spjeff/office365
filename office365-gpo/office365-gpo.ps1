<#
Office365 - Group Policy

 * leverages 3 libraries (SPO, PNP, CSOM)
 * leverages parallel PowerShell
 * grant Site Collection Admin for support staff
 * apply Site Collection quota 2GB
 * enable Site Collection auditing
 * enable Site Collection Custom Action JS (JQuery + "office365-gpo.js")
 * disable external sharing
#>

# Core
workflow GPOWorkflow {
	Param ($sites, $UserName, $Password)

	Function VerifySite([string]$SiteUrl,[string]$UserName,[string]$Password) {
		Function Get-SPOCredentials([string]$UserName,[string]$Password) {
		   if([string]::IsNullOrEmpty($Password)) {
			  $SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString 
		   }
		   else {
			  $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
		   }
		   return New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
		}
		Function Get-ActionBySequence([Microsoft.SharePoint.Client.ClientContext]$Context,[int]$Sequence) {
			 $customActions = $Context.Site.UserCustomActions
			 $Context.Load($customActions)
			 $Context.ExecuteQuery()
			 $customActions | where { $_.Sequence -eq $Sequence }
		}
		Function Delete-Action([Microsoft.SharePoint.Client.UserCustomAction]$UserCustomAction) {
			 $Context = $UserCustomAction.Context
			 $UserCustomAction.DeleteObject()
			 $Context.ExecuteQuery()
			 "DELETED"
		}
		Function Verify-ScriptLinkAction([Microsoft.SharePoint.Client.ClientContext]$Context,[string]$ScriptSrc,[string]$ScriptBlock, [int]$Sequence) {
			$actions = Get-ActionBySequence -Context $Context -Sequence $Sequence
			
			if (!$actions) {
				$action = $Context.Site.UserCustomActions.Add();
				$action.Location = "ScriptLink"
				if($ScriptSrc) {
					$action.ScriptSrc = $ScriptSrc
				}
				if($ScriptBlock) {
					$action.ScriptBlock = $ScriptBlock
				}
				$action.Sequence = $Sequence
				$action.Update()
				$Context.ExecuteQuery()
				"ADDED"
			}
		}
		Function Verify-General([Microsoft.SharePoint.Client.ClientContext]$Context) {
			#SPSite Collection
			$update = $false
			$site = $Context.Site
			$Context.Load($site)
			$Context.ExecuteQuery()

            if (!$site.TrimAuditLog) {
                $site.TrimAuditLog = $true
				$site.AuditLogTrimmingRetention = 180
				$update = $true
            }
			
			if ($site.ShareByEmailEnabled) {
				$site.ShareByEmailEnabled = $false
				$update = $true
			}
			
			if ($site.ShareByLinkEnabled) {
				$site.ShareByLinkEnabled = $false
				$update = $true
			}
			
			if (!$site.DisableCompanyWideSharingLinks) {
				$site.DisableCompanyWideSharingLinks = $true 
				$update = $true
			}
			
			if ($update) {
				 $Context.ExecuteQuery()
			}
			
			#RootWeb
			$update = $false
			$root = $Context.Site.RootWeb
			$Context.Load($root)
			$Context.ExecuteQuery()
			
			if (!$root.AllProperties["_auditlogreportstoragelocation"]) {
				$url = $Context.Site.ServerRelativeUrl
				if ($url -eq "/") {$url = ""}
				$root.AllProperties["_auditlogreportstoragelocation"] = "$url/SiteAssets"
				$root.Update()
				$Context.ExecuteQuery()
			}
		}
		
		# Assembly CSOM
		[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
				
		Try {
			$context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
			$cred = Get-SPOCredentials -UserName $UserName -Password $Password
			$context.Credentials = $cred


			# JS CUSTOM ACTION
			$scriptUrl = "https://tenant.sharepoint.com/SiteAssets/Office365-GPO/jquery-2.2.3.js"
			Verify-ScriptLinkAction -Context $context -ScriptSrc $scriptUrl -Sequence 2000
			$scriptUrl = "https://tenant.sharepoint.com/SiteAssets/Office365-GPO/office365-gpo.js"
			Verify-ScriptLinkAction -Context $context -ScriptSrc $scriptUrl -Sequence 2001

			# SITE COLLECTION CONFIG
			Verify-General -Context $context

			$context.Dispose()
		} Catch {
			Write-Error "ERROR -- $SiteUrl -- $($_.Exception.Message)"
		}
	}


	#Parallel Loop - CSOM
	ForEach -Parallel -ThrottleLimit 100 ($s in $sites) {
		Write-Output "Start thread >> $($s.Url)"
		VerifySite -SiteUrl $s.Url -UserName $UserName -Password $Password
	}

	"DONE"
}

Function Main {
	# Start
	$start = Get-Date

	#SPO and PNP modules
	Import-Module -WarningAction SilentlyContinue Microsoft.Online.SharePoint.PowerShell -Prefix MS
	Import-Module -WarningAction SilentlyContinue SharePointPnPPowerShellOnline -Prefix PNP
	
	#Config
	$AdminUrl = "https://tenant-admin.sharepoint.com"
	$UserName = "user@tenant.onmicrosoft.com"
	$Password = "pass@word1"
	
	#Credential
	$secpw = ConvertTo-SecureString -String $Password -AsPlainText -Force
	$c = New-Object System.Management.Automation.PSCredential ($UserName, $secpw)

	#Connect Office 365
	Connect-MSSPOService -URL $AdminUrl -Credential $c
	
	#Scope
	Write-Host "Opening list of sites ... " -Fore Green
	$sites = Get-MSSPOSite
	$sites.Count
	
	#Serial loop
    ForEach ($s in $sites) {
		#SPO
        #Storage quota
        if (!$s.StorageQuota) {
            Set-MSSPOSite -Identity $s.Url -StorageQuota 2000 -StorageQuotaWarningLevel 1900
            Write-Output "set 2GB quota on $($s.Url)"
        }

        #Site collection admin
        Set-MSSPOUser -site $s.Url -Loginname $UserName -IsSiteCollectionAdmin $true

		#PNP
        Connect-PNPSPOnline -Url $s.Url -Credentials $c
        $audit = Get-PNPSPOAuditing
        if ($audit.AuditFlags -ne 7099) {
            Set-PNPSPOAuditing -RetentionTime 180 -TrimAuditLog -EditItems -CheckOutCheckInItems -MoveCopyItems -DeleteRestoreItems -EditContentTypesColumns -EditUsersPermissions
            Write-Output "set audit flags on $($s.Url)"
        }
    }

   	#Parallel loop
	#CSOM
	GPOWorkflow $sites $UserName $Password

	#Duration
	[Math]::Round(((Get-Date) - $start).TotalMinutes, 2)
}
Main