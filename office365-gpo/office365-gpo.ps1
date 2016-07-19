<#
Office365 - Group Policy

 * leverages 3 libraries (SPO, PNP, CSOM)
 * leverages parallel PowerShell
 * grant Site Collection Admin for support staff
 * apply Site Collection quota 2GB
 * enable Site Collection auditing
 * enable Site Collection Custom Action JS (JQuery + "office365-gpo.js")
 
 * last updated 07-19-16
#>

#Core
workflow GPOWorkflow { param ($sites, $UserName, $Password)

	Function VerifySite([string]$SiteUrl, $UserName, $Password) {
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
			}
		}
		Function Verify-Features([Microsoft.SharePoint.Client.ClientContext]$Context) {	
			#list of Site features
            $feat = $Context.Site.Features
            $Context.Load($feat)
			$Context.ExecuteQuery()
			
			#SPSite - Enable Workflow
			$id = New-Object System.Guid "0af5989a-3aea-4519-8ab0-85d91abe39ff"
			$found = $feat |? {$_.DefinitionId -eq $id}
			if (!$found) {
				$feat.Add($id, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::Farm)
			}
            $Context.ExecuteQuery()

            #SPWeb - Disable Minimal Download Strategy (MDS)
            Loop-WebFeature $Context $Context.Site.RootWeb $false "87294c72-f260-42f3-a41b-981a2ffce37a"
		}
        Function Loop-WebFeature ($Context, $currWeb, $wantActive, $featureId) {
            #get parent
            $Context.Load($currWeb)
            $Context.ExecuteQuery()
			
			#ensure parent
            Ensure-WebFeature $Context $currWeb $wantActive $featureId

            #get child
            $webs = $currWeb.Webs
            $Context.Load($webs)
            $Context.ExecuteQuery()
			
			#loop child subwebs
            foreach ($web in $webs) {
				Write-Host "ensure feature on " + $web.url
                #ensure child
                Ensure-WebFeature $Context $web $wantActive $featureId

                #Recurse
                $subWebs = $web.Webs
                $Context.Load($subWebs)
                $Context.ExecuteQuery()
                $subWebs | ForEach-Object { Loop-WebFeature $Context $_ $wantActive $featureId }
            }
        }
        Function Ensure-WebFeature ($Context, $web, $wantActive, $featureId) {
            #list of Web features
            if ($web.Url) {
                Write-Host " - $($web.Url)"
			    $feat = $web.Features
			    $Context.Load($feat)
			    $Context.ExecuteQuery()

                #Disable/Enable Web features
                $id = New-Object System.Guid $featureId
                $found = $feat |? {$_.DefinitionId -eq $id}
                if ($wantActive) {
					Write-Host "ADD FEAT" -Fore Yellow
                    if (!$found) {
						$feat.Add($id, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::Farm)
						$Context.ExecuteQuery()
					}
                } else {
					Write-Host "REMOVE FEAT" -Fore Yellow
			        if ($found) {
						$feat.Remove($id, $true)
						$Context.ExecuteQuery()
					}
                }
				#no changes. already OK
            }
        }
		Function Verify-General([Microsoft.SharePoint.Client.ClientContext]$Context) {
			#Defaults
			$update = $false
			
			#SPSite
			$site = $Context.Site
			$Context.Load($site)
			$Context.ExecuteQuery()
			
			#SPWeb
			$rootWeb = $site.RootWeb
			$Context.Load($rootWeb)
			$Context.ExecuteQuery()
			
			#Access Request SPList
			<#
			$arList = $rootWeb.Lists.GetByTitle("Access Requests");
			$Context.Load($arList)
			$Context.ExecuteQuery()
			if ($arList) {
				$arList.Hidden = $false
				$arList.Update()
				$update = $true
			}
			#>

			#Trim Audit Log
            if (!$site.TrimAuditLog) {
                $site.TrimAuditLog = $true
				$site.AuditLogTrimmingRetention = 180
				$update = $true
            }
			
			#Audit Log Storage
			if (!$rootWeb.AllProperties["_auditlogreportstoragelocation"]) {
				$url = $Context.Site.ServerRelativeUrl
				if ($url -eq "/") {$url = ""}
				$rootWeb.AllProperties["_auditlogreportstoragelocation"] = "$url/SiteAssets"
				$rootWeb.Update()
				$update = $true
			}
			
			#Update
			if ($update) {
				 $Context.ExecuteQuery()
			}
		}
		
		#Assembly CSOM
		[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
				
		Try {
			$context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
			$cred = Get-SPOCredentials -UserName $UserName -Password $Password
			$context.Credentials = $cred

			#JS CUSTOM ACTION
			$scriptUrl = "https://tenant.sharepoint.com/SiteAssets/Office365-GPO/jquery-2.2.3.js"
			Verify-ScriptLinkAction -Context $context -ScriptSrc $scriptUrl -Sequence 2000
			$scriptUrl = "https://tenant.sharepoint.com/SiteAssets/Office365-GPO/office365-gpo.js"
			Verify-ScriptLinkAction -Context $context -ScriptSrc $scriptUrl -Sequence 2001

			#SITE COLLECTION CONFIG
			Verify-General -Context $context
			
			#FEATURES
			Verify-Features -Context $context

			$context.Dispose()
		} Catch {
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
	Write-Host "Opening list of sites ... " -Fore Green
	$sites = Get-MSPOSite
	$sites.Count

	#Serial loop
    Write-Host "Serial loop"
    ForEach ($s in $sites) {
        Write-Host "." -NoNewLine
		#SPO
        #Storage quota
        if (!$s.StorageQuota) {
            Set-MSPOSite -Identity $s.Url -StorageQuota 2000 -StorageQuotaWarningLevel 1900
            Write-Output "set 2GB quota on $($s.Url)"
        }

        #Site collection admin
		$scaUser = "SharePoint Service Administrator"
        $user = Get-MSPOUser -Site $s.Url -Loginname $scaUser -ErrorAction SilentlyContinue
        if (!$user.IsSiteAdmin) {
            Set-MSPOUser -Site $s.Url -Loginname $scaUser -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue | Out-Null
        }

		#PNP
        Connect-PSPOnline -Url $s.Url -Credentials $c
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