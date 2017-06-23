# ==============================================================================================
#    NAME: o365-auto-license.ps1
#    Initial Alpha Date: 12/15/2015
#    Released Version Date: 12/17/2015
#    Last Modification Date: 5/10/2016
#    Released Version: 0.8
#    COMMENT: This script should configured as a scheduled task to automatically assign Office 365
#			  users a license based on any filter required. It will not automatically remove licenses,
#			  it will in fact ensure that any currently assigned licenses do not get disabled which 
#			  could cause loss of data. Ensure the filters for users are adjusted properly before
#			  putting into operation.
#    
#  
# ==============================================================================================

##################
# Misc variables #
##################

# Set date, age limit, & path for reports
$Date = Get-Date -Format MM-dd-yyyy-HHmm 
$ReportAgeLimit = -45
$ReportPath = "D:\EnableUsersReports"
$TranscriptPath = "D:\ScriptLogs\ScriptLog" + $Date + ".log"
$Script:ScriptDC1 = "DC1.company.com"
$Script:ScriptDC2 = "DC2.company.com"

#Create array to put users in for reporting
$Script:AllUsers = @()

# Set usage location for license assignments
$UsageLocation = "US"

#################################
# License email alert variables #
#################################
$Script:SMTPServer = "mail.company.com"
[array]$Script:SMTPTOAddresses = "support@company.com"
$Script:SMTPFROMAddress = "no-reply@company.com"
$Script:SMTPSubect = "URGENT: Office 365 Available Licenses in Acceptance Tenant"
$Script:SoftAvailLicThreshold = 25
$Script:SoftAvailLicPercent = .05
$Script:HardAvailLicThreshold = 2
#This value is used to determine what number of active units for a sku will be checked for threshold rather than percent. 
#E.g. if the value of the below is 1000, the script will check for all sku's with over 1000 active units (licenses)
#and check to see if those skus are below the SoftAvailLicThreshold variable. Active units lower than 1000 will be checked
#to see if they are below the SoftAvailLicPercent. 
$Script:ActiveUnitsCheck = 1000

####################
# Filter Variables #
####################

$VisioGroups = @()
$VisioGroups += "MSFT-VISIO-365-C2R-X86"

$ProjectGroups = @()
$ProjectGroups += "MSFT-PROJECT-365-C2R-X86"

#############################
# Services to add Variables #
#############################
$EPServicesToAdd = "OFFICESUBSCRIPTION","EXCHANGE_S_ENTERPRISE","SHAREPOINTWAC","SHAREPOINTENTERPRISE"
$EMSServicesToAdd = "RMS_S_PREMIUM","INTUNE_A","RMS_S_ENTERPRISE","AAD_PREMIUM","MFA_PREMIUM"
$KioskServicesToAdd = "OFFICESUBSCRIPTION"
$VisioServicesToAdd = "VISIO_CLIENT_SUBSCRIPTION"
$ProjectServicesToAdd = "PROJECT_CLIENT_SUBSCRIPTION"

###################
# Connect to MSOL #
###################

# Get Credentials
$username = "srvlic@tenant.onmicrosoft.com"
$password = Get-Content "D:\Encrypted.txt" | ConvertTo-SecureString
$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username, $password

# Import modules needed to run script
Import-Module MSOnline
Import-Module Active*

# Connect to MSonline with Above Creds
Connect-MsolService -Credential $cred

###########################################################
# Get MSOLAccount Sku's for licenses that will be enabled #
###########################################################
$EPAccountSku = Get-MsolAccountSku | ?{$_.AccountSkuID -ilike "*ENTERPRISEPACK"}
$EPServiceNamePlans = (Get-MsolAccountSku | ?{$_.AccountSkuID -eq $EPAccountSku.AccountSkuID} | %{$_.ServiceStatus}).ServicePlan | Select -ExpandProperty ServiceName
$EMSAccountSku = Get-MsolAccountSku | ?{$_.AccountSkuID -ilike "*EMS"}
$EMSServiceNamePlans = (Get-MsolAccountSku | ?{$_.AccountSkuID -eq $EMSAccountSku.AccountSkuID} | %{$_.ServiceStatus}).ServicePlan | Select -ExpandProperty ServiceName
$VisioAccountSku = Get-MsolAccountSku | ?{$_.AccountSkuID -ilike "*Intune*"}
$VisioServiceNamePlans = (Get-MsolAccountSku | ?{$_.AccountSkuID -eq $VisioAccountSku.AccountSkuID} | %{$_.ServiceStatus}).ServicePlan | Select -ExpandProperty ServiceName
$ProjectAccountSku = Get-MsolAccountSku | ?{$_.AccountSkuID -ilike "*Intune*"}
$ProjectServiceNamePlans = (Get-MsolAccountSku | ?{$_.AccountSkuID -eq $ProjectAccountSku.AccountSkuID} | %{$_.ServiceStatus}).ServicePlan | Select -ExpandProperty ServiceName

#####################################################
# Create array's of disabled plans per license pack #
#####################################################

##################################################################################
# Creating Default license plan.                                                 #
# To enabled services use variables in "Services to add Variables" section above #
##################################################################################

# Get plans to do comparison of what users have during assignment and to loop through and determine what plans to disable below
[System.Collections.ArrayList]$EPAvailablePlans = @()
[System.Collections.ArrayList]$EPDisabledPlans = @()
ForEach ($epplan in $EPServiceNamePlans) {
    $EPAvailablePlans.Add($epplan) | Out-Null
    $EPDisabledPlans.Add($epplan) | Out-Null
}

#Using available services list and remove all plans from the variable in the "Services to add Variables" section.
ForEach ($enabledsvc in $EPServicesToAdd) {
    $EPdisabledPlans.Remove($enabledsvc)
}

Write-Output "[PLANINFO]Available services for default users are: `r`n"
$EPAvailablePlans

Write-Output "[PLANINFO]Services to be disabled for default users (unless already enabled) are: `r`n"
$EPDisabledPlans

######################################
# EMS License Pack for default users #
######################################
# Get plans to do comparison of what users have during assignment and to loop through and determine what plans to disable below
[System.Collections.ArrayList]$EMSAvailablePlans = @()
[System.Collections.ArrayList]$EMSDisabledPlans = @()
ForEach ($emsplan in $EMSServiceNamePlans) {
    $EMSAvailablePlans.Add($emsplan) | Out-Null
    $EMSDisabledPlans.Add($emsplan) | Out-Null
}

#Using available services list and remove all plans from the variable in the "Services to add Variables" section.
ForEach ($enabledsvc in $EMSServicesToAdd) {
    $EMSDisabledPlans.Remove($enabledsvc)
}

Write-Output "[PLANINFO]Available EMS services for default users are: `r`n"
$EMSAvailablePlans

Write-Output "[PLANINFO]EMS Services to be disabled for default users (unless already enabled) are: `r`n"
If ([string]::IsNullOrEmpty($EMSDisabledPlans)) {
    Write-Output "N/A"
}
Else {
    $EMSDisabledPlans
}

##################################################################################
# Creating Kiosk license plan.                                                   #
# To enabled services use variables in "Services to add Variables" section above #
##################################################################################

# Get plans to do comparison of what users have during assignment and to loop through and determine what plans to disable below
[System.Collections.ArrayList]$KioskAvailablePlans = @()
[System.Collections.ArrayList]$KioskDisabledPlans = @()
ForEach ($kioskplan in $EPServiceNamePlans) {
    $KioskAvailablePlans.Add($kioskplan) | Out-Null
    $KioskDisabledPlans.Add($kioskplan) | Out-Null
}

#Using available services list and remove all plans from the variable in the "Services to add Variables" section.
ForEach ($kioskenabledsvc in $KioskServicesToAdd) {
    $KioskDisabledPlans.Remove($kioskenabledsvc)
}

Write-Output "[PLANINFO]Available services for kiosk users are: `r`n"
$KioskAvailablePlans

Write-Output "[PLANINFO]Services to be disabled for kiosk users (unless already enabled) are: `r`n"
$KioskDisabledPlans
##################################################################################
# Creating Visio license plan.                                                   #
# To enabled services use variables in "Services to add Variables" section above #
##################################################################################

# Get plans to do comparison of what users have during assignment and to loop through and determine what plans to disable below
[System.Collections.ArrayList]$VisioAvailablePlans = @()
[System.Collections.ArrayList]$VisioDisabledPlans = @()
ForEach ($visioplan in $VisioServiceNamePlans) {
    $VisioAvailablePlans.Add($visioplan) | Out-Null
    $VisioDisabledPlans.Add($visioplan) | Out-Null
}

#Using available services list and remove all plans from the variable in the "Services to add Variables" section.
ForEach ($visioenabledsvc in $VisioServicesToAdd) {
    $VisioDisabledPlans.Remove($visioenabledsvc)
}

Write-Output "[PLANINFO]Available services for Visio users are: `r`n"
$VisioAvailablePlans

Write-Output "[PLANINFO]Services to be disabled for Visio users (unless already enabled) are: `r`n"
$VisioDisabledPlans
##################################################################################
# Creating Project license plan.                                                 #
# To enabled services use variables in "Services to add Variables" section above #
##################################################################################

# Get plans to do comparison of what users have during assignment and to loop through and determine what plans to disable below
[System.Collections.ArrayList]$ProjectAvailablePlans = @()
[System.Collections.ArrayList]$ProjectDisabledPlans = @()
ForEach ($projectplan in $ProjectServiceNamePlans) {
    $ProjectAvailablePlans.Add($projectplan) | Out-Null
    $ProjectDisabledPlans.Add($projectplan) | Out-Null
}

#Using available services list and remove all plans from the variable in the "Services to add Variables" section.
ForEach ($projectenabledsvc in $ProjectServicesToAdd) {
    $ProjectDisabledPlans.Remove($projectenabledsvc)
}

Write-Output "[PLANINFO]Available services for Visio users are: `r`n"
$ProjectAvailablePlans

Write-Output "[PLANINFO]Services to be disabled for Visio users (unless already enabled) are: `r`n"
$ProjectDisabledPlans

##########################################
# Get only enabled plans and get a count #
##########################################
Function getEnabledPlans {
    Param ($RefObj,$DiffObj)
	
    ForEach ($planObj in $DiffObj) {
        $RefObj.Remove($planObj)
    }
    $Script:RefObjCount = $RefObj.Count
		
}

##########################
# Existing license check #
##########################

# Removes existing licenses from disabled plans array
Function removeInUseLicences {

    # Finds licenses the users already has and removes them from the disabled list
    $AddToPlan = (($LicOptUser.Licenses | ?{$_.AccountSkuID -eq $AccountSku.AccountSkuID}).ServiceStatus | ?{$_.ProvisioningStatus -eq "Success"}).ServicePlan.ServiceName
    Foreach ($status in $AddToPlan) {

        $Plan.Remove($status.ServiceName)

    }

}
# Check users with licenses already assigned and ensure no existing licenses are removed
Function createLicencseOption {
    Param ($LicOptUser, $LicPackType, $ExistingLic = $False)



    Switch ($LicPackType) {
        "DefaultUser" {
            # Enable Office Pro Plus and Exchange Online Licenses
            [System.Collections.ArrayList]$Plan = $EPdisabledPlans
            If ($ExistingLic -eq $True) {
                removeInUseLicences
            }
            Write-Output "[LICPLAN]Services to be disabled for $($LicOptUser.UserPrincipalName) are: `r`n "
            $Plan

            #creates new license to assign to user with existing plans left enabled
            $Script:LicenseOptions = New-MsolLicenseOptions -AccountSkuId $EPAccountSku.AccountSkuId -DisabledPlans $Plan

        }

        "EMSUser" {
            # Enable EMS Licenses
            [System.Collections.ArrayList]$Plan = $EMSDisabledPlans
            If ($ExistingLic -eq $True) {
                removeInUseLicences
            }
            Write-Output "[LICPLAN]Services to be disabled for $($LicOptUser.UserPrincipalName) are: `r`n "
            $Plan

            #creates new license to assign to user with existing plans left enabled
            $Script:LicenseOptions = New-MsolLicenseOptions -AccountSkuId $EMSAccountSku.AccountSkuId -DisabledPlans $Plan

        }

        "KioskUser" {
            # Enable Office Pro Plus Only
            [System.Collections.ArrayList]$Plan = $KioskdisabledPlans
            If ($ExistingLic -eq $True) {
                removeInUseLicences
            }
            Write-Output "[LICPLAN]Services to be disabled for $($LicOptUser.UserPrincipalName) are: `r`n "
            $Plan

            #creates new license to assign to user with existing plans left enabled
            $Script:LicenseOptions = New-MsolLicenseOptions -AccountSkuId $EPAccountSku.AccountSkuId -DisabledPlans $Plan

        }

        "VisioUser" {
            # Enable Visio Licenses
            [System.Collections.ArrayList]$Plan = $VisiodisabledPlans

            If ($ExistingLic -eq $True) {
                removeInUseLicences
            }
            Write-Output "[LICPLAN]Services to be disabled for $($LicOptUser.UserPrincipalName) are: `r`n "
            $Plan

            #creates new license to assign to user with existing plans left enabled
            $Script:LicenseOptions = New-MsolLicenseOptions -AccountSkuId $VisioAccountSku.AccountSkuId -DisabledPlans $Plan


        }

        "ProjectUser" {
            # Enable Project Licenses
            [System.Collections.ArrayList]$Plan = $ProjectdisabledPlans

            If ($ExistingLic -eq $True) {
                removeInUseLicences
            }
            Write-Output "[LICPLAN]Services to be disabled for $($LicOptUser.UserPrincipalName) are: `r`n "
            $Plan			

            #creates new license to assign to user with existing plans left enabled
            $Script:LicenseOptions = New-MsolLicenseOptions -AccountSkuId $ProjectAccountSku.AccountSkuId -DisabledPlans $Plan


        }
    }

}


##########################################
# Get user information for export to csv #
##########################################

Function createNewObj {
    Param ($MSOLUserObj,$ADUserObj,$LicTypeObj,$Note)

    If ($MSOLUserObj -ne $null) {
	
        $licuser = New-Object System.Object
        $licuser | Add-Member -Type NoteProperty DisplayName -Value $MSOLUserObj.DisplayName
        $licuser | Add-Member -Type NoteProperty UserPrincipalName -Value $MSOLUserObj.UserPrincipalName
        $licuser | Add-Member -Type NoteProperty EmployeeType -Value $ADUserObj.employeeType
        $licuser | Add-Member -Type NoteProperty extensionAttribute15 -Value $ADUserObj.extensionAttribute15
        $licuser | Add-Member -Type NoteProperty extensionAttribute9 -Value $ADUserObj.extensionAttribute9
        $licuser | Add-Member -Type NoteProperty ProxyAddresses -Value ($MSOLUserObj.ProxyAddresses -join ";")
        $licuser | Add-Member -Type NoteProperty IsLicensed -Value $MSOLUserObj.IsLicensed
        $licuser | Add-Member -Type NoteProperty Licenses -Value ($MSOLUserObj.Licenses.AccountSkuid -join ";")
        $licuser | Add-Member -Type NoteProperty OverallProvisioningStatus -Value $MSOLUserObj.OverallProvisioningStatus
        $licuser | Add-Member -Type NoteProperty ObjectId -Value $MSOLUserObj.ObjectId
        $licuser | Add-Member -Type NoteProperty LiveId -Value $MSOLUserObj.LiveId
        $licuser | Add-Member -Type NoteProperty ImmutableId -Value $MSOLUserObj.ImmutableId
        $licuser | Add-Member -Type NoteProperty LicenseType -Value $LicTypeObj
        $licuser | Add-Member -Type NoteProperty Note -Value $Note
        ForEach ($license in $($MSOLUserObj.Licenses)) {
        $licacctsku = $license.AccountSkuID
        ForEach ($status in $($license.ServiceStatus)) {
        $licnoteprop = $licacctsku + " - " + $status.ServicePlan.ServiceName
        $licuser | Add-Member -Type NoteProperty $licnoteprop -Value $status.ProvisioningStatus
    }
						
}
}
ElseIf ($MSOLUserObj -eq $null) {

    $licuser = New-Object System.Object
    $licuser | Add-Member -Type NoteProperty DisplayName -Value $ADUserObj.DisplayName
    $licuser | Add-Member -Type NoteProperty UserPrincipalName -Value $ADUserObj.UserPrincipalName
    $licuser | Add-Member -Type NoteProperty EmployeeType -Value $ADUserObj.employeeType
    $licuser | Add-Member -Type NoteProperty extensionAttribute15 -Value $ADUserObj.extensionAttribute15
    $licuser | Add-Member -Type NoteProperty extensionAttribute9 -Value $ADUserObj.extensionAttribute9
    $licuser | Add-Member -Type NoteProperty ProxyAddresses -Value ($ADUserObj.ProxyAddresses -join ";")
    $licuser | Add-Member -Type NoteProperty IsLicensed -Value $null
    $licuser | Add-Member -Type NoteProperty Licenses -Value $null
    $licuser | Add-Member -Type NoteProperty OverallProvisioningStatus -Value $null
    $licuser | Add-Member -Type NoteProperty ObjectId -Value $null
    $licuser | Add-Member -Type NoteProperty LiveId -Value $null
    $licuser | Add-Member -Type NoteProperty ImmutableId -Value $null
    $licuser | Add-Member -Type NoteProperty LicenseType -Value $LicTypeObj
    $licuser | Add-Member -Type NoteProperty Note -Value $Note
	
	
}

$Script:AllUsers += $licuser

}

############################
# Assign licenses to users #
############################

Function assignLicense {
    Param ($assignLicUPN, $assignLicMSOLUser, $assignLicADUser, $assignLicType, $assignAccountSku, [int]$assignLicSvcsCount) 

    
    # Determine if user needs a license or not, ensure no existing enabled services are not disabled, and add user to the report
    If ($assignLicSvcsCount -eq $RefObjCount) {
        Write-Output "[LICENSE] User $assignLicUPN is licensed for $assignLicType licenses, see report for details."
			
        createNewObj -MSOLUserObj $assignLicMSOLUser -ADUserObj $assignLicADUser -LicTypeObj $assignLicType -Note "Already Licensed ($assignLicType)"					
    }
    ElseIf ($assignLicSvcsCount -ge 0) {
        Write-Output "[LICENSE] User $assignLicUPN is licensed for $assignLicType licenses, but does not have all services enabled."
        Write-Output "[ADDSERVICE] Adding service to existing license..."
			
        createLicencseOption -LicOptUser $assignLicMSOLUser -LicPackType $assignLicType -ExistingLic $True
			
        # Set the users Region
        Set-MsolUser -UserPrincipalName $assignLicUPN -UsageLocation $UsageLocation

        # Grants a user a license and disables all above services, exluding office
        Set-MsolUserLicense -UserPrincipalName $assignLicUPN -LicenseOptions $LicenseOptions
			
        createNewObj -MSOLUserObj $assignLicMSOLUser -ADUserObj $assignLicADUser -LicTypeObj $assignLicType -Note "Added Service to License ($assignLicType)"
    }
    ElseIf ($assignLicSvcsCount -eq -99) {
        Write-Output "[NOLICENSE] User $assignLicUPN needs a license"
            
        If ((Get-MsolAccountSku | ?{$_.AccountSkuId -eq $assignAccountSku.AccountSkuId}).ActiveUnits - (Get-MsolAccountSku | ?{$_.AccountSkuId -eq $assignAccountSku.AccountSkuId}).ConsumedUnits -ge $HardAvailLicThreshold) {
            Write-Output "[ADDLICENSE] Adding $assignLicType license to user $assignLicUPN"
			
            createLicencseOption -LicOptUser $assignLicMSOLUser -LicPackType $assignLicType
			
            # Set the users Region
            Set-MsolUser -UserPrincipalName $assignLicUPN -UsageLocation $UsageLocation

            # Grants a user a license and disables all above services, exluding office
            Set-MsolUserLicense -UserPrincipalName $assignLicUPN -AddLicenses $assignAccountSku.AccountSkuId -LicenseOptions $LicenseOptions
            $assignLicMSOLUserTemp = Get-MsolUser -UserPrincipalName $assignLicUPN -ErrorAction SilentlyContinue
            createNewObj -MSOLUserObj $assignLicMSOLUserTemp -ADUserObj $assignLicADUser -LicTypeObj $assignLicType -Note "Add license ($assignLicType)"
        }
        Else {
            Write-Output "[NOLIC]Not enough licenses available to assign more."
            Write-EventLog -Source O365LicenseScript -LogName Application -EventId 22365 -EntryType Error -Message "Unable to assign a license to user $assignLicUPN. License threshold of $HardAvailLicThreshold met."
            createNewObj -MSOLUserObj $assignLicMSOLUserTemp -ADUserObj $assignLicADUse -LicTypeObj $assignLicTyper -Note "FAILED: No more licenses ($assignLicType)"
        }

    }		
    Else {
        Write-Output "[?ERROR?] Something is hosed. Fix it"
        createNewObj -MSOLUserObj $assignLicMSOLUser -ADUserObj $assignLicADUser -LicTypeObj $assignLicType -Note "Error Needs Fixing ($assignLicType)"
    }
		
		
		
}

########################################
# Determine what users need a license  #
########################################

Function evalLicenseReq {
    Param ($LicType, $UserList)


    # Set License Sku for each user type
    If ($LicType -eq "DefaultUser") {
        $AccountSku = $EPAccountSku
        $AvailablePlans = $EPAvailablePlans
    }
    ElseIf ($LicType -eq "EMSUser") {
        $AccountSku = $EMSAccountSku
        $AvailablePlans = $EMSAvailablePlans
    }
    ElseIf ($LicType -eq "KioskUser") {
        $AccountSku = $EPAccountSku
        $AvailablePlans = $KioskAvailablePlans
    }
    ElseIf ($LicType -eq "VisioUser") {
        $AccountSku = $VisioAccountSku
        $AvailablePlans = $VisioAvailablePlans
    }
    ElseIf ($LicType -eq "ProjectUser") {
        $AccountSku = $ProjectAccountSku
        $AvailablePlans = $ProjectAvailablePlans
    }
    Else {
        Write-Output "[?ERROR?] Something is hosed. Fix it"
    }
	
 
    ForEach ($user in $UserList) {
        $UPN = $user.UserPrincipalName
        $o365lic = Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue	
	
        If ($o365lic -ne $null) {
            If ($($o365lic.isLicensed) -eq $True) {
					
            $EnabledServices = (($o365lic.Licenses | ?{$_.AccountSkuID -eq $AccountSku.AccountSkuID}).ServiceStatus | ?{$_.ProvisioningStatus -eq "Success"}).ServicePlan.ServiceName
            If ($EnabledServices -ne $null) {
                [array]$EnabledServicesCompare = Compare-Object -ReferenceObject $AvailablePlans -DifferenceObject $EnabledServices -SyncWindow 15 -IncludeEqual | ?{$_.SideIndicator -eq "=="}
                [int]$EnabledServicesCount = $EnabledServicesCompare.Count
                assignLicense -assignLicUPN $UPN -assignLicMSOLUser $o365lic -assignLicADUser $user -assignLicType $LicType -assignAccountSku $AccountSku -assignLicSvcsCount $EnabledServicesCount
				
            }
            Else {
                If (($o365lic.Licenses | ?{$_.AccountSkuID -eq $AccountSku.AccountSkuID}) -eq $null) {
				        
                    assignLicense -assignLicUPN $UPN -assignLicMSOLUser $o365lic -assignLicADUser $user -assignLicType $LicType -assignAccountSku $AccountSku -assignLicSvcsCount -99
					
                }
					
                Else {
					    
                    assignLicense -assignLicUPN $UPN -assignLicMSOLUser $o365lic -assignLicADUser $user -assignLicType $LicType -assignAccountSku $AccountSku -assignLicSvcsCount 99
					
                }
				
            }
        }
        Else {
            Write-Output "[NOLICENSE] IsLicensed is false. assigning -99 to user $UPN"
            assignLicense -assignLicUPN $UPN -assignLicMSOLUser $o365lic -assignLicADUser $user -assignLicType $LicType -assignAccountSku $AccountSku -assignLicSvcsCount -99
			
        }
		
    }
		
    Else {
        Write-Output "[NOTFOUND] User $UPN not found in cloud"
			
        createNewObj -MSOLUserObj $o365lic -ADUserObj $user -Note "No Cloud Account Found"
			
    }	
	 
}	

}

##################################
# Remove disabled users licenses #
##################################

Function removeLicense {
    Write-Output "Getting all licensed users, this may take a few minutes..."
    $LicensedUsers = Get-MsolUser -All | ?{$_.isLicensed -eq $true}
    $disabledUsers = @()
    ForEach ($licensedUser in $LicensedUsers) {

        $UACcheck = Get-ADUser -Server $ScriptDC -Filter {UserPrincipalName -eq $licensedUser.UserPrincipalName -and enabled -eq $False}
        $memberOfcheck =  Get-ADUser -Server $ScriptDC -Properties memberOf -Filter {UserPrincipalName -eq $licensedUser.UserPrincipalName}
        $visioGroupCheck = $null
        $projectGroupCheck = $null
        If ($UACcheck -ne $null) {
            Write-Output "[DISABLEDUSER]$($licensedUser.UserPrincipalName) is disabled."
            $disabledUsers += $UACcheck
        }
        Else {
            Write-Output "[ENABLEDUSER]$($licensedUser.UserPrincipalName) is enabled."
            If ($licensedUser.Licenses.AccountSkuID -contains $VisioAccountSku.AccountSkuId) {
                ForEach ($visioLicGroup in $VisioGroups) {
                    If ($memberOfcheck.memberOf -contains (Get-ADGroup -Identity $visioLicGroup).DistinguishedName) {     
                        Write-Output "[MEMBER]$($licensedUser.UserPrincipalName) user is a member of $visioLicGroup"
                        $visioGroupCheck = "Member"
                    }
                    Else {
                        Write-Output "[NOTMEMBER]$($licensedUser.UserPrincipalName) user is not a member of $visioLicGroup"
                        $visioGroupCheck = "NotMember"
                    }
                }

                If ($visioGroupCheck -eq "NotMember") {
                    Write-Output "[REMOVELIC]$($licensedUser.UserPrincipalName) was not a member of any Visio groups. Visio license is being removed..."
                    #Set-MsolUserLicense -UserPrincipalName $licensedUser.UserPrincipalName -RemoveLicenses $VisioAccountSku.AccountSkuId
                }
                ElseIf ($visioGroupCheck -eq "Member") {
                    Write-Output "[NOCHANGE]$($licensedUser.UserPrincipalName) is a member of a Visio groups. Visio license will be kept."
                }
                Else {
                    Write-Output "[?ERROR?]Unabled to determine $($licensedUser.UserPrincipalName) Visio group status."
                }

            }

            Else {
                Write-Output "[NOVISGROUPEVAL]$($licensedUser.UserPrincipalName) does not have Visio licenses. No evaluation required."
            }

            If ($licensedUser.Licenses.AccountSkuID -contains $ProjectAccountSku.AccountSkuId) {

                ForEach ($projectLicGroup in $ProjectGroups) {
                    If ($memberOfcheck.memberOf -contains (Get-ADGroup -Identity $projectLicGroup).DistinguishedName) {
                        Write-Output "[MEMBER]$($licensedUser.UserPrincipalName) user is a member of $projectLicGroup"
                        $projectGroupCheck = "Member"
                    }
                    Else {
                        Write-Output "[NOTMEMBER]$($licensedUser.UserPrincipalName) user is not a member of $projectLicGroup"
                        $projectGroupCheck = "NotMember"
                    }
                }

                If ($projectGroupCheck -eq "NotMember") {
                    Write-Output "[REMOVELIC]$($licensedUser.UserPrincipalName) was not a member of any Project groups. Project license is being removed..."
                    #Set-MsolUserLicense -UserPrincipalName $licensedUser.UserPrincipalName -RemoveLicenses $ProjectAccountSku.AccountSkuId
                }
                ElseIf ($projectGroupCheck -eq "Member") {
                    Write-Output "[NOCHANGE]$($licensedUser.UserPrincipalName) is a member of a Project groups. Project license will be kept."
                }
                Else {
                    Write-Output "[?ERROR?]Unabled to determine $($licensedUser.UserPrincipalName) Project group status."
                }   

            }
           
            Else {
                Write-Output "[NOPROJGROUPEVAL]$($licensedUser.UserPrincipalName) does not have Project licenses. No evaluation required."
            }
       
        }
         
        
    }

    ForEach ($disuser in $disabledUsers) {
        Write-Output "[REMOVELIC]Removing licenses from user $($disuser.UserPrincipalName)"
        Set-MsolUserLicense -UserPrincipalName $disuser.UserPrincipalName -RemoveLicenses (Get-MsolUser -UserPrincipalName $disuser.UserPrincipalName).Licenses.AccountSkuID
        
    }


}

#######################
# License count check #
#######################

Function licenseCountCheck {
    $AccountSkus = Get-MsolAccountSku
    $LicenseLimitBody = @()

    ForEach ($sku in $AccountSkus) {

        If ($sku.ActiveUnits -eq 0) {
            Write-Output "[AVAILLICCOUNT]Not required to do license check for $($sku.AccountSkuId), Active Units equals 0"
        }
        Else {
        
            If ($sku.ActiveUnits -ge $ActiveUnitsCheck) {
                If ($sku.ActiveUnits - $sku.ConsumedUnits -ge $SoftAvailLicThreshold) {
                    Write-Output "[AVAILLICCOUNT]Available license count for $($sku.AccountSkuId) is ABOVE the set threshold of $SoftAvailLicThreshold"
                }
                Else {
                    Write-Output "[AVAILLICCOUNT]Available license count for $($sku.AccountSkuId) is BELOW the set threshold of $SoftAvailLicThreshold"
                    $LicenseLimitMSG = New-Object System.Object
                    $LicenseLimitMSG | Add-Member -Type NoteProperty AccountSKUID -Value $sku.AccountSkuId
                    $LicenseLimitMSG | Add-Member -Type NoteProperty ActiveUnits -Value $sku.ActiveUnits
                    $LicenseLimitMSG | Add-Member -Type NoteProperty ConsumedUnits -Value $sku.ConsumedUnits
                    $LicenseLimitMSG | Add-Member -Type NoteProperty AvailableUnits -Value $($sku.ActiveUnits - $sku.ConsumedUnits)
                $LicenseLimitMSG | Add-Member -Type NoteProperty ThresholdValue -Value $SoftAvailLicThreshold
                $LicenseLimitBody += $LicenseLimitMSG
                Write-EventLog -Source O365LicenseScript -LogName Application -EventId 11365 -EntryType Information -Message ($LicenseLimitBody | Out-String)
            
            }
        }
        Else {
            If ($sku.ActiveUnits - $sku.ConsumedUnits -ge $sku.ActiveUnits*$SoftAvailLicPercent) {
                Write-Output "[AVAILLICCOUNT]Available license count for $($sku.AccountSkuId) is ABOVE the set threshold of $("{0:P0}" -f $SoftAvailLicPercent)"
            }
            Else {
                Write-Output "[AVAILLICCOUNT]Available license count for $($sku.AccountSkuId) is BELOW the set threshold of $("{0:P0}" -f $SoftAvailLicPercent)"
                $LicenseLimitMSG = New-Object System.Object
                $LicenseLimitMSG | Add-Member -Type NoteProperty AccountSKUID -Value $sku.AccountSkuId
                $LicenseLimitMSG | Add-Member -Type NoteProperty ActiveUnits -Value $sku.ActiveUnits
                $LicenseLimitMSG | Add-Member -Type NoteProperty ConsumedUnits -Value $sku.ConsumedUnits
                $LicenseLimitMSG | Add-Member -Type NoteProperty AvailableUnits -Value $($sku.ActiveUnits - $sku.ConsumedUnits)
            $LicenseLimitMSG | Add-Member -Type NoteProperty ThresholdValue -Value $("{0:P0}" -f $SoftAvailLicPercent)
        $LicenseLimitBody += $LicenseLimitMSG
        Write-EventLog -Source O365LicenseScript -LogName Application -EventId 11365 -EntryType Information -Message ($LicenseLimitBody | Out-String)
            
    }

}
}   
    
}
    
If ($LicenseLimitBody -ne $null) {
    Write-Output "[LICALERT]License limits reached, sending email alert."
    Send-MailMessage -SmtpServer $SMTPServer -To $SMTPTOAddresses -From $SMTPFROMAddress -Subject $SMTPSubect -Body ($LicenseLimitBody | Out-String)
}
Else {
    Write-Output "[NOLICALERT]No license limits reached."
}
}

##################################################################
# Remove reports created by this script older than set age limit #
##################################################################

Function removeReports {
    # Clean up old reports
    #Get report age limit
    $Limit = (Get-Date).AddDays($ReportAgeLimit)
    Get-ChildItem $ReportPath -Include *.csv -Recurse | ? { $_.CreationTime -lt $Limit } | Remove-Item -Force

}

######################################
# Get list of users per license type #
######################################

# Default users license. Change filter as necessary to filter by group membership or different attributes. These are attributes found on premise, not in the cloud. 
Function getDefaultUsers {

    # Get Default users to assign licenses
    $DefaultUsers = Get-ADUser -Server $ScriptDC -Properties extensionAttribute9,extensionAttribute15,msExchRecipientTypeDetails,employeeType -ResultSetSize $null -Filter {UserprincipalName -like "testuser4*" -and enabled -eq $True} | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and ($_.msExchRecipientTypeDetails -eq 1 -or $_.msExchRecipientTypeDetails -eq 2147483648)}
	
    # Calls function to find all enabled plans
    getEnabledPlans -RefObj $EPAvailablePlans -DiffObj $EPdisabledPlans 
	
    # Calls function to assign the license
    evalLicenseReq -LicType "DefaultUser" -UserList $DefaultUsers

    $AllUsers | Export-Csv "$ReportPath\Default_enableusers_$Date.csv" -NoTypeInformation
    $Script:AllUsers = @()
	

}

# Default users license. Change filter as necessary to filter by group membership or different attributes. These are attributes found on premise, not in the cloud. 
Function getEMSUsers {

    # Get EMS users to assign licenses
    $EMSUsers = Get-ADUser -Server $ScriptDC -Properties extensionAttribute9,extensionAttribute15,msExchRecipientTypeDetails,employeeType -ResultSetSize $null -Filter {UserprincipalName -like "testuser4*" -and enabled -eq $True} | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and ($_.msExchRecipientTypeDetails -eq 1 -or $_.msExchRecipientTypeDetails -eq 2147483648)}
	
    # Calls function to find all enabled plans
    getEnabledPlans -RefObj $EMSAvailablePlans -DiffObj $EMSDisabledPlans
	
    # Calls function to assign the license
    evalLicenseReq -LicType "EMSUser" -UserList $EMSUsers

    $AllUsers | Export-Csv "$ReportPath\EMS_enableusers_$Date.csv" -NoTypeInformation
    $Script:AllUsers = @()
	

}

Function getKioskUsers {

    # Get Default users to assign licenses
    $KioskUsers = Get-ADUser -Server $ScriptDC -Properties extensionAttribute9,extensionAttribute15,msExchRecipientTypeDetails,employeeType -ResultSetSize $null -Filter {UserprincipalName -like "w2umdg*" -and enabled -eq $True} | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and $_.msExchRecipientTypeDetails -eq $null}
	
    # Calls function to find all enabled plans
    getEnabledPlans -RefObj $KioskAvailablePlans -DiffObj $KioskdisabledPlans 
	
    # Calls function to assign the license
    evalLicenseReq -LicType "KioskUser" -UserList $KioskUsers

    $AllUsers | Export-Csv "$ReportPath\Kiosk_enableusers_$Date.csv" -NoTypeInformation
    $Script:AllUsers = @()
	

}

# Visio users license. Change filter as necessary to filter by group membership or different attributes. These attributes are found on premise, not in the cloud.
Function getVisioUsers {

    $VisioUsers = @()
    # Get Groups and members
    ForEach ($group in $VisioGroups) {
        Write-Output "Group $group"
        $VisioUsers += Get-ADGroupMember -Identity $group | Get-ADUser -Server $ScriptDC -Properties extensionAttribute15,employeeType,userAccountControl | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and $_.userAccountControl -ne 514 -and $_.userAccountControl -ne 546 -and $_.userAccountControl -ne 66050}
    }


    # Calls function to find all enabled plans
    getEnabledPlans -RefObj $VisioAvailablePlans -DiffObj $VisiodisabledPlans 
	
    # Calls function to assign the license
    evalLicenseReq -LicType "VisioUser" -UserList $VisioUsers

    $AllUsers | Export-Csv "$ReportPath\Visio_enableusers_$Date.csv" -NoTypeInformation
    $Script:AllUsers = @()

}

Function getProjectUsers {

    $ProjectUsers = @()
    # Get Groups and members
    ForEach ($group in $ProjectGroups) {
        Write-Output "Group $group"
        $ProjectUsers += Get-ADGroupMember -Identity $group | Get-ADUser -Server $ScriptDC -Properties extensionAttribute15,employeeType,userAccountControl | ?{$_.employeeType -eq "Contractor" -or $_.employeeType -eq "Employee" -and $_.userAccountControl -ne 514 -and $_.userAccountControl -ne 546 -and $_.userAccountControl -ne 66050}
    }

    # Calls function to find all enabled plans
    getEnabledPlans -RefObj $ProjectAvailablePlans -DiffObj $ProjectdisabledPlans 
	
    # Calls function to assign the license
    evalLicenseReq -LicType "ProjectUser" -UserList $ProjectUsers

    $AllUsers | Export-Csv "$ReportPath\Project_enableusers_$Date.csv" -NoTypeInformation
    $Script:AllUsers = @()


}

Function checkDC {
    # Checks connectivity with DC's listed in the DC variable
    $DC1 = Test-ComputerSecureChannel -Server $ScriptDC1 -ErrorAction SilentlyContinue
    $ScriptDC = $null

    If ($DC1 -eq $True) {
        $Script:ScriptDC = $ScriptDC1
    }
    Else {
        Write-Output "[DCDOWN]$ScriptDC1 is not responding. Trying $ScriptDC2..."
        $DC2 = Test-ComputerSecureChannel -Server $ScriptDC2 -ErrorAction SilentlyContinue
        
        If ($DC2 -eq $True) {
            $Script:ScriptDC = $ScriptDC2
        }
        
        Else {
            Write-Output "[DCDOWN]$ScriptDC2 is not responding. Armageddon is upon us, no DC's are available."
            Write-Output "[EXIT]Exiting script"
            Stop-Transcript
            Exit
        }
    }
    Write-Output "[DCUSED]$ScriptDC will be the domain controller used for all commands"


}

Start-Transcript -Path $TranscriptPath -NoClobber

checkDC

Write-Output "Removing disabled users licenses."

removeLicense

Write-Output "Getting Contractors and Employee users to assign them the default & EMS license."

getDefaultUsers

getEMSUsers

Write-Output "Getting Kiosk users to assign them the Office Pro Plus license."

getKioskUsers

Write-Output "Getting Visio users to assign them a license."

getVisioUsers

Write-Output "Getting Project users to assign them a license."

getProjectUsers

Write-Output "Checking post license assignment count. Sending alerts if low"

licenseCountCheck

Write-Output "Removing old reports."

removeReports

Stop-Transcript