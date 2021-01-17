# from https://gist.github.com/asadrefai/ecfb32db81acaa80282d
# from https://www.microsoft.com/en-us/download/confirmation.aspx?id=42038
# installed file [sharepointclientcomponents_16-6906-1200_x64-en-us] 
Try{
    Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
    Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
    Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll'
} 
catch {
    Write-Host $_.Exception.Message
    Write-Host "No further parts of the migration will be completed" 
}
# [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
# [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
# [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles")

# from https://community.idera.com/database-tools/powershell/ask_the_experts/f/learn_powershell_from_don_jones-24/2824/exporting-key-value-pair-using-export-csv-cmdlet
function ConvertTo-Object($hashtable) 
{
$hashtable = $ups
   $object = New-Object PSObject
   $hashtable.Keys | 
      ForEach-Object {
          Add-Member -inputObject $object -memberType NoteProperty -name $_ -value $hashtable[$_]
        }
   $object
}

# from https://sharepoint.stackexchange.com/questions/108664/powershell-script-for-user-profile-properties-in-sharepoint-online-2013
$SiteURL = "https://spjeff.sharepoint.com"
Connect-pnponline -Url $SiteURL
$Context = Get-PNPContext

#Identify users in the Site Collection
$Users = $Context.Web.SiteUsers
$Context.Load($Users)
$Context.ExecuteQuery()

#Create People Manager object to retrieve profile data
$PeopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($Context)
$coll = @()
$i=0
Foreach ($User in $Users)
{
    $i++
$UserProfile = $PeopleManager.GetPropertiesFor($User.LoginName)
$Context.Load($UserProfile)
$Context.ExecuteQuery()
   If ($UserProfile.Email -ne $null)
    {
    Write-Host "User:" $User.LoginName -ForegroundColor Green
    $ups = $UserProfile.UserProfileProperties
    $obj = ConvertTo-Object $ups
    $coll += $obj
    Write-Host ""
    }  
}

$coll | Export-Csv "Export-SPO-All-UPS.csv" -NoTypeInformation
Write-Host "DONE"