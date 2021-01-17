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
Connect-pnponline -Url $SiteURL -UseWebLogin
$Context = Get-PNPContext

#Identify users in the Site Collection
$Users = $Context.Web.SiteUsers
$Context.Load($Users)
$Context.ExecuteQuery()

#Create People Manager object to retrieve profile data
$PeopleManager = New-Object   Microsoft.SharePoint.Client.UserProfiles.PeopleManager($Context)
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

$coll | Export-Csv "D:\Export-SPO-All-UPS.csv" -NoTypeInformation
Write-Host "DONE"