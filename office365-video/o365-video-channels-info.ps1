<# 
https://docs.microsoft.com/en-us/stream/migration-o365video-prep

Powershell tool to output information about Office 365 Video, to be used to help prep for O365 Video to Stream migration.

Change log:  
1/13/2020 by Marc Mroz
- Fix text file/csv encoding to UTF8 to support non-ASCII characters

1/9/2020 by Marc Mroz
- Fixed bug where owners/editor/viewer data for a channel wasn't output when the site collection's language wasn't in English
- Add 3 extra columns to the report to export channel's owners/editors/viewers permissions that don't have email addresses

11/20/2019 by Marc Mroz
-Added support for multifactor authentication (MFA)
-Added ability to loop over each video in the channel and sum up the view counts over time
#>

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

#Do you want to download the videos from O365 Video to your PC
$DownloadFile = $false

#--------------Don't make any changes to the script below------------

<# 
Creates 2 strings, one with a list of email addresses, the other with a list of names/titles
If a user in the permissions list doesn't have an email address (it's a non-mail SG, special SP ACL, etc) then we output the title
of that ACL into a seperate string.

Parameters:
  $permissionListObj (input) - The object listing users that have permission to a chanel as gotten from the O365 Video REST API:
                               /portals/hub/_api/VideoService/Channels('<Channel GUID>')/GetPermissionGroup(<enum>)/Users
  $permissionEmailList (output) - Formatted string with ; list of email addresses gotten from $permissionListObj
  $permisisonOtherList (output) - Formatted string with ; list of titles gotten from $permissionListObj
#>
function CreatePermissionsListStrings($permissionListObj, [ref]$permissionEmailList, [ref]$permisisonOtherList)
{ 
    $permissionEmailList.Value = ""
    $permisisonOtherList.Value = ""

    for ($i=0;$i -lt $permissionListObj.value.Count; $i++)
    {  
        $email = $permissionListObj.value[$i].Email
        if ($email.Contains("@")) {
            $permissionEmailList.Value = $permissionEmailList.Value + $email + '; '
        }
        else {
            $title = $permissionListObj.value[$i].Title
            $permisisonOtherList.Value = $permisisonOtherList.Value + $title + '; '
        }
    }
   
}


#Use the same folder the script is running in to output the reports
$PathforFiles = $PSScriptRoot + "\output\"

#Create "output" directory if it doesn't exist
New-Item -ItemType Directory -Force -Path $PathforFiles

Clear-Host
Write-Host "Reports on Office 365 Video channels, to be used for O365 Video to Stream migration prep."
Write-Host ""
Write-Host "CSV and text files will be output to this folder: "  
Write-Host $PathforFiles
Write-Host ""
Write-Host ""
Write-Host "The script will prompt you for a login. You must login with a O365 Global Admin user."
Write-Host ""
Write-Host "Enter the following info to run the script..."

#Get the user's SPO url
$O365SPOUrl = Read-Host -Prompt "SharePoint Online url (eg https://contoso.sharepoint.com)"

#Ask if the user wants to output view counts in the report (view counts make the script take forever to run)
$YesOrNo = Read-Host "Report on sum of views for all videos in each channel? This makes the script take much longer to run. (y/n)"
while("y","n" -notcontains $YesOrNo )
{
 $YesOrNo = Read-Host "Report on sum of views for all videos in each channel? This makes the script take much longer to run. (y/n)"
}
$IncludeViewCounts = $false
if ($YesOrNo -eq 'y') {$IncludeViewCounts = $true}

#Prompt the user for login and PW. They need to login as an O365 Global Admin other wise some API calls won't return data.
#-UseWebLogin will show a normal login window which supports MFA logins
Write-Host ""
Write-Host "You will be prompted to login to your organization. Make sure you login as an Office 365 global admin..."
Connect-PnPOnline -Url $O365SPOUrl -UseWebLogin

$csvFile = $PathforFiles + "Channels-Info.csv"
$LogFile = $PathforFiles + "Log-Trace.txt"
$FilesToExport = $PathforFiles + "Videos-File-List.txt"
$FileDownloadPath = $PathforFiles + "Downloads\"


$O365VideoPortalHubUrl = $O365SPOUrl + "/portals/hub"
$O365VideoRESTUrl = $O365VideoPortalHubUrl + "/_api/VideoService"

#Prompt for admin login - MFA will be supported as it uses the web login
Connect-PnPOnline -Url $O365SPOUrl -UseWebLogin


Write-Host ""
Write-Host ""
Write-Host "Running script" -NoNewline

#Get channel list from the O365 Video API
$O365VideoChannelsRestAPI = $O365VideoRESTUrl + "/Channels"
Add-Content $LogFile "Connecting to SPO..."

$Channels = Invoke-PnPSPRestMethod -Url $O365VideoChannelsRestAPI


#Get info about each channel
if ($Channels -ne $null)
{
    #Create channel object to hold info about a single channel that will be added to CSV
    $ChannelObj = New-Object -TypeName psobject
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel name' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel URL' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel GUID' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel owners with email addresses' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel owners without email addresses' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel editors with email addresses' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel editors without email addresses' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel viewers with email addresses' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Channel viewers without email addresses' -Value 'Missing data'
    $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Count of videos in channel' -Value 'Missing data'
    if ($IncludeViewCounts) {
        $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Sum of views (last 3 months) for videos in channel' -Value 'Missing data'
        $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Sum of views (last 6 months) for videos in channel' -Value 'Missing data'
        $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Sum of views (last 12 months) for videos in channel' -Value 'Missing data'
        $ChannelObj | Add-Member -MemberType NoteProperty -Name 'Sum of views (last 24 months) for videos in channel' -Value 'Missing data'
    }



    $TotalChannelMsg = 'There are total of ' + $Channels.value.Count + ' Channels'
    Add-Content $LogFile 'Channels retrieved, looping through channels and videos now...' -Encoding UTF8
    Add-Content $FilesToExport 'Channels retrieved, looping through channels and videos now...' -Encoding UTF8
    Add-Content $FilesToExport $TotalChannelMsg -Encoding UTF8

    for ($i=0;$i -lt $Channels.value.Count; $i++)
    {
        Write-Host "." -NoNewline

        $ChannelGUID = $Channels.value[$i].Id
        $ChannelRestAPIUrl= $O365VideoRESTUrl + "/Channels('" + $ChannelGUID + "')"
        
        $ChannelURL = $O365SPOUrl +  $Channels.value[$i].ServerRelativeUrl
        $ChannelURLText = 'Channel Site Collection URL is:  ' + $ChannelURL
        Add-Content $LogFile $ChannelURLText -Encoding UTF8

        $CreatorsRestAPIUrl = $ChannelRestAPIUrl + "/GetPermissionGroup(" + "2)/Users"
        $ContributorsRestAPIUrl = $ChannelRestAPIUrl + "/GetPermissionGroup(" + "0)/Users"
        $ViewersRestAPIUrl = $ChannelRestAPIUrl + "/GetPermissionGroup(" + "1)/Users"

        try {$CreatorsList = Invoke-PnPSPRestMethod -Url $CreatorsRestAPIUrl -ErrorAction Stop}
        catch {
            Add-Content $LogFile "Error calling: $CreatorsRestAPIUrl" -Encoding UTF8
            Add-Content $LogFile "+Error msg: $($PSItem.ToString())" -Encoding UTF8
        }
         
        try {$ContributorsList = Invoke-PnPSPRestMethod -Url $ContributorsRestAPIUrl -ErrorAction Stop}
        catch {
            Add-Content $LogFile "Error calling: $ContributorsRestAPIUrl" -Encoding UTF8
            Add-Content $LogFile "+Error msg: $($PSItem.ToString())" -Encoding UTF8
        }
        
        try {$ViewersList = Invoke-PnPSPRestMethod -Url $ViewersRestAPIUrl -ErrorAction Stop}
        catch {
            Add-Content $LogFile "Error calling: $ViewersRestAPIUrl" -Encoding UTF8
            Add-Content $LogFile "+Error msg: $($PSItem.ToString())" -Encoding UTF8
        }        
        
        Add-Content $LogFile '==========================================' -Encoding UTF8
        Add-Content $LogFile 'Enumerating Channel Owners' -Encoding UTF8
        Add-Content $LogFile $CreatorsList.value.Email -Encoding UTF8
        Add-Content $LogFile $CreatorsList.value.Title -Encoding UTF8
        Add-Content $LogFile '==========================================' -Encoding UTF8
        Add-Content $LogFile 'Enumerating Channel Editors' -Encoding UTF8
        Add-Content $LogFile $ContributorsList.value.Email -Encoding UTF8
        Add-Content $LogFile $ContributorsList.value.Title -Encoding UTF8
        Add-Content $LogFile '==========================================' -Encoding UTF8
        Add-Content $LogFile 'Enumerating Channel Viewers' -Encoding UTF8
        Add-Content $LogFile $ViewersList.value.Email -Encoding UTF8
        Add-Content $LogFile $ViewersList.value.Title -Encoding UTF8
        Add-Content $LogFile '=========================================='

        Add-Content $FilesToExport $ChannelURLText -Encoding UTF8
        $ChannelGUIDString = 'Channel GUID is:  ' + $ChannelGUID
        Add-Content $LogFile $ChannelGUIDString -Encoding UTF8
        Add-Content $FilesToExport $ChannelGUIDString -Encoding UTF8
        
        #add info to channel object which will be output to the CSV
        $ChannelObj.'Channel name' = $Channels.value[$i].Title
        $ChannelObj.'Channel URL' = $ChannelURL
        $ChannelObj.'Channel GUID' = $ChannelGUID

        #email property for users only populated for 
        # 1. licensed users - if the user isn't licensed SPO doesn't populate the email property
        # 2. Security groups that aren't mail enabled - if not mail enabled obviously no email property
        # 3. Special SP ACLs like "Everyone except external users"
        # so we are splitting into 2 columns in the report one with email addresses and one where we just show the titles of the permission entites

        $permissionEmailList = ""
        $permisisonOtherList = ""

        CreatePermissionsListStrings $CreatorsList ([ref]$permissionEmailList) ([ref]$permisisonOtherList)
        $ChannelObj.'Channel owners with email addresses' = $permissionEmailList
        $ChannelObj.'Channel owners without email addresses' = $permisisonOtherList
   
        CreatePermissionsListStrings $ContributorsList ([ref]$permissionEmailList) ([ref]$permisisonOtherList)
        $ChannelObj.'Channel editors with email addresses' = $permissionEmailList
        $ChannelObj.'Channel editors without email addresses' = $permisisonOtherList

        CreatePermissionsListStrings $ViewersList ([ref]$permissionEmailList) ([ref]$permisisonOtherList)
        $ChannelObj.'Channel viewers with email addresses' = $permissionEmailList
        $ChannelObj.'Channel viewers without email addresses' = $permisisonOtherList

        #Get all the videos in each channel
        $VideoChannelRESTUrl = $ChannelRestAPIUrl + "/Videos"
    
        try {$VideosInChannel = Invoke-PnPSPRestMethod -Url $VideoChannelRESTUrl -ErrorAction Stop}
        catch {
            Add-Content $LogFile "Error calling: $VideoChannelRESTUrl" -Encoding UTF8
            Add-Content $LogFile "+Error msg: $($PSItem.ToString())" -Encoding UTF8
        }     


        #Clear the sums of views on all videos in channel. Using null because we want to know if we were able to get any data or the API itself to get the
        #analytics was not returning anything at all. We don't want to confuse 0 views with we weren't able to get any view counts at all for this video because
        #it's new (not in search index) or the search analytics counts aren't tabulated or is broken. Will check at bottom if each sum is not null or not.
        $SumVideoViews3Months = $null
        $SumVideoViews6Months = $null
        $SumVideoViews12Months = $null
        $SumVideoViews24Months = $null

    
        if ($VideosInChannel -ne $null)
        {
            $videocountMsg = 'Channel has ' + $VideosInChannel.value.Count + ' videos'
            Add-Content $LogFile 'Channel is not empty, looping through videos now' -Encoding UTF8
            Add-Content $LogFile $videocountMsg -Encoding UTF8
            Add-Content $FilesToExport $videocountMsg -Encoding UTF8

            #add count of videos in the channel to object which will be output to CSV
            $ChannelObj.'Count of videos in channel' = $VideosInChannel.value.Count

            #Get info about each video in the channel
            for ($j=0;$j -lt $VideosInChannel.value.Count; $j++)
            {
                
                if ($IncludeViewCounts) {
                    #Get view counts for the last 24 months for a video
                    $VideoViewsOverTimeUrl = $VideoChannelRESTUrl + '(guid''' + $VideosInChannel.value[$j].ID + ''')/GetVideoDetailedViewCount'
                    Add-Content $LogFile $VideoViewsOverTimeUrl -Encoding UTF8

                    $VideoViews = $null
                    $CurrentVideoViewsLast3Months = $null
                    $CurrentVideoViewsLast6Months = $null
                    $CurrentVideoViewsLast12Months = $null
                    $CurrentVideoViewsLast24Months = $null
                    
                    try {$VideoViews = Invoke-PnPSPRestMethod -Url $VideoViewsOverTimeUrl -ErrorAction Stop}
                    catch {
                        Add-Content $LogFile "Error calling: $VideoViewsOverTimeUrl" -Encoding UTF8
                        Add-Content $LogFile "+Error msg: $($PSItem.ToString())" -Encoding UTF8
                    }   

                    #$MonthsCnt = $VideoViews.Months.Count
                    #$MonthsCntMsg = "Vidoe's month node count:"+ $MonthsCnt
                    #Add-Content $LogFile $MonthsCntMsg -Encoding UTF8

                    if ($VideoViews.Months.Count -ne 0) 
                    {

                        for ($k=0;$k -lt 24; $k++)
                        {
                            $MonthTotalHits = $VideoViews.Months[$k].TotalHits
                            if ($k -lt 3) {$CurrentVideoViewsLast3Months = $CurrentVideoViewsLast3Months + $MonthTotalHits}
                            if ($k -lt 6) {$CurrentVideoViewsLast6Months = $CurrentVideoViewsLast6Months + $MonthTotalHits}
                            if ($k -lt 12) {$CurrentVideoViewsLast12Months = $CurrentVideoViewsLast12Months + $MonthTotalHits}
                            if ($k -lt 24) {$CurrentVideoViewsLast24Months = $CurrentVideoViewsLast24Months + $MonthTotalHits}
                        }
                        
                        $SumVideoViews3Months = $SumVideoViews3Months + $CurrentVideoViewsLast3Months
                        $SumVideoViews6Months = $SumVideoViews6Months + $CurrentVideoViewsLast6Months
                        $SumVideoViews12Months = $SumVideoViews12Months + $CurrentVideoViewsLast12Months
                        $SumVideoViews24Months = $SumVideoViews24Months + $CurrentVideoViewsLast24Months

                    }
                }

                
                $VideoPath = $O365SPOUrl + $VideosInChannel.value[$j].ServerRelativeUrl
                #Add-Content $LogFile $VideoPath -Encoding UTF8

                $VideoFilePathFragments = $VideosInChannel.value[$j].ServerRelativeUrl.Split('/')
                $VideoFilePath = "/" + $VideoFilePathFragments[3] + "/" + $VideoFilePathFragments[4]
                $VideoFileDownloadPath = $FileDownloadPath + "\" + $VideoFilePathFragments[2]
                
                if ($DownloadFile)
                {
                    if ( -not(Test-Path $VideoFileDownloadPath))
                    {
                        New-Item -Path $FileDownloadPath -Name $VideoFilePathFragments[2] -ItemType "directory"
                    }
                    
                    $FileURL = $VideosInChannel.value[$j].ServerRelativeUrl
                    $FileURL = $FileURL.Replace("'", "''")
                    #$VideosInChannel.value[$j].ServerRelativeUrl
                    
                    Download-SPOFile -WebUrl $ChannelURL -UserName $UserName -Password $SecurePassword -FileUrl $FileURL -DownloadPath $VideoFileDownloadPath
                    
                }
                else
                {
                    Add-Content $FilesToExport $VideoPath -Encoding UTF8
                }
            }
            Add-Content $LogFile '**********************************************************' -Encoding UTF8
			Add-Content $LogFile '**********************************************************' -Encoding UTF8
        }
        
        #If we were able to get view counts all the videos in the channel then output those sums to the CSV
        if ($IncludeViewCounts) {
            $ChannelObj.'Sum of views (last 3 months) for videos in channel' = $SumVideoViews3Months
            $ChannelObj.'Sum of views (last 6 months) for videos in channel' = $SumVideoViews6Months
            $ChannelObj.'Sum of views (last 12 months) for videos in channel' = $SumVideoViews12Months
            $ChannelObj.'Sum of views (last 24 months) for videos in channel' = $SumVideoViews24Months
        }

        #Write out the CSV to the file now
        $ChannelObj | Export-Csv -Path $csvFile -Append -NoTypeInformation -Encoding UTF8
        
    }
    Write-Host "Done"
    Write-Host ""
    Write-Host "CSV and text files in this folder: " $PathforFiles
    Write-Host $csvFile
    Write-Host "  CSV spreadsheet of all channels in O365 Video with which users and mail enabled security groups have access"
    Write-Host ""
    Write-Host "$FilesToExport"
    Write-Host "  text file with links to all the SPO URLs for each channel and all videos within each channel"
    Write-Host ""
    Write-Host "$LogFile"
    Write-Host "  diagnostic log file for script if needed for debugging"

}





