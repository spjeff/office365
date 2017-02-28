<#
.SYNOPSIS
	Download all PBIVIZ visualization files from the MS Gallery.

.DESCRIPTION
	Download JSON feed with list of available visulizations and downloads each to the loca folder.

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Name		: office365-PowerBI-download-PBIVIZ.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.10
	Last Modified	: 02-28-2017
.LINK
	Source Code
		https://github.com/spjeff/office365/blob/master/office365-PowerBI-download-PBIVIZ/
#>

# Configure
$path = Split-Path $MyInvocation.MyCommand.Path
$base = "https://visuals.azureedge.net/gallery-prod/"
$catalog = $base + "visualCatalog.json" 

# Download JSON catalog
$wr = Invoke-WebRequest -Uri $catalog
$json = $wr.Content | ConvertFrom-Json

# Download each PBIVIZ binary file
$client = New-Object System.Net.WebClient
foreach ($j in $json) {
   $u = $base + $j.downloadUrl
   $u
   $file = $path + $j.downloadUrl
   $client.DownloadFile($u, $file)
}

# Summary
$n = $json.Count
Write-Host "Downloaded $n Files" -Fore Green