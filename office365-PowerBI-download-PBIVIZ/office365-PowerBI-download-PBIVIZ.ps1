$path = Split-Path $MyInvocation.MyCommand.Path
$base = "https://visuals.azureedge.net/gallery-prod/"
$catalog = $base + "visualCatalog.json" 
$wr = Invoke-WebRequest -Uri $catalog
$json = $wr.Content | ConvertFrom-Json
$client = New-Object System.Net.WebClient

foreach ($j in $json) {
   $u = $base + $j.downloadUrl
   $u
   $file = $path + $j.downloadUrl
   $client.DownloadFile($u, $file)
}

$n = $json.Count
Write-Host "Downloaded $n Files" -Fore Green