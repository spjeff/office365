# Backup-All-MS-Flow-JSON-and-ZIP.ps1
# from https://github.com/pnp/script-samples/blob/main/scripts/flow-export-all-flows-in-environment/README.md
# NOTE - Reference code above needs several changes to support new Micrsoft MS Flow V2 API standards.   Below code uses latest V2 API standards.
 
# Load PnP PowerShell Module
Import-Module "PnP.PowerShell"
 
# Connect to SharePoint Online
$url = "https://spjeffdev-admin.sharepoint.com"
Connect-PnPOnline -Url $url -Interactive
 
# Loop all Flows in all Environments and export as ZIP & JSON
$FlowEnvs = Get-PnPPowerPlatformEnvironment
foreach ($FlowEnv in $FlowEnvs) {
 
    # Display Name of Environment
    $environmentName = $FlowEnv.Name
    Write-Host "Getting All Flows in $environmentName Environment"

    #Remove -AsAdmin Parameter to only target Flows you have permission to access
    $flows = Get-AdminFlow -Environment $environmentName
 
    # Display Count of Flows
    Write-Host "Found $($flows.Count) Flows to export..."
 
    # Loop all Flows and export as ZIP & JSON
    foreach ($flow in $flows) {
        # Display Name of Flow
        Write-Host "Exporting as ZIP & JSON... $($flow.DisplayName)"
        $filename = $flow.DisplayName.Replace(" ", "")
 
        # Build Export Path
        $timestamp = Get-Date -Format "yyyyMMddhhmmss"
        $exportPath = "$($filename)_$($timestamp)"
        $exportPath = $exportPath.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
 
        # Execute Export to ZIP & JSON
        $flow | ft -a
        Export-PnPFlow -Environment $FlowEnv -Identity $flow.FlowName -PackageDisplayName $flow.DisplayName -AsZipPackage -OutPath "$exportPath.zip" -Force
        Export-PnPFlow -Environment $FlowEnv -Identity $flow.FlowName | Out-File "$exportPath.json"
    }
}
