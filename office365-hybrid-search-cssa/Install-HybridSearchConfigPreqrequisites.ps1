# This script installs the two required prerequisites: 
# AdministrationConfig-EN.msi and msoidcli_64.msi.     
# It is assumed that these are available in the same folder as the script itself.      
# See the following links for downloading manually:     
# – http://www.microsoft.com/en-us/download/details.aspx?id=39267     
# – http://go.microsoft.com/fwlink/p/?linkid=236297     
#     
function Install-MSI {      
    param(     
        [Parameter(Mandatory=$true)]      
        [ValidateNotNullOrEmpty()]      
        [String] $path      
    )     
    $parameters = "/qn /i " + $path      
    $installStatement = [System.Diagnostics.Process]::Start( "msiexec", $parameters )       
    $installStatement.WaitForExit()      
}     
$scriptFolder = Split-Path $script:MyInvocation.MyCommand.Path      
$MSOIdCRLRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL" -ErrorAction SilentlyContinue      
if ($MSOIdCRLRegKey -eq $null) {     
    Write-Host "Installing Office Single Sign On Assistant" -Foreground Yellow      
    Install-MSI ($scriptFolder + "\msoidcli_64.msi")      
    Write-Host "Successfully installed!" -Foreground Green      
}     
else {     
    Write-Host "Office Single Sign On Assistant is already installed." -Foreground Green      
}     
$MSOLPSRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOnlinePowershell" -ErrorAction SilentlyContinue      
if ($MSOLPSRegKey -eq $null) {     
    Write-Host "Installing AAD PowerShell" -Foreground Yellow      
    Install-MSI ($scriptFolder + "\AdministrationConfig-EN.msi")      
    Write-Host "Successfully installed!" -Foreground Green      
}     
else {
    Write-Host "AAD PowerShell is already installed." -Foreground Green      
}     
