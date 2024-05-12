<#
    Demo feasiblity and prototype code for download large data set rows from MS Dataverse table into local PowerShell console.
    Support pagingation and FetchXML query language for high performance and scale.

    * High scale processing High speed performance
    * Line number breakpoint precision debug
    * Line by line transcript LOG text run history
    * Import any third party module (SharePoint, PNP, SQL, etc.)

    References
    https://vishalgrade.com/2023/10/03/how-to-use-powershell-in-dynamics-crm-to-perform-crud-operations/

.EXAMPLE
	.\XRM-Cmdlet-Query-Dataverse-Rows.ps1
	
.NOTES  
	File Name:  XRM-Cmdlet-Query-Dataverse-Rows
	Author   :  Jeff Jones  - @spjeff
	Modified :  2024-05-11

.LINK
	https://admin.powerplatform.microsoft.com/environments
#>

# Modules
$ModuleName = "Microsoft.Xrm.Data.PowerShell"
Install-Module $ModuleName -Scope "CurrentUser"
Import-Module $ModuleName -Force
$xrmCommands = Get-Command -Module $ModuleName
$xrmCommands.Count

# Configuration
$url   = "https://org12345678.crm.dynamics.com/"
$fetch = @'
<fetch xmlns:generator='MarkMpn.SQL4CDS'>
  <entity name='aaduser'>
    <all-attributes />
    <filter>
      <condition attribute='city' operator='eq' value='Chicago' />
    </filter>
  </entity>
</fetch>
'@

# Connect
$conn = Get-CrmConnection -InteractiveMode
$conn 

# Download data
$result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch
$rows = $result.CrmRecords

# Display data 
$rows | Out-GridView
