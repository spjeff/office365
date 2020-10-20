<#
.SYNOPSIS
Connect to Exchange Online without the Click2Run
 
.DESCRIPTION
Connect to Exchange Online without the Click2Run
 
.PARAMETER UserPrincipalName
UserPrincipalName of the Admin Account
 
.EXAMPLE
Connect to Exchange Online
Connect-EXOPSSession -UserPrincipalName admin@contoso.com
 
.NOTES
Ref : https://www.michev.info/Blog/Post/1771/hacking-your-way-around-modern-authentication-and-the-powershell-modules-for-office-365
Only Support User Connection no Application Connect (As Of : 2019-05)
 
#>

Function Connect-EXOPSSession
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
    [cmdletbinding()]
    param (
    [parameter(Mandatory=$False)]
        $UserPrincipalName
    )
    if([string]::IsNullOrEmpty($UserPrincipalName))
    {
        $UserPrincipalName = Get-CurrentUPN
    }
    if([string]::IsNullOrEmpty($UserPrincipalName))
    {
        Throw "Can't determine User Principal Name, please use the parameter -UserPrincipalName to specify it."
    }
    else
    {
        $resourceUri = "https://outlook.office365.com"
        $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
        $clientid = "a0c73c16-a7e3-4564-9a95-2bdf47383716"

        if($Script:UPNEXOHeader){
            # Setting DateTime to Universal time to work in all timezones
            $DateTime = (Get-Date).ToUniversalTime()

            # If the authToken exists checking when it expires
            $TokenExpires = ($Script:UPNEXOHeader.ExpiresOn.datetime - $DateTime).Minutes
            $UPNMismatch = $UserPrincipalName -ne $Script:UPNEXOHeader.UserID
            $AppIDMismatch = $ClientID -ne $Script:UPNEXOHeader.AppID
            if($TokenExpires -le 0 -or $UPNMismatch -or $AppIDMismatch){
                Write-PSFMessage -Level Host -Message "Authentication need to be refresh" -ForegroundColor Yellow
                $Script:UPNEXOHeader = Get-OAuthHeaderUPN -clientId $ClientID -redirectUri $redirectUri -resourceAppIdURI $resourceURI -UserPrincipalName $UserPrincipalName
            }
        }
        # Authentication doesn't exist, calling Get-GraphAuthHeaderBasedOnUPN function
        else {
            $Script:UPNEXOHeader = Get-OAuthHeaderUPN -clientId $ClientID -redirectUri $redirectUri -resourceAppIdURI $resourceURI -UserPrincipalName $UserPrincipalName
        }
        $Result = $Script:UPNEXOHeader

        $Authorization =  $Result.Authorization
        $Password = ConvertTo-SecureString -AsPlainText $Authorization -Force
        $Ctoken = New-Object System.Management.Automation.PSCredential -ArgumentList $UserPrincipalName, $Password
        $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/PowerShell-LiveId?BasicAuthToOAuthConversion=true -Credential $Ctoken -Authentication Basic -AllowRedirection
        Import-Module (Import-PSSession $EXOSession -AllowClobber) -Global -DisableNameChecking
    }
}