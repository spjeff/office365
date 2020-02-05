# from https://blog.mastykarz.nl/building-applications-office-365-apis-any-platform/

# Config
$clientID = "34f34d49-86b7-4437-a332-6fecaf95a244"
$tenantName = "spjeff.onmicrosoft.com"
$ClientSecret = "secret-goes-here"
$Username = "spjeff@spjeff.com"
$Password = "password-goes-here"

# Access Token
$ReqTokenBody = @{
    Grant_Type    = "Password"
    client_Id     = $clientID
    Client_Secret = $clientSecret
    Username      = $Username
    Password      = $Password
    Scope         = "https://graph.microsoft.com/.default"
} 
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
$TokenResponse


# Data call - READ
$apis = @(
'https://graph.microsoft.com/v1.0/me/contacts',
'https://graph.microsoft.com/v1.0/me',
'https://graph.microsoft.com/v1.0/users',
'https://graph.microsoft.com/v1.0/users/george@spjeff.com/contacts')
$apis |% {
    Write-Host $_ -Fore Yellow
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $_ -Method GET -Body $body -ContentType "text/plain"
}

# Data call - WRITE
$newcontact = '{"givenName": "Test","surname": "Contact","emailAddresses": [{"address": "test@contact.com","name": "Pavel Bansky"}],"businessPhones": ["+1 732 555 0102"]}'
$api = 'https://graph.microsoft.com/v1.0/users/george@spjeff.com/contacts'
Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $api -Method "POST" -Body $newcontact -ContentType "application/json"