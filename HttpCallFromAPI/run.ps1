using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."
Write-Output $Request
write-output $request.body
$convertedBody = [HashTable]::New($Request.Body, [StringComparer]::OrdinalIgnoreCase)
$data = $convertedBody | ConvertFrom-Json
# Interact with query parameters or the body of the request.
$name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body
}
$clientID = $env:Teams_application_id
$ClientSecret =  $env:Teams_application_secret
$Username = $env:Teams_User_id
$Password =  $env:Teams_User_Password
$tenant = $env:Teams_Tenant
$teamid = $env:Teams_TeamID
$channelid = $env:Teams_ChannelID
# Build Token Body
$ReqTokenBody = @{
    Grant_Type    = "Password"
    client_Id     = $clientID
    Client_Secret = $clientSecret
    Username      = $Username
    Password      = $Password
    Scope         = "https://graph.microsoft.com/.default"
} 

# Get Token
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
$apiUrl = "https://graph.microsoft.com/v1.0/teams/$teamid/channels/$channelid/messages"
$title = $Request.Body.data.from.phone_number
$subtitle = $Request.Body.data.subject
$text = $data
# Body to send the message
$body = @"
{
    "subject": null,
    "body": {
        "contentType": "html",
        "content": "<attachment id=\"74d20c7f34aa4a7fb74e2b30004247c5\"></attachment>"
    },
    "attachments": [
        {
            "id": "74d20c7f34aa4a7fb74e2b30004247c5",
            "contentType": "application/vnd.microsoft.card.thumbnail",
            "contentUrl": null,
            "content": "{\r\n  \"title\": \"$title\",\r\n  \"subtitle\": \"$subtitle\",\r\n  \"text\": \"$text\",\r\n  \r\n}",
            "name": null,
            "thumbnailUrl": null
        }
    ]
}
"@
# Send Teams Message
Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Body $Body -Method Post -ContentType 'application/json'

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

if ($name) {
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
