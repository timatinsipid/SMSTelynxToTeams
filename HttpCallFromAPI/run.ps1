using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."
$convertedBody = [HashTable]::New($Request.Body, [StringComparer]::OrdinalIgnoreCase)
write-output $convertedBody.data.payload

# Interact with query parameters or the body of the request.
$clientID = $env:Teams_application_id
$ClientSecret =  $env:Teams_application_secret
$Username = $env:Teams_User_id
$Password =  $env:Teams_User_Password
$tenant = $env:Teams_Tenant
$teamid = $env:Teams_TeamID
$channelid = $env:Teams_ChannelID
$chatid = $env:Teams_ChatID
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
# Send a message to a channel
#$apiUrl = "https://graph.microsoft.com/v1.0/teams/$teamid/channels/$channelid/messages"
# Send a message to a chat
$apiUrl = "https://graph.microsoft.com/v1.0/chats/$chatid/messages"
$JSON = convertfrom-json -inputobject $($convertedBody.data.payload.from)
Write-Output $JSON
$title = $convertedBody.data.payload.from.phone_number
$subtitle = $convertedBody.data.payload.received_at
$text = $convertedBody.data.payload.text
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
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
