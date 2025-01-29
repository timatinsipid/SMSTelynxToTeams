using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.
$name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body
}
Write-Output $env:Teams_application_id
Write-Output $env:Teams_application_secret
Write-Output $env:Teams_User_id
Write-Output $env:Teams_User_Password
$clientID = $env:Teams_application_id
$ClientSecret =  $env:Teams_application_secret
$Username = $env:Teams_User_id
$Password =  $env:Teams_User_Password
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
write-output $TokenResponse

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

if ($name) {
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})


# param(
#     [parameter(Mandatory=$true)]
#     $chatid,
#     [parameter(Mandatory=$true)]
#     $title,
#     $subtitle,
#     [parameter(Mandatory=$true)]
#     $text
#  )

# # If passing the variable for chat ID, then get the chat ID, if not, will use the Chat ID passed in
# IF ($chatid -like "var_*") {
#     $chatid = Get-AutomationVariable -Name $chatid
# }
# #$chatid = Get-AutomationVariable -Name "var_TeamsChat_Collaboration"
# #var_TeamsChat_TestChat
# #var_TeamsChat_Collaboration

# Get the tenant
# $Tenant = Get-AutomationVariable -Name TeckTenantProd

# Get App Credentials to send Chat
# $appCredential = Get-AutomationPSCredential -Name 'AzureApp.collaboration-teams-automation'
# $clientID = $appCredential.UserName
# $ClientSecret =  $appCredential.GetNetworkCredential().Password

# # Get User Credentials to send Chat
# $userCredential = Get-AutomationPSCredential -Name 'svc.collab.bot@teck.com'
# $Username = $userCredential.UserName
# $Password =  $userCredential.GetNetworkCredential().Password

# # Build Token Body
# $ReqTokenBody = @{
#     Grant_Type    = "Password"
#     client_Id     = $clientID
#     Client_Secret = $clientSecret
#     Username      = $Username
#     Password      = $Password
#     Scope         = "https://graph.microsoft.com/.default"
# } 

# # Get Token
# $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

# # URI for Chat
# $apiUrl = "https://graph.microsoft.com/v1.0/chats/$ChatID@thread.v2/messages"

# # Body to send the message
# $body = @"
# {
#     "subject": null,
#     "body": {
#         "contentType": "html",
#         "content": "<attachment id=\"74d20c7f34aa4a7fb74e2b30004247c5\"></attachment>"
#     },
#     "attachments": [
#         {
#             "id": "74d20c7f34aa4a7fb74e2b30004247c5",
#             "contentType": "application/vnd.microsoft.card.thumbnail",
#             "contentUrl": null,
#             "content": "{\r\n  \"title\": \"$title\",\r\n  \"subtitle\": \"$subtitle\",\r\n  \"text\": \"$text\",\r\n  \r\n}",
#             "name": null,
#             "thumbnailUrl": null
#         }
#     ]
# }
# "@

# #"content": "{\r\n  \"title\": \"This is an example of posting a card\",\r\n  \"subtitle\": \"<h3>This is the subtitle</h3>\",\r\n  \"text\": \"Here is some body text. <br>\\r\\nAnd a <a href=\\\"http://microsoft.com/\\\">hyperlink</a>. <br>\\r\\nAnd below that is some buttons:\",\r\n  \"buttons\": [\r\n    {\r\n      \"type\": \"messageBack\",\r\n      \"title\": \"Login to FakeBot\",\r\n      \"text\": \"login\",\r\n      \"displayText\": \"login\",\r\n      \"value\": \"login\"\r\n    }\r\n  ]\r\n}",

# write-output $body
# # Send Teams Message
# Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Body $Body -Method Post -ContentType 'application/json'


