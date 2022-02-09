param (
    [Parameter(Mandatory=$True,HelpMessage="Application Id (obtained when registering the application in Azure AD")]
    [ValidateNotNullOrEmpty()]
    [string]$AppId = "",

	[Parameter(Mandatory=$False,HelpMessage="Redirect Url (specified when registering the application in Azure AD, use localhost).  Required for delegate permissions.")]
	[ValidateNotNullOrEmpty()]
	[string]$RedirectUrl = "http://localhost/code",

    [Parameter(Mandatory=$True,HelpMessage="Tenant domain")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId = ""

)

$createChatJSON = @"
{
  "chatType": "oneOnOne",
  "members": [
    {
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      "roles": ["owner"],
      "user@odata.bind": "https://graph.microsoft.com/v1.0/users('%user1%')"
    },
    {
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      "roles": ["owner"],
      "user@odata.bind": "https://graph.microsoft.com/v1.0/users('%user2%')"
    }
  ]
}
"@

# https://graph.microsoft.com/v1.0/chats

# Get token


# Acquire auth code (needed to request token)
$authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize?client_id=$AppId&response_type=code&redirect_uri=$RedirectUrl&response_mode=query&scope=openid%20profile%20email%20offline_access%20"
Start-Process $authUrl

Write-Host "Please complete log-in via the web browser, and then paste the redirect URL (including auth code) here to continue" -ForegroundColor Green
$authcode = Read-Host "Auth code"
$codeStart = $authcode.IndexOf("?code=")
if ($codeStart -gt 0)
{
    $authcode = $authcode.Substring($codeStart+6)
}
$codeEnd = $authcode.IndexOf("&session_state=")
if ($codeEnd -gt 0)
{
    $authcode = $authcode.Substring(0, $codeEnd)
}
Write-Verbose "Using auth code: $authcode"

# Acquire token (using the auth code)
$body = @{grant_type="authorization_code";scope="https://graph.microsoft.com/.default";client_id=$AppId;code=$authcode;redirect_uri=$RedirectUrl}
try
{
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
}
catch
{
    Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
    exit # Failed to obtain a token
}
$global:token = $oauth

# Prepare the request headers
$headers = @{
    'Authorization'  = "$($oauth.token_type) $($oauth.access_token)";
}

# Retrieve my profile to get user Id, then apply that as first chat member
$myProfile = Invoke-RestMethod -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/me" -Method Get
$createChatJSON = $createChatJSON.Replace("%user1%", $myProfile.id)


# Retrieve list of users (we will try to initiate a chat with everyone)
$graphUri = "https://graph.microsoft.com/v1.0/users"
while (![String]::IsNullOrEmpty($graphUri))
{
    Write-Host "Requesting: $graphUri"
    $users = Invoke-RestMethod -Headers $Headers -Uri $graphUri -Method Get
    $graphUri = $users."@odata.nextLink"

    foreach ($user in $users.value)
    {
        # Create a chat with this user
        if (![String]::IsNullOrEmpty($user.mail) -and ($user.id -ne $myProfile.id) )
        {
            Write-Host "Creating chat with: $($user.mail)"
            $json = $createChatJSON.Replace("%user2%", $user.id)
            $newChat = Invoke-RestMethod -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/chats" -Method Post -Body $json -ContentType "application/json"            

            if (![String]::IsNullOrEmpty($newChat.id))
            {
                # Send hello to the chat
                Write-Host "Saying hello to: $($user.mail)"
                $json = @"
{
    "body": {
        "content": "Hello $($user.mail)"
    }
}
"@
                $messagePost = Invoke-RestMethod -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/chats/$($newChat.id)/messages" -Method Post -Body $json -ContentType "application/json"
            }
        }
        # We only create one chat per second to stay within throttling limits (for creating chat messages, that is 2ps, or 20ps across the tenant)
        Start-Sleep -Seconds 1
    }
}

