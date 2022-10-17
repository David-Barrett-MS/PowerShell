param (
    [Parameter(Mandatory=$True,HelpMessage="Application Id (obtained when registering the application in Azure AD")]
    [ValidateNotNullOrEmpty()]
    [string]$AppId = "",

    [Parameter(Mandatory=$True,HelpMessage="Application secret key (obtained when registering the application in Azure AD)")]
    [ValidateNotNullOrEmpty()]
    [string]$AppSecretKey = "",

    [Parameter(Mandatory=$True,HelpMessage="Tenant domain")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantDomain,

    [Parameter(Mandatory=$True,HelpMessage="Name of the report to retrieve")]
    [ValidateNotNullOrEmpty()]
    [string]$UserId = "",

    [Parameter(Mandatory=$False,HelpMessage="If specified, channels will be retrieved as well as chats")]
    [switch]$Channels
)


# Get token
$Body = @{    
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default" 
    client_Id     = $AppId
    Client_Secret = $AppSecretKey
} 
$authResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantDomain/oauth2/v2.0/token" -Method POST -Body $Body -ErrorAction STOP

# Get all messages
$Headers = @{
    'Authorization' = "Bearer $($authResponse.access_token)"
}
$chatMessages = Invoke-RestMethod -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/users/$UserId/chats/getAllMessages" -Method Get

# Output data
if (![String]::IsNullOrEmpty($SavePath))
{
    $currentDate = get-date -Format d
    $currentDate = $currentDate.Replace('/','-')
    $chatMessages | out-file -FilePath "$SavePath\Chats-$UserId-$currentDate.csv"
}
else
{
    $report
}

if ($Channels)
{
    $channelMessages = Invoke-RestMethod -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/users/$UserId/chats/getAllMessages" -Method Get
}