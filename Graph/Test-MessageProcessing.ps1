#
# Test-MessageProcessing.ps1
#
# By David Barrett, Microsoft Ltd. 2025. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

<#
.SYNOPSIS
Simulates actions taken by an app automatically processing messages in a mailbox.

.DESCRIPTION
Script used for testing various Graph calls interacting with messages in a mailbox.

.EXAMPLE

Delegate permissions (requires Mail.ReadWrite):
.\Test-MessageProcessing.ps1 -AppId $clientId -TenantId $tenantId -TestCount 1

Application permissions (requires Mail.ReadWrite):
.\Test-MessageProcessing.ps1 -mailbox $Mailbox -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey  -TestCount 1

#>


param (
    [Parameter(Mandatory=$False,HelpMessage="Application Id (obtained when registering the application in Azure AD).")]
    [ValidateNotNullOrEmpty()]
    [string]$AppId,

    [Parameter(Mandatory=$False,HelpMessage="Application secret key (obtained when registering the application in Azure AD).  If not specified, delegate permissions are assumed.")]
    [ValidateNotNullOrEmpty()]
    [string]$AppSecretKey,

    [Parameter(Mandatory=$False,HelpMessage="Redirect Url (specified when registering the application in Azure AD, use localhost).  Required for delegate permissions.")]
    [ValidateNotNullOrEmpty()]
    [string]$RedirectUrl = "http://localhost/code",

    [Parameter(Mandatory=$False,HelpMessage="Tenant Id.  Default is common, which requires the application to be registered for multi-tenant use (and consented to in target tenant).")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId = "common",

    [Parameter(Mandatory=$False,HelpMessage="Mailbox (me to access own mailbox when using delegate permissions).")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox = "me",

    [Parameter(Mandatory=$False,HelpMessage="If specified, output will be logged to file (using same name as this script).")]
    [switch]$LogToFile,

    [Parameter(Mandatory=$False,HelpMessage="Number of times to run the test.")]
    [int]$TestCount = 10
)

$script:ScriptVersion = "1.0.0"
if ($LogToFile) {
    $script:logFile = "$($MyInvocation.InvocationName).log"
}

Function LogToFile([string]$LogEntry)
{
	if ( [String]::IsNullOrEmpty($script:logFile) ) { return }
	$LogEntry | Out-File $script:logFile -Append
}

Function Log([string]$Details, [System.ConsoleColor]$ForegroundColor = "Gray")
{
    $logEntry = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details"
    Write-Host $logEntry -ForegroundColor $ForegroundColor
    LogToFile $logEntry
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting"



function GetAppAuthToken()
{
    # Acquire token for application permissions
    $body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$AppId;client_secret=$AppSecretKey}
    try
    {
        $script:oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
        $script:token_expires = (Get-Date).AddSeconds($oauth.expires_in)
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
        exit # Failed to obtain a token
    }
    $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/"
}

function RenewOAuthToken()
{
    # If using app flow, we just obtain a new token

    if (![String]::IsNullOrEmpty($AppSecretKey))
    {
        GetAppAuthToken
        return
    }

    # If using delegate flow, we need to refresh the token
    $body = @{grant_type="refresh_token";scope="https://graph.microsoft.com/.default";client_id=$AppId;refresh_token=$oauth.refresh_token}
    try
    {
        $script:oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
        $script:token_expires = (Get-Date).AddSeconds($oauth.expires_in)
    }
    catch
    {
        Write-Host "Failed to renew OAuth token" -ForegroundColor Red
        exit # Failed to renew the token
    }
    Write-Host "Successfully renewed OAuth token" -ForegroundColor Green
}

function UpdateHeaders()
{
    # Prepare the request headers

    if ($null -eq $script:headers -or $script.headers.Count -eq 0)
    {
        $script:headers = @{
            'Authorization'  = "$($oauth.token_type) $($oauth.access_token)";
            'Content-Type'   = 'application/json';
        }    
    }

    if ($script:token_expires -lt (Get-Date).AddMinutes(5))
    {
        RenewOAuthToken
    }

}

function GET($url)
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = Invoke-RestMethod -Method Get -Uri $url -Headers $script:headers
    }
    catch
    {
        Write-Host "Failed to obtain data from URL: $url" -ForegroundColor Red
        exit
    }
    return $result
}

function PATCH($url, $body)
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = Invoke-RestMethod -Method Patch -Uri $url -Headers $script:headers -Body $body
    }
    catch
    {
        Write-Host "Failed to obtain data from URL: $url" -ForegroundColor Red
        exit
    }
    return $result
}

function DELETE($url)
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = Invoke-RestMethod -Method Delete -Uri $url -Headers $script:headers
    }
    catch
    {
        Write-Host "Failed to obtain data from URL: $url" -ForegroundColor Red
        exit
    }
    return $result
}

function GetMailboxFolder($queryString)
{
    $url = $script:graphBaseUrl + "mailFolders/$queryString"
    return GET $url
}

function GetFolderMessages($folderId, $queryString)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$queryString"
    return GET $url
}

function MarkMessageAsRead($messageId)
{
    $url = $script:graphBaseUrl + "messages/$messageId"
    $body = @{
        isRead = $true
    } | ConvertTo-Json
    return PATCH $url $body
}

function GetFolderMessageById($folderId, $messageId, $queryString)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId/$queryString"
    return GET $url
}

function GetFolderMessageByIdAttachments($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId/attachments"
    return GET $url
}

function DeleteFolderMessageById($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId"
    return DELETE $url
}

## Authentication

if ([String]::IsNullOrEmpty($AppSecretKey))
{
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
    $graphBaseUrl = "https://graph.microsoft.com/v1.0/me/"
}
else
{
    GetAppAuthToken
}
Write-Host "Successfully obtained OAuth token" -ForegroundColor Green

# GET /mailFolders/?$filter=displayName+eq+'INBOX'
$inbox = GetMailboxFolder("?`$filter=displayName eq 'INBOX'")
if ($null -eq $inbox)
{
    Write-Host "Failed to obtain inbox folder" -ForegroundColor Red
    exit # Failed to obtain inbox folder
}
$inboxId = $inbox.value.id
Write-Host "Inbox folder id: $($inboxId)" -ForegroundColor Green

# GET /mailfolders/inbox/messages/?$filter=isRead+eq+false&$count=true&$top=20&$orderBy=receivedDateTime&select=internetMessageId,id
$unreadMessages = GetFolderMessages $inboxId "?`$filter=isRead eq false&`$count=true&`$top=20&`$orderBy=receivedDateTime&select=internetMessageId,id"
if ($null -eq $unreadMessages)
{
    Write-Host "Failed to obtain unread messages" -ForegroundColor Red
    exit # Failed to obtain unread messages
}
if ($unreadMessages.value.Count -eq 0)
{
    Write-Host "No unread messages to process" -ForegroundColor Yellow
    exit # No unread messages to process
}
Write-Host "$($unreadMessages.value.Count) unread messages found" -ForegroundColor Green

$global:messageToProcess = $unreadMessages.value[0]
# PATCH message (mark as read)
if ($null -eq $messageToProcess)
{
    Write-Host "Failed to obtain message to process" -ForegroundColor Red
    exit # Failed to obtain message to process
}
$global:updatedMessage = MarkMessageAsRead $messageToProcess.id
if ($null -eq $updatedMessage)
{
    Write-Host "Failed to update message" -ForegroundColor Red
    exit # Failed to update message
}
Write-Host "Message marked as read: $($updatedMessage.subject)" -ForegroundColor Green


# GET /mailFolders/?$filter=displayName+eq+'INBOX'
$inbox = GetMailboxFolder("?`$filter=displayName eq 'INBOX'")
if ($null -eq $inbox)
{
    Write-Host "Failed to obtain inbox folder" -ForegroundColor Red
    exit # Failed to obtain inbox folder
}
$inboxId = $inbox.value.id
Write-Host "Inbox folder id: $($inboxId)" -ForegroundColor Green

# GET /mailfolders/inbox(by Id)/MessageID?$select=*
$message = GetFolderMessageById $inboxId $messageToProcess.id "?`$select=*"
if ($null -eq $message)
{
    Write-Host "Failed to obtain message by id: $($messageToProcess.id)" -ForegroundColor Red
    exit # Failed to obtain message
}

# GET /mailfolders/inbox(by Id)/MessageID/attachments
$global:attachments = GetFolderMessageByIdAttachments $inboxId $messageToProcess.id
if ($null -eq $attachments)
{
    Write-Host "Failed to obtain message attachments" -ForegroundColor Red
    exit # Failed to obtain message attachments
}
Write-Host "$($attachments.value.Count) attachments found" -ForegroundColor Green

# DELETE /mailfolders/inbox(by Id)/MessageID
#DeleteFolderMessageById $inboxId $messageToProcess.id

Write-Host "Completed"
Write-Host "`$messageToProcess contains the message that was processed"
Write-Host "`$updatedMessage contains the message after marked as read"
Write-Host "`$attachments contains the attachments for the message that was processed"