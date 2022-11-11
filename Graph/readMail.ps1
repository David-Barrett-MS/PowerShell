#
# readMail.ps1
#
# By David Barrett, Microsoft Ltd. 2022. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

<#
.SYNOPSIS
Reads messages from a mailbox using Graph.

.DESCRIPTION
This script demonstrates how to read messages using the Graph API.
https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http

An application must be registered with the Microsoft Graph Mail.Read permission assigned.  This script uses Application permissions
and a secret key - pass the relevant information via parameters.

.EXAMPLE
.\readMail.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<AppSecretKey>" -Mailbox "<mailbox>"

#>


param (
	[Parameter(Mandatory=$False,HelpMessage="Application Id (obtained when registering the application in Azure AD")]
	[ValidateNotNullOrEmpty()]
	[string]$AppId,

	[Parameter(Mandatory=$False,HelpMessage="Application secret key (obtained when registering the application in Azure AD")]
	[ValidateNotNullOrEmpty()]
	[string]$AppSecretKey,

	[Parameter(Mandatory=$False,HelpMessage="Tenant Id")]
	[ValidateNotNullOrEmpty()]
	[string]$TenantId,

	[Parameter(Mandatory=$False,HelpMessage="Mailbox")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,

	[Parameter(Mandatory=$False,HelpMessage="Determines how many messages will be retrieved.")]
	$Top = 10
)


$graphUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/messages"

# Acquire token for application permissions
$body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$AppId;client_secret=$AppSecretKey}
try
{
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
}
catch
{
    Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
    exit # Failed to obtain a token
}
$authHeader = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
Write-Host "Successfully obtained OAuth token" -ForegroundColor Green

# Retrieve messages
$global:messages = Invoke-RestMethod -Headers $authHeader -Uri $graphUrl -Method Get

# Output data
if (![String]::IsNullOrEmpty($SavePath))
{
    $currentDate = get-date -Format d
    $currentDate = $currentDate.Replace('/','-')
    $messages | out-file -FilePath "$SavePath\Messages-$UserId-$currentDate.json"
}
else
{
    $messages.value
}
