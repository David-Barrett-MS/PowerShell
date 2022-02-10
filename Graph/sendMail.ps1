#
# sendMail.ps1
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
Send a message (to self) using Graph sendMail.

.DESCRIPTION
This script demonstrates how to send an email using the Graph API.
https://docs.microsoft.com/en-us/graph/api/user-sendmail

An application must be registered with the Microsoft Graph Mail.Send permission assigned.  This script uses Application permissions
and a secret key - pass the relevant information via parameters.  A test message is sent from the specified mailbox back to the same
mailbox.

.EXAMPLE
.\sendMail.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<AppSecretKey>" -Mailbox "<mailbox>"

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
	[string]$Mailbox
)


$graphUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/"

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


# Send message to the mailbox owner (i.e. send message to self)
$createUrl = "$($graphUrl)sendMail"
$sendMessageJson = "
{
    ""message"": {
        ""subject"":""Test Message"",
        ""importance"":""Low"",
        ""body"":{
            ""contentType"":""Text"",
            ""content"":""This is a test message sent using Graph sendMail.""
        },
        ""toRecipients"":[
            {
                ""emailAddress"":{
                    ""address"":""$Mailbox""
                }
            }
        ]
    },
    ""saveToSentItems"": ""true""
}"

try
{
    Write-Host "Sending request to: $createUrl" -ForegroundColor White
    $global:sendMessageResults = Invoke-RestMethod -Method Post -Uri $createUrl -Headers $authHeader -Body $sendMessageJson -ContentType "application/json"
}
catch
{
    Write-Host "Failed to send message" -ForegroundColor Red
    exit
}
