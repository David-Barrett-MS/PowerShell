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
# Send simple HTML message to self using application permissions
.\sendMail.ps1 -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey -Mailbox $Mailbox

# Send adaptive card in message to self
$cardContent = Get-Content "c:\temp\adaptiveCard.json" -Raw
.\sendMail.ps1 -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey -Mailbox $Mailbox -AdaptiveCardPayload $cardContent

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

    [Parameter(Mandatory=$False,HelpMessage="If specified, message is sent to this recipient instead of the mailbox owner.")]
    [string]$Recipient = "",

    [Parameter(Mandatory=$False,HelpMessage="If specified, this picture will be attached to the message.")]
    [string]$AttachPicture,

    [Parameter(Mandatory=$False,HelpMessage="If specified, this is the HTML body of the message.")]
    [string]$MessageHTML = "",

    [Parameter(Mandatory=$False,HelpMessage="If specified, this Adaptive Card JSON is embedded in the message body.")]
    [string]$AdaptiveCardPayload = "",

	[Parameter(Mandatory=$False,HelpMessage="If specified, message will be saved to Sent Items.")]
	[Switch]$SaveToSentItems

)

function Set-DefaultMessageHtml
{
    param([string]$Content)

    if ([String]::IsNullOrEmpty($script:MessageHTML))
    {
        $script:MessageHTML = $Content
    }
}


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

$attachmentList = @()
if (-not [String]::IsNullOrEmpty($AttachPicture))
{
    # Add attachment
    $pictureType = $AttachPicture.Substring($AttachPicture.Length-3).ToLower()
    if ($pictureType -ne "jpg" -and $pictureType -ne "png")
    {
        Write-Host "Attachment must be jpg or png." -ForegroundColor Red
    }
    else
    {
        $pictureFile = Get-Item $AttachPicture
        if (!$pictureFile)
        {
            Write-Host "Failed to read picture: $AttachPicture" -ForegroundColor Red
        }
        else
        {
            # Read the byte data for the picture
            Write-Host "Reading $($pictureFile.VersionInfo.FileName)" -ForegroundColor Green
            $fileStream = New-Object -TypeName System.IO.FileStream -ArgumentList ($pictureFile.VersionInfo.FileName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            $fileReader = New-Object -TypeName System.IO.BinaryReader -ArgumentList $fileStream
            if (!$fileReader) { exit }

            $pictureBytes = $fileReader.ReadBytes($pictureFile.Length)
            $fileReader.Dispose()
            $fileStream.Dispose()

            # Convert the data into a Base64 string
            $attachBytesBase64 = [Convert]::ToBase64String($pictureBytes)

            # Add the attachment JSON
            $attachmentList += [pscustomobject]@{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                name = $pictureFile.Name
                contentId = 'image001.png@01D82E0D.6377FC90'
                contentType = "image/$pictureType"
                contentBytes = $attachBytesBase64
            }
            # If no message HTML has been provided, embed the image in a simple body
            Set-DefaultMessageHtml "<html><body><p>This is the image: <img id='Picture_x0020_1' src='cid:image001.png@01D82E0D.6377FC90'></p></body></html>"
        }
    }
}

if (-not [String]::IsNullOrEmpty($AdaptiveCardPayload))
{
    try
    {
        $cardJsonObject = $AdaptiveCardPayload | ConvertFrom-Json
    }
    catch
    {
        Write-Host "Adaptive Card payload must be valid JSON." -ForegroundColor Red
        exit
    }

    $cardJson = $cardJsonObject | ConvertTo-Json -Depth 20 -Compress
    $adaptiveCardScript = "<script type='application/adaptivecard+json'>$cardJson</script>"

    $originalMessageHtml = $MessageHTML
    Set-DefaultMessageHtml "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body><p>This message contains an Adaptive Card.</p>$adaptiveCardScript</body></html>"

    if ($MessageHTML -eq $originalMessageHtml)
    {
        $bodyCloseRegex = [System.Text.RegularExpressions.Regex]::new("</body>", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($bodyCloseRegex.IsMatch($MessageHTML))
        {
            $MessageHTML = $bodyCloseRegex.Replace($MessageHTML, "$adaptiveCardScript</body>", 1)
        }
        else
        {
            $MessageHTML = "$MessageHTML$adaptiveCardScript"
        }
    }
}

$bodyContentType = "Text"
$bodyContent = "This is a test message sent using Graph sendMail."

if ($MessageHTML)
{
    $bodyContentType = "HTML"
    $bodyContent = $MessageHTML
}

$targetRecipient = $Mailbox
if ($Recipient)
{
    $targetRecipient = $Recipient
}

$messagePayload = @{
    message = @{
        subject = "Test Message"
        importance = "Low"
        body = @{
            contentType = $bodyContentType
            content = $bodyContent
        }
        toRecipients = @(
            @{
                emailAddress = @{
                    address = $targetRecipient
                }
            }
        )
    }
    saveToSentItems = [bool]$SaveToSentItems
}

if ($attachmentList.Count -gt 0)
{
    $messagePayload.message.attachments = $attachmentList
}

$sendMessageJson = $messagePayload | ConvertTo-Json -Depth 20

$sendMessageJson

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
