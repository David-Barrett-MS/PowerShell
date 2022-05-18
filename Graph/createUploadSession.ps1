#
# createUploadSession.ps1
#
# By David Barrett, Microsoft Ltd. 2021. Use at your own risk.  No warranties are given.
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
Create a Draft message and add one or multiple large attachments.

.DESCRIPTION
This script demonstrates how to create a draft message and add a large attachment using createUploadSession.
https://docs.microsoft.com/en-us/graph/api/attachment-createuploadsession

.EXAMPLE
.\createUploadSession.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<AppSecretKey>" -Mailbox "<mailbox>"

#>

param (
	[Parameter(Mandatory=$False,HelpMessage="Application Id (obtained when registering the application in Azure AD).")]
	[ValidateNotNullOrEmpty()]
	[string]$AppId,

	[Parameter(Mandatory=$False,HelpMessage="Application secret key (obtained when registering the application in Azure AD).")]
	[ValidateNotNullOrEmpty()]
	[string]$AppSecretKey,

	[Parameter(Mandatory=$False,HelpMessage="Tenant Id.")]
	[ValidateNotNullOrEmpty()]
	[string]$TenantId,

	[Parameter(Mandatory=$False,HelpMessage="Mailbox.")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,

	[Parameter(Mandatory=$False,HelpMessage="The file(s) that will be attached.")]
	[ValidateNotNullOrEmpty()]
    $FileAttachments
)
$script:ScriptVersion = "1.0.1"

$graphUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/"

# Validate the attachment(s)
foreach ($filePath in $FileAttachments)
{
    $attachmentFile = Get-Item $filePath
    if (!$attachmentFile)
    {
        Write-Host "Failed to locate attachment: $filePath" -ForegroundColor Red
        exit
    }
}

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


# Create draft message to the mailbox owner (i.e. send message to self)
$createUrl = "$($graphUrl)messages"
$createMessageJson = "{
    ""subject"":""Test Message with attachment"",
    ""importance"":""Low"",
    ""body"":{
        ""contentType"":""HTML"",
        ""content"":""Please check attachment(s) and discard.""
    },
    ""toRecipients"":[
        {
            ""emailAddress"":{
                ""address"":""$Mailbox""
            }
        }
    ]
}"

try
{
    Write-Host "Sending request to: $createUrl" -ForegroundColor White
    $createMessageResults = Invoke-RestMethod -Method Post -Uri $createUrl -Headers $authHeader -Body $createMessageJson -ContentType "application/json"
}
catch
{
    Write-Host "Failed to create message" -ForegroundColor Red
    exit
}

if ([String]::IsNullOrEmpty(($createMessageResults.id)))
{
    Write-Host "Failed to read message Id of created message." -ForegroundColor Red
    exit
}


# Upload the attachments

foreach ($filePath in $FileAttachments)
{
    Write-Host "Starting upload session for $filePath"
    $attachmentFile = Get-Item $filePath

    # Create upload session
    $createUploadSessionUrl = "$($graphUrl)messages/$($createMessageResults.id)/attachments/createUploadSession"
    $createUploadSessionJson = "{
        ""AttachmentItem"": {
            ""attachmentType"": ""file"",
            ""name"": ""$($attachmentFile.Name)"", 
            ""size"": $($attachmentFile.Length)
        }
    }"

    try
    {
        Write-Host "Sending request to: $createUploadSessionUrl" -ForegroundColor White
        $createUploadSessionResults = Invoke-RestMethod -Method Post -Uri $createUploadSessionUrl -Headers $authHeader -Body $createUploadSessionJson -ContentType "application/json"
    }
    catch
    {
        Write-Host "Failed to create upload session" -ForegroundColor Red
        exit 
    }

    if ([String]::IsNullOrEmpty(($createUploadSessionResults.uploadUrl)))
    {
        Write-Host "Failed to create upload session." -ForegroundColor Red
        exit
    }

    # Upload the attachment (in blocks)
    $fileStream = New-Object -TypeName System.IO.FileStream -ArgumentList ($attachmentFile.VersionInfo.FileName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
    $fileReader = New-Object -TypeName System.IO.BinaryReader -ArgumentList $fileStream
    if (!$fileReader) { exit }


    $offset = 0
    $blockSize = 3Mb

    while ($offset -lt $attachmentFile.Length)
    {
        # Read the next block of data from the file
        $blockBytes = $fileReader.ReadBytes($blockSize)

        # Prepare the request headers
        $headers = @{
            #'Content-Length'=$blockBytes.Length;
            'Content-Range'="bytes $offset-$($offset+$blockBytes.Length-1)/$($attachmentFile.Length)"
        }
        if  ($PSVersionTable.PSVersion.Major -lt 7)
        {
            $headers.Add("Content-Length",$blockBytes.Length)
        }
        $offset += $blockBytes.Length

        # Upload the data
        try
        {
            Write-Host "Uploading to: $($createUploadSessionResults.uploadUrl)" -ForegroundColor White
            if ($PSVersionTable.PSVersion.Major -lt 7)
            {
                $global:uploadResults = Invoke-WebRequest -Method Put -Uri $createUploadSessionResults.uploadUrl -Body $blockBytes -Headers $headers -ContentType "application/octet-stream"
            }
            else
            {
                $global:uploadResults = Invoke-WebRequest -Method Put -Uri $createUploadSessionResults.uploadUrl -Body $blockBytes -Headers $headers -ContentType "application/octet-stream" -SkipHeaderValidation
            }
        }
        catch
        {
            Write-Host "Failed to upload file: $filePath" -ForegroundColor Red
            $fileReader.Dispose()
            $fileStream.Dispose()
            exit 
        }
    }
    Write-Host "Finished uploading: $filePath"


    $fileReader.Dispose()
    $fileStream.Dispose()
}


Write-Host "Message creation and attachment upload succeeded." -ForegroundColor Green