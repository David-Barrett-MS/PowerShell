#
# uploadUserPhoto.ps1
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
Uploads a profile photo for a user.

.DESCRIPTION
This script demonstrates how to upload a profile picture for a user (delegate or application permissions).
https://docs.microsoft.com/en-us/graph/api/profilephoto-update?view=graph-rest-1.0&tabs=http

.EXAMPLE
.\uploadUserPhoto.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -RedirectUrl "<RedirectUrl>" -Photo "<path to photo jpeg>"

#>


param (
	[Parameter(Mandatory=$False,HelpMessage="Application Id (obtained when registering the application in Azure AD")]
	[ValidateNotNullOrEmpty()]
	[string]$AppId,

	[Parameter(Mandatory=$False,HelpMessage="Application secret key (obtained when registering the application in Azure AD).  If not specified, delegate permissions are assumed.")]
	[ValidateNotNullOrEmpty()]
	[string]$AppSecretKey,

	[Parameter(Mandatory=$False,HelpMessage="Redirect Url (specified when registering the application in Azure AD, use localhost).  Required for delegate permissions.")]
	[ValidateNotNullOrEmpty()]
	[string]$RedirectUrl,

	[Parameter(Mandatory=$False,HelpMessage="Tenant Id.  Default is common, which requires the application to be registered for multi-tenant use (and consented to in target tenant)")]
	[ValidateNotNullOrEmpty()]
	[string]$TenantId = "common",

	[Parameter(Mandatory=$False,HelpMessage="Mailbox (not required if using delegate permissions)")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox = "me",

	[Parameter(Mandatory=$False,HelpMessage="Path to the photo that will be attached")]
	[ValidateNotNullOrEmpty()]
    [string]$Photo
)


# Check the photo is valid
$photoFile = Get-Item $Photo
if (!$photoFile)
{
    Write-Host "Failed to read photo: $Photo" -ForegroundColor Red
    exit
}

if ($Mailbox -eq "me")
{
    $graphUrl = "https://graph.microsoft.com/v1.0/me/photo/`$value"

    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize?client_id=$AppId&response_type=code&redirect_uri=$RedirectUrl&response_mode=query&scope=openid%20profile%20email%20offline_access%20$Scope"
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
}
else
{
    $graphUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/photo/`$value"

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
}
Write-Host "Successfully obtained OAuth token" -ForegroundColor Green

# Read the photo
$fileStream = New-Object -TypeName System.IO.FileStream -ArgumentList ($photoFile.VersionInfo.FileName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
$fileReader = New-Object -TypeName System.IO.BinaryReader -ArgumentList $fileStream
if (!$fileReader) { exit }

$photoBytes = $fileReader.ReadBytes($photoFile.Length)
$fileReader.Dispose()
$fileStream.Dispose()

# Prepare the request headers
$headers = @{
    'Authorization'  = "$($oauth.token_type) $($oauth.access_token)";
    'Content-Type'   = 'image/jpeg';
}

# PUT the photo
try
{
    Write-Host "Sending request to: $graphUrl" -ForegroundColor White
    if ($psversiontable.PSVersion.Major -gt 6)
    {
        $global:uploadResults = Invoke-WebRequest -Method Put -Uri $graphUrl -Body $photoBytes -Headers $headers -SkipHeaderValidation
    }
    else
    {
        $global:uploadResults = Invoke-WebRequest -Method Put -Uri $graphUrl -Body $photoBytes -Headers $headers
    }
}
catch
{
    Write-Host "Failed to upload photo" -ForegroundColor Red
    $Error[0]
    exit
}

Write-Host "Photo upload succeeded." -ForegroundColor Green