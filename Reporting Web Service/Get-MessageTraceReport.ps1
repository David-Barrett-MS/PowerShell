#
# Get-MessageTraceReport.ps1
#
# By David Barrett, Microsoft Ltd. 2023. Use at your own risk.  No warranties are given.
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
Retrieves a message trace report from the Office 365 Reporting Web Service.

.DESCRIPTION
This script demonstrates how to retrieve a message trace report from the Office 365 Reporting Web Service.  Implements OAuth via auth code only (certificate auth not yet implemented, basic auth is not supported).  Not all parameters available in the API are implemented, this script is purely for testing.
For application registration instructions, please see https://learn.microsoft.com/en-us/previous-versions/office/developer/o365-enterprise-developers/jj984325(v=office.15)#register-your-application-in-azure-ad

.EXAMPLE
PS> .\Get-MessageTraceReport.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<SecretKey>" -ReportSavePath "c:\Reports"

This will get the message trace report for the previous 48 hour period (default).

.EXAMPLE
PS> $messageTrace = .\Get-MessageTraceReport.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -AppSecretKey "<SecretKey>" -ReportSavePath "c:\Reports" -OutHost -StartDate 2023/02/17 -EndDate 2023/02/19

Retrieves the message trace report for the specified period and assigns the result to $messageTrace (which will be an XmlElement that can be queried further using PowerShell or .Net XML functions).
#>

param (
	[Parameter(Mandatory=$True,HelpMessage="Application Id (obtained when registering the application in Azure AD.")]
	[ValidateNotNullOrEmpty()]
	[string]$AppId,

	[Parameter(Mandatory=$True,HelpMessage="Application secret key (obtained when registering the application in Azure AD.")]
	[ValidateNotNullOrEmpty()]
	[string]$AppSecretKey,

    [Parameter(Mandatory=$False,HelpMessage="Required when using application permissions.  Please note that certificate auth requires the MSAL dll to be available.")]
    $OAuthCertificate = $null,
    
	[Parameter(Mandatory=$True,HelpMessage="Tenant Id.")]
	[ValidateNotNullOrEmpty()]
	[string]$TenantId,

    [Parameter(Mandatory=$False,HelpMessage="The redirect Uri of the Azure registered application.")]
    [string]$RedirectUrl = "http://localhost/code",

	[Parameter(Mandatory=$False,HelpMessage="Report save path (reports are prepended by the StartDate-EndDate or current date if no date parameters specified).  If missing, current folder is used.")]
	[ValidateNotNullOrEmpty()]
	[string]$ReportSavePath,

	[Parameter(Mandatory=$False,HelpMessage="When specified, the message trace report will be output to the pipeline.")]
	[ValidateNotNullOrEmpty()]
	[switch]$OutHost,

	#[Parameter(Mandatory=$False,HelpMessage="The Internet MessageID header of the message.")]
	#[ValidateNotNullOrEmpty()]
	#[string]$MessageId,

	#[Parameter(Mandatory=$False,HelpMessage="An identifier used to get the detailed message transfer trace information.")]
	#[ValidateNotNullOrEmpty()]
	#[string]$MessageTraceId,

	[Parameter(Mandatory=$False,HelpMessage="This field is used to limit the report period.")]
    [ValidateNotNullOrEmpty()]
	[DateTime]$StartDate = [DateTime]::MinValue,

	[Parameter(Mandatory=$False,HelpMessage="This field is used to limit the report period.")]
    [ValidateNotNullOrEmpty()]
	[DateTime]$EndDate = [DateTime]::MinValue

)
$script:ScriptVersion = "0.1.0"
Write-Host "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting" -ForegroundCOlor Green

function LoadLibraries()
{
    param (
        [bool]$searchProgramFiles,
        $dllNames,
        [ref]$dllLocations = @()
    )
    # Attempt to find and load the specified libraries

    foreach ($dllName in $dllNames)
    {
        # First check if the dll is in current directory
        LogDebug "Searching for DLL: $dllName"
        $dll = $null
        try
        {
            $dll = Get-ChildItem $dllName -ErrorAction SilentlyContinue
        }
        catch {}

        if ($searchProgramFiles)
        {
            if ($dll -eq $null)
            {
	            $dll = Get-ChildItem -Recurse "C:\Program Files (x86)" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            if (!$dll)
	            {
		            $dll = Get-ChildItem -Recurse "C:\Program Files" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            }
            }
        }

        if ($dll -eq $null)
        {
            Log "Unable to load locate $dll" Red
            return $false
        }
        else
        {
            try
            {
		        LogVerbose ([string]::Format("Loading {2} v{0} found at: {1}", $dll.VersionInfo.FileVersion, $dll.VersionInfo.FileName, $dllName))
		        Add-Type -Path $dll.VersionInfo.FileName
                if ($dllLocations)
                {
                    $dllLocations.value += $dll.VersionInfo.FileName
                    ReportError
                }
            }
            catch
            {
                ReportError "LoadLibraries"
                return $false
            }
        }
    }
    return $true
}

function GetTokenWithCertificate
{
    # We use MSAL with certificate auth
    if (!$script:msalApiLoaded)
    {
        $msalLocation = @()
        $script:msalApiLoaded = $(LoadLibraries -searchProgramFiles $false -dllNames @("Microsoft.Identity.Client.dll") -dllLocations ([ref]$msalLocation))
        if (!$script:msalApiLoaded)
        {
            Log "Failed to load MSAL.  Cannot continue with certificate authentication." Red
            exit
        }
    }   

    $cca1 = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($OAuthClientId)
    $cca2 = $cca1.WithCertificate($OAuthCertificate)
    $cca3 = $cca2.WithTenantId($OAuthTenantId)
    $cca = $cca3.Build()

    $scopes = New-Object System.Collections.Generic.List[string]
    $scopes.Add("https://outlook.office365.com/.default")
    $acquire = $cca.AcquireTokenForClient($scopes)
    $authResult = $acquire.ExecuteAsync().Result
    $script:oauthToken = $authResult
    $script:oAuthAccessToken = $script:oAuthToken.AccessToken
}

if ( $OAuthCertificate -eq $null )
{
    # Unusually (for auth code flow), this auth requires a secret key as well as auth code to obtain the access token

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
    else
    {
        Write-Host "Failed to obtain auth code" -ForegroundColor Red
        exit # Failed to obtain a token
    }

    $codeEnd = $authcode.IndexOf("&session_state=")
    if ($codeEnd -gt 0)
    {
        $authcode = $authcode.Substring(0, $codeEnd)
    }
    Write-Verbose "Using auth code: $authcode"

    # Acquire token (using the auth code) - even though we know the tenant Id, we need to use the common endpoint to obtain our token
    $body = @{resource="https://outlook.office365.com";grant_type="authorization_code";client_id=$AppId;code=$authcode;redirect_uri=$RedirectUrl;client_secret=$AppSecretKey}
    try
    {
        $oauth = Invoke-RestMethod -Method Post -Uri https://login.windows.net/common/oauth2/token -Body $body
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
        exit # Failed to obtain a token
    }
}
else
{
    # Acquire token using application permissions (requires certificate auth)
    # THIS IS NOT YET IMPLEMENTED
    # resource=https%3A%2F%2Foutlook.office365.com&amp;client_id={your_app_client_id}&amp;grant_type=client_credentials&amp;client_assertion_type=urn%3Aietf%3Aparams%3Aoauth%3Aclient-assertion-type%3Ajwt-bearer&amp;client_assertion={encoded_signed_JWT_token}
    Write-Host "Certificate/application auth is not yet implemented." -ForegroundColor Red
    Exit
}

$token = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
Write-Host "Successfully obtained OAuth token" -ForegroundColor Green

# Now we have our OAuth token, we can send requests to the reporting web service endpoint

if (![String]::IsNullOrEmpty($ReportSavePath) -and !$ReportSavePath.EndsWith("\"))
{
    $ReportSavePath = "$ReportSavePath\"
}

# https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace[?ODATA options]
$reportUri = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace?"
$outputFile = "$ReportSavePath$([DateTime]::Today.ToString("yyyyMMdd"))MessageTrace.xml"
if ($StartDate -gt [DateTime]::MinValue -or $EndDate -gt [DateTime]::MinValue)
{
    if ($EndDate -eq [DateTime]::MinValue)
    {
        Write-Host "EndDate required when StartDate is used" -ForegroundColor Red
        Exit
    }
    if ($StartDate -eq [DateTime]::MinValue)
    {
        Write-Host "StartDate required when EndDate is used" -ForegroundColor Red
        Exit
    }
    # Example Uri: https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace?\$filter=StartDate%20eq%20datetime%272020-02-01T00:00:00Z%27%20and%20EndDate%20eq%20datetime%272020-02-06T00:00:00Z
    $reportUri = "$($reportUri)`$filter=StartDate eq datetime%27$($StartDate.ToString("o"))%27 and EndDate eq datetime%27$($EndDate.ToString("o"))%27"
    $outputFile = "$ReportSavePath$($StartDate.ToString("yyyyMMdd"))-$($EndDate.ToString("yyyyMMdd")) Message Trace.xml"
}

$results = $null

try
{
    Write-Host "GET $reportUri" -ForegroundColor Gray
    
    $results = Invoke-RestMethod -Method Get -Uri $reportUri -Headers $token -PassThru -OutFile $outputFile
    Write-Host "Saved to: $outputFile" -ForegroundColor Green
}
catch
{
}

if ($OutHost)
{
    # Write the output to the pipeline
    $results
}