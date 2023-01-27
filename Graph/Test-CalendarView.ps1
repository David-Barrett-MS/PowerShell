#
# Test-CalendarView.ps1
#
# By David Barrett, Microsoft Ltd. 2022. Use at your own risk.  No warranties are given.
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
Tests Graph calendarView and collects traces for troubleshooting purposes.

.DESCRIPTION
This script performs a series of requests to Graph calendarView to isolate problematic items.

.EXAMPLE
.\Test-CalendarView.ps1 -AppId "<AppId>" -TenantId "<TenantId>" -RedirectUrl "<RedirectUrl>" -StartDate xx -EndDate xx

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
	[string]$RedirectUrl = "http://localhost",

	[Parameter(Mandatory=$False,HelpMessage="Tenant Id.  Default is common, which requires the application to be registered for multi-tenant use (and consented to in target tenant)")]
	[ValidateNotNullOrEmpty()]
	[string]$TenantId = "common",

	[Parameter(Mandatory=$False,HelpMessage="Mailbox (not required if using delegate permissions)")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox = "me",

	[Parameter(Mandatory=$False,HelpMessage="Folder to which diagnostics will be saved (defaults to current folder).")]
	[ValidateNotNullOrEmpty()]
	[string]$OutputFolder = "",

    [Parameter(Mandatory=$False,HelpMessage="Start of the time range to check.")]
    [datetime]$StartDate = [DateTime]::Today,
	
    [Parameter(Mandatory=$False,HelpMessage="End of the time range to check.")]
    [datetime]$EndDate = [DateTime]::Today.AddDays(1),

    [Parameter(Mandatory=$False,HelpMessage="If the time range is longer than this many hours, the analysis will be split into periods of this length.")]
    $AnalysisPeriod = 24,

	[Parameter(Mandatory=$False,HelpMessage="Query string that will be appended to calendar view request.")]
	[ValidateNotNullOrEmpty()]
    [string]$ViewQuery #= "`$filter=contains(subject, 'test')"
)

# Create Log file
$script:ScriptVersion = "1.0.0"
$script:logFile = "$($MyInvocation.InvocationName).log"
if (![String]::IsNullOrEmpty($OutputFolder))
{
    if (!$OutputFolder.EndsWith("\"))
    {
        $OutputFolder = "$OutputFolder\"
    }
    $logFile = "$OutputFolder$($script:logFile)"
}

Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($script:logFile) ) { return }
	"$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details" | Out-File $LogFile -Append
}
LogToFile "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting"


if ($Mailbox -eq "me")
{
    $graphUrl = "https://graph.microsoft.com/v1.0/me/calendarView" #?startDateTime=<STARTDATE>&endDateTime=<ENDDATE>"

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
}
else
{
    $graphUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/calendarView" #?startDateTime=<STARTDATE>&endDateTime=<ENDDATE>"

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



function RetrieveCalendarView($start, $end)
{
    # Retrieve the calendar view

    # Build the query string
    $viewStart = "$($start.ToUniversalTime().ToString("s"))Z"
    $viewEnd = "$($end.ToUniversalTime().ToString("s"))Z"
    $calenderViewRequest = "$($graphUrl)?startDateTime=$viewStart&endDateTime=$viewEnd"
    if (-not [String]::IsNullOrEmpty($viewQuery))
    {
        $calenderViewRequest = "$calenderViewRequest&$viewQuery"
    }

    # Send the request and return any items found
    $calendarItems = $null
    try
    {
        Write-Verbose "Requesting calendar: $calenderViewRequest"
        if ($psversiontable.PSVersion.Major -gt 6)
        {
            $thisCalendarView = Invoke-WebRequest -Method Get -Uri $calenderViewRequest -Headers $script:headers -SkipHeaderValidation
        }
        else
        {
            $thisCalendarView = Invoke-WebRequest -Method Get -Uri $calenderViewRequest -Headers $script:headers
        }
        $calendarItems = $thisCalendarView.Content | ConvertFrom-Json
        Write-Host "$($calendarItems.value.Count) item(s) returned: $calenderViewRequest" -ForegroundColor Gray
        LogToFile "$calenderViewRequest SUCCESS ($($calendarItems.value.Count) item(s) )"
        $global:successQueries += $calenderViewRequest
        return $calendarItems.value
    }
    catch
    {
        $global:errorQueries += $calenderViewRequest # Log the calls that generate an error
        Write-Host "ERROR: $calenderViewRequest" -ForegroundColor Red
        LogToFile "ERROR: $calenderViewRequest"
    }

    # We get here if an error occurred.  If we have an additional query, retry the request without the query
    if ([String]::IsNullOrEmpty($viewQuery))
    {
        return $null
    }

    $calenderViewRequest = "$($graphUrl)?startDateTime=$viewStart&endDateTime=$viewEnd"
    try
    {
        Write-Verbose "Requesting calendar: $calenderViewRequest"
        if ($psversiontable.PSVersion.Major -gt 6)
        {
            $thisCalendarView = Invoke-WebRequest -Method Get -Uri $calenderViewRequest -Headers $script:headers -SkipHeaderValidation
        }
        else
        {
            $thisCalendarView = Invoke-WebRequest -Method Get -Uri $calenderViewRequest -Headers $script:headers
        }
        $calendarItems = $thisCalendarView.Content | ConvertFrom-Json
        Write-Host "$($calendarItems.value.Count) item(s) returned: $calenderViewRequest" -ForegroundColor Gray
        LogToFile "$calenderViewRequest SUCCESS ($($calendarItems.value.Count) item(s) )"
        $global:successQueries += $calenderViewRequest
        return $calendarItems.value
    }
    catch
    {
        $global:errorQueries += $calenderViewRequest # Log the calls that generate an error
        Write-Host "ERROR: $calenderViewRequest" -ForegroundColor Red
        LogToFile "ERROR: $calenderViewRequest"
    }
    return $null
}



# Test calendar for corrupt items/queries that result in errors
# We query the whole day in half hour intervals.  This script ideally wants to be run with Fiddler also running to capture the full responses.

# https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0&tabs=http#query-parameters
# The values of startDateTime and endDateTime are interpreted using the timezone offset specified in the value and are not impacted by the value of the Prefer: outlook.timezone header if present.
# If no timezone offset is included in the value, it is interpreted as UTC.

if ($EndDate.Subtract($StartDate).TotalHours -gt $AnalysisPeriod)
{
}

$global:successQueries = @()
$global:errorQueries = @()


# Prepare the request headers
$script:headers = @{
    'Authorization'  = "$($oauth.token_type) $($oauth.access_token)";
}

# Retrieve calendar view for whole period
$calendarItems = RetrieveCalendarView $StartDate $EndDate

# Now retrieve calendar view in half hour segments
$periodStartDate = $StartDate
$periodEndDate = $StartDate.AddMinutes(30)

do {
    $calendarItems = RetrieveCalendarView $periodStartDate $periodEndDate
    $periodStartDate = $periodStartDate.AddMinutes(30)
    $periodEndDate = $periodEndDate.AddMinutes(30)
} while ($periodEndDate -le $EndDate)

Write-Host "Finished." -ForegroundColor Green
LogToFile "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) finished."