#
# Import-AppointmentsFromCSV.ps1
#
# By David Barrett, Microsoft Ltd. Use at your own risk.  No warranties are given.
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
Imports appointments from a CSV file into a mailbox calendar using Microsoft Graph.

.DESCRIPTION
This script demonstrates how to import appointments from a CSV file into a mailbox calendar using Microsoft Graph.

Permissions required:
	Calendars.ReadWrite

This script can use delegated or application permissions.  If using delegated permissions, the script will prompt for user authentication via a web browser.

.EXAMPLE
Generate sample appointments and then import into a mailbox calendar using application permissions:
.\Import-AppointmentsFromCSV.ps1 -GenerateSampleCSV -SourceFile "c:\Temp\appointments.csv" -Mailbox "mailbox" -OAuthClientId "AppId" -OAuthTenantId "TenantId" -AppSecretKey "SecretKey"

Import appointments from CSV into a mailbox calendar using application permissions:
.\Import-AppointmentsFromCSV.ps1 -SourceFile "c:\Temp\appointments.csv" -Mailbox "mailbox" -OAuthClientId "AppId" -OAuthTenantId "TenantId" -AppSecretKey "SecretKey"

Import appointments from CSV into a mailbox calendar using delegated permissions (assuming app registration has redirect URL of http://localhost/code):
.\Import-AppointmentsFromCSV.ps1 -SourceFile "c:\Temp\appointments.csv" -OAuthClientId "AppId" -OAuthTenantId "TenantId"

#>
param (
    [Parameter(Position=1,Mandatory=$True,HelpMessage="CSV file containing the appointments to import.")]
    [ValidateNotNullOrEmpty()]
    [string]$SourceFile,

    [Parameter(Position,Mandatory=$False,HelpMessage="Timezone in which appointments will be created (default is UTC).  This timezone is used if no timezone is specified in the CSV file.")]
    [ValidateNotNullOrEmpty()]
    [string]$Timezone = "UTC",

    [Parameter(Mandatory=$False,HelpMessage="If specified, a sample CSV will be created (and then processed).")]
    [ValidateNotNullOrEmpty()]
    [switch]$GenerateSampleCSV,

<# Logging.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="If included, activity is logged to a file (same as script name with .log appended).")]	
    [switch]$LogToFile,
<# Logging.ps1 %PARAMS_END% #>

<# Auth.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="Application Id (obtained when registering the application in Azure AD).")]
    [ValidateNotNullOrEmpty()]
    [string]$OAuthClientId,

    [Parameter(Mandatory=$False,HelpMessage="Application secret key (obtained when registering the application in Azure AD).  If not specified, delegate permissions are assumed.")]
    [ValidateNotNullOrEmpty()]
    [string]$AppSecretKey,

    [Parameter(Mandatory=$False,HelpMessage="Redirect Url (specified when registering the application in Azure AD; use localhost).  Required for delegate permissions.  Default is http://localhost/code")]
    [ValidateNotNullOrEmpty()]
    [string]$RedirectUrl = "http://localhost/code",

    [Parameter(Mandatory=$False,HelpMessage="Tenant Id.  Default is common, which requires the application to be registered for multi-tenant use (and consented to in target tenant).")]
    [ValidateNotNullOrEmpty()]
    [string]$OAuthTenantId = "common",

    [Parameter(Mandatory=$False,HelpMessage="Mailbox.  Default is me to access own mailbox when using delegate permissions; must be set to user id for application permissions.")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox = "me"
)
<# Auth.ps1 %PARAMS_END% #>

$script:ScriptVersion = "1.0.0"
if ($TraceGraphCalls) {
    $script:traceFile = "$($MyInvocation.InvocationName).trace"
}

<# Logging.ps1 %FUNCTIONS_START% #>
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
<# Logging.ps1 %FUNCTIONS_END% #>

<# Auth.ps1 %FUNCTIONS_START% #>
function GetAppAuthToken()
{
    # Acquire token for application permissions
    $body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$OAuthClientId;client_secret=$AppSecretKey}
    try
    {
        $script:oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:token_expires = (Get-Date).AddSeconds($oauth.expires_in)
    }
    catch
    {
        Log "Failed to obtain OAuth token" Red
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
    $body = @{grant_type="refresh_token";scope="https://graph.microsoft.com/.default";client_id=$OAuthClientId;refresh_token=$oauth.refresh_token}
    try
    {
        $script:oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
        $script:token_expires = (Get-Date).AddSeconds($oauth.expires_in)
    }
    catch
    {
        Log "Failed to renew OAuth token" Red
        exit # Failed to renew the token
    }
    Log "Successfully renewed OAuth token" Green
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
    Log "Calling URL: $url"
    try
    {
        $result = Invoke-RestMethod -Method Get -Uri $url -Headers $script:headers
    }
    catch
    {
        Log "Failed to obtain data from URL: $url" Red
        return $null
    }
    return $result
}

function POST($url, $body)
{
    UpdateHeaders
    Log "Calling URL: $url"
    try
    {
        $result = Invoke-RestMethod -Method Post -Uri $url -Headers $script:headers -Body $body
    }
    catch
    {
        Log "Error during POST to URL: $url" Red
        return $null
    }
    return $result
}

function PATCH($url, $body)
{
    UpdateHeaders
    Log "Calling URL: $url"
    try
    {
        $result = Invoke-RestMethod -Method Patch -Uri $url -Headers $script:headers -Body $body
    }
    catch
    {
        Log "Failed to obtain data from URL: $url" Red
        return $null
    }
    return $result
}

function DELETE($url)
{
    UpdateHeaders
    Log "Calling URL: $url"
    try
    {
        $result = Invoke-RestMethod -Method Delete -Uri $url -Headers $script:headers
    }
    catch
    {
        Log "Failed to obtain data from URL: $url" Red
        return $null
    }
    return $result
}

## Authentication

if ([String]::IsNullOrEmpty($AppSecretKey))
{
    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/authorize?client_id=$OAuthClientId&response_type=code&redirect_uri=$RedirectUrl&response_mode=query&scope=openid%20profile%20email%20offline_access%20"
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
    $body = @{grant_type="authorization_code";scope="https://graph.microsoft.com/.default";client_id=$OAuthClientId;code=$authcode;redirect_uri=$RedirectUrl}
    try
    {
        $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$OAuthTenantId/oauth2/v2.0/token -Body $body
    }
    catch
    {
        Log "Failed to obtain OAuth token" Red
        exit # Failed to obtain a token
    }
    $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/me/"
}
else
{
    GetAppAuthToken
}
Log "Successfully obtained OAuth token" Green
<# Auth.ps1 %FUNCTIONS_END% #>



# Main Script

# Check if we need to generate a sample CSV first of all
if ($GenerateSampleCSV)
{
    # CSV Headers
    $csvData = new-object Text.StringBuilder
    [void]$csvData.AppendLine("Subject,Start Date,Start Time,End Date,End Time")

    # Now generate some sample appointments
    $i = 1
    while ($i -lt 10)
    {
        # Generate the start and end time
        $startDate = [DateTime]::Today.AddDays(7*$i).AddHours(10)
        $endDate = $startDate.AddHours(1)

        # Add the appointment to the CSV data
        [void]$csvData.Append("CSV Import Test $i,")
        [void]$csvData.Append($startDate.ToString("yyyy-MM-dd"))
        [void]$csvData.Append(",")
        [void]$csvData.Append($startDate.ToString("HH:mm:ss"))
        [void]$csvData.Append(",")
        [void]$csvData.Append($endDate.ToString("yyyy-MM-dd"))
        [void]$csvData.Append(",")
        [void]$csvData.AppendLine($endDate.ToString("HH:mm:ss"))
        $i++
    }

    # And export the data to CSV
    try
    {
        $csvData.ToString() | Out-File $SourceFile
        Log "Successfully exported sample CSV data to $SourceFile" Green
    }
    catch
    {
        Log "Failed to export sample CSV data to $SourceFile" Red
        exit
    }
}

# CSV File Check
if (!(Get-Item -Path $SourceFile -ErrorAction SilentlyContinue))
{
	Log "Unable to open file: $SourceFile" Red
    exit
}
 
# Import CSV File
try
{
	$CSVFile = Import-Csv -Path $SourceFile
}
catch { }
if (!$CSVFile)
{
    $csvHeader = 'Subject','Start Date','Start Time','End Date','End Time'
	Log "CSV header line not found, trying to import using predefined header: $csvHeader"
	$CSVFile = Import-Csv -Path $SourceFile -header $csvHeader
}
if (!$CSVFile)
{
	Log "Unable to open file: $SourceFile" Red
    exit
}

# Check file has required fields

function GetFieldValue($Item, $FieldName)
{
    if ($Item.$FieldName)
    {
        return $Item.$FieldName
    }
    if ($Item.$($FieldMap.$FieldName))
    {
        return $Item.$($FieldMap.$FieldName)
    }
    return ""
}
$RequiredFields=@("Subject","StartDate","StartTime","EndDate","EndTime")

$FieldMap=@{
	"subject" = "Subject";
	"startDate" = "Start Date";
	"startTime" = "Start Time";
	"endDate" = "End Date";
	"endTime" = "End Time";
    "isAllDay" = "All day event";
    "location" = "Location";
    "timeZone" = "Time zone";
}

foreach ($requiredField in $RequiredFields)
{
	if ([String]::IsNullOrEmpty($(GetFieldValue $CSVFile[0] $requiredField)))
    {
		# Missing required field
		Log "Import file is missing required field: $requiredField" Red
        exit
    }
}

# Parse the CSV file and add the appointments

$eventsURL = "$($script:graphBaseUrl)calendar/events"

foreach ($CalendarItem in $CSVFile)
{    
	# Create the appointment and set the required fields (the minimum required to create an appointment))
    $appointmentTimezone = GetFieldValue $CalendarItem "timeZone"
    if ([String]::IsNullOrEmpty($appointmentTimezone))
    {
        $appointmentTimezone = $Timezone
    }
    $appointment = @{
        subject = GetFieldValue $CalendarItem "subject"
        body = @{
            contentType = "HTML"
            content = ""
        }
        start = @{
            dateTime = "$(GetFieldValue $CalendarItem "startDate")T$(GetFieldValue $CalendarItem "startTime")"
            timeZone = $appointmentTimezone
        }
        end = @{
            dateTime = "$(GetFieldValue $CalendarItem "endDate")T$(GetFieldValue $CalendarItem "endTime")"
            timeZone = $appointmentTimezone
        }
        location = @{
            displayName = ""
        }
    }

    if ([String]::IsNullOrEmpty($appointment.subject))
    {
        Log "Skipping appointment with no subject" Yellow
        continue
    }

    # Add optional fields if specified
    foreach ($optionalField in $FieldMap.Keys)
    {
        if ($RequiredFields -contains $optionalField) { continue } # Already processed
        $fieldValue = GetFieldValue $CalendarItem $optionalField
        if (![String]::IsNullOrEmpty($fieldValue))
        {
            $appointment.$optionalField = $fieldValue
        }
    }

    # Add the appointment to the calendar
    try
    {
        POST $eventsURL ($appointment | ConvertTo-Json -Depth 3) | Out-Null
        Log "Created $($CalendarItem."Subject")" green
    }
    catch
    {
        Log "Failed to create appointment (error on save): $($CalendarItem."Subject")" red
    }
}    
