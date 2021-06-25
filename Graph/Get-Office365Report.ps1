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

    [Parameter(Mandatory=$False,HelpMessage="Name of the report to retrieve")]
    [ValidateNotNullOrEmpty()]
    [string]$ReportName = "getEmailActivityCounts",

    [Parameter(Mandatory=$False,HelpMessage="Name of the report to retrieve")]
    [ValidateNotNullOrEmpty()]
    [string]$Period = "D7",
    
    [Parameter(Mandatory=$False,HelpMessage="Path to save the report")]
    [ValidateNotNullOrEmpty()]
    [string]$SavePath = ""
)


# Get token
$Body = @{    
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default" 
    client_Id     = $AppId
    Client_Secret = $AppSecretKey
} 
$authResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantDomain/oauth2/v2.0/token" -Method POST -Body $Body -ErrorAction STOP

# Get report
$Headers = @{
    'Authorization' = "Bearer $($authResponse.access_token)"
}
$report = Invoke-RestMethod -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/reports/$($ReportName)(period='$Period')" -Method Get

# Output report
if (![String]::IsNullOrEmpty($SavePath))
{
    $currentDate = get-date -Format d
    $currentDate = $currentDate.Replace('/','-')
    $report | out-file -FilePath "$SavePath\$ReportName$currentDate.csv"
}
else
{
    $report
}