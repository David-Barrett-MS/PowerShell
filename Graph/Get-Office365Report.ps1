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



$Body = @{    
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default" 
    client_Id     = $AppId
    Client_Secret = $AppSecretKey
} 

$authResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantDomain/oauth2/v2.0/token" -Method POST -Body $Body -ErrorAction STOP

$Headers = @{
    'Authorization' = "Bearer $($authResponse.access_token)"
}

$currentDate = get-date -Format d
$currentDate = $currentDate.Replace('/','-')

$apiUrl = "https://graph.microsoft.com/v1.0/reports/$($ReportName)(period='$Period')"
$report = Invoke-RestMethod -Headers $Headers -Uri $apiUrl -Method Get

if (![String]::IsNullOrEmpty($SavePath))
{
    $report | out-file -FilePath "$SavePath\$ReportName$currentDate.csv"
}
else
{
    $report
}