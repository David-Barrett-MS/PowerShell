param (
	[Parameter(Mandatory=$True,HelpMessage="Path from which to read event logs")]
	[ValidateNotNullOrEmpty()]
	[string]$SourceFolder
)

$global:events = @()

dir "$SourceFolder\*.json" | foreach {
    $fileContent = Get-Content $_
    $fileEvents = ConvertFrom-Json $fileContent
    foreach ($fileEvent in $fileEvents) {
        Add-Member -InputObject $fileEvent -Type NoteProperty -PassThru -Name 'SourceAuditFile' -Value "$($_.Name)" | out-null
        $global:events += $fileEvent
    }
}
if ($global:events.Count -lt 1)
{
    Write-Host "Failed to read any events" -ForegroundColor Red
    exit
}

Write-Host "$($global:events.Count) events loaded" -ForegroundColor Green

Write-Host "Example event query:"
Write-Host "`$events | where-object -FilterScript { `$_.recordtype -eq 84 -or `$_.recordtype -eq 83 } | ft" -ForegroundColor Yellow
Write-Host
Write-Host "Example export of events:"
Write-Host "`$events | where-object -FilterScript { `$_.recordtype -eq 84 -or `$_.recordtype -eq 83 } | export-csv `"o365api.csv`" -NoTypeInformation" -ForegroundColor Yellow

<#
$events | where-object -FilterScript { $_.recordtype -eq 84 -or $_.recordtype -eq 83 -or $_.recordtype -eq 43} | ft

CreationTime        Id                                   Operation                    OrganizationId
------------        --                                   ---------                    --------------
2022-11-17T12:41:41 2c0e208e-0284-41d7-b3ac-218ea8b07bf1 MipLabel                     fc69f6a8-90cd-4047-977d-0c768925b8ec
2022-11-17T11:08:53 48796a12-c5b1-4ad6-b9b1-0eb9d919f212 SensitivityLabelApplied      fc69f6a8-90cd-4047-977d-0c768925b8ec
2022-11-17T11:08:14 8ebb2619-956c-446e-a931-acdd4c7fa226 SensitivityLabelApplied      fc69f6a8-90cd-4047-977d-0c768925b8ec
2022-11-17T12:38:04 bb682ce7-573f-4aba-b420-7996aeae7094 SensitivityLabelApplied      fc69f6a8-90cd-4047-977d-0c768925b8ec
2022-11-17T12:40:45 a474e8ba-faab-4e94-b8b9-ea6595ebb92e SensitivityLabelRemoved      fc69f6a8-90cd-4047-977d-0c768925b8ec
2022-11-17T12:40:26 120d8aa3-3133-4f4f-96b6-7bb14ee8c74b SensitivityLabeledFileOpened fc69f6a8-90cd-4047-977d-0c768925b8ec
2022-11-17T12:41:36 fd65bc3e-a8b9-4548-9869-5d22dcb7d5c4 SensitivityLabelApplied      fc69f6a8-90cd-4047-977d-0c768925b8ec

$events | where-object -FilterScript { $_.recordtype -eq 84 -or $_.recordtype -eq 83 -or $_.recordtype -eq 43} | export-csv "o365api.csv" -NoTypeInformation



$powerevents = $events | where-object -FilterScript { $_.Workload.StartsWith("PowerPlatform") }
$purviewevents = Import-CSV  -Path "E:\Scripts\demonmaths.co.uk\Analysis\From Purview. (PowerPlatform workload)csv.csv"

# Check each API event is present in Purview export
$missingEvents = @()
foreach ($powerevent in $powerevents) {
	$foundPurviewEvent = $false
	foreach ($purviewevent in $purviewevents) {
		if ($powerevent.Id -eq $purviewevent.RecordId) {
			Write-Host "Found $($purviewevent.RecordId)" -ForegroundColor Green
			$foundPurviewEvent = $true
			break;
		}
	}
	if (!$foundPurviewEvent) {
		$missingEvents += $powerevent
		Write-Host "Missing from Purview: $($powerevent.Id)" -ForegroundColor Red
	}
}

# Check each Purview event is present in API feed
$missingEvents = @()
foreach ($purviewevent in $purviewevents) {	
	$foundAPIEvent = $false
	foreach ($powerevent in $powerevents) {
		if ($powerevent.Id -eq $purviewevent.RecordId) {
			Write-Host "Found $($powerevent.Id)" -ForegroundColor Green
			$foundAPIEvent = $true
			break;
		}
	}
	if (!$foundAPIEvent) {
		$missingEvents += $purviewevent
		Write-Host "Missing from API: $($purviewevent.RecordId)" -ForegroundColor Red
	}
}



#>