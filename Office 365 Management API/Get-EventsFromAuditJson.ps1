param (
	[Parameter(Mandatory=$True,HelpMessage="Path from which to read event logs")]
	[ValidateNotNullOrEmpty()]
	[string]$SourceFolder,
	
	[Parameter(Mandatory=$False,HelpMessage="File mask to filter files that will be ingested (default is *.json)")]
	[ValidateNotNullOrEmpty()]
	[string]$FileMask = "*.json",	

	[Parameter(Mandatory=$False,HelpMessage="If specified, events from the API feed are deduplicated.")]
	[ValidateNotNullOrEmpty()]
	[switch]$Deduplicate,
	
	[Parameter(Mandatory=$False,HelpMessage="If specified, this file (export from Purview Portal) will be tested to check the API feed contains the same events")]
	[ValidateNotNullOrEmpty()]
	[string]$PurviewExportCompare,

	[Parameter(Mandatory=$False,HelpMessage="If specified, a comparison is done both ways between the Purview export and the API data (this isn't usually useful)")]
	[ValidateNotNullOrEmpty()]
	[switch]$CompareBothWays
)

$global:events = @()
$eventIds = @()
$duplicateCount = 0

dir "$SourceFolder\$FileMask" | foreach {
    $fileContent = Get-Content $_
    $fileEvents = ConvertFrom-Json $fileContent
    foreach ($fileEvent in $fileEvents) {
		$addEvent = $true
		if ($Deduplicate) {
			if ($eventIds.Contains($fileEvent.Id)) {
				Write-Verbose "Duplicate event ignored: $($fileEvent.Id)"
				$addEvent = $false
				$duplicateCount++
			} else {
				$eventIds += $fileEvent.Id
			}
		}
		if ($addEvent) {
			Add-Member -InputObject $fileEvent -Type NoteProperty -PassThru -Name 'SourceAuditFile' -Value "$($_.Name)" | out-null
			$global:events += $fileEvent
		}
    }
}
if ($Deduplicate) {
	Write-Host "$duplicateCount duplicate(s) ignored from API data"
}
if ($global:events.Count -lt 1)
{
    Write-Host "No events read" -ForegroundColor Red
    exit
}

if ([String]::IsNullOrEmpty($PurviewExportCompare)) {
	Write-Host "$($global:events.Count) events loaded" -ForegroundColor Green
	Write-Host "Example event query:"
	Write-Host "`$events | where-object -FilterScript { `$_.recordtype -eq 84 -or `$_.recordtype -eq 83 } | ft" -ForegroundColor Yellow
	Write-Host
	Write-Host "Example export of events:"
	Write-Host "`$events | where-object -FilterScript { `$_.recordtype -eq 84 -or `$_.recordtype -eq 83 } | export-csv `"o365api.csv`" -NoTypeInformation" -ForegroundColor Yellow
	exit
}

$purviewevents = Import-CSV  -Path $PurviewExportCompare
if ($purviewevents.Count -lt 1) {
	Write-Host "Failed to read any events from Purview export file: $PurviewExportCompare" -ForegroundColor Red
	exit
}
Write-Host "Purview import contains $($purviewevents.Count) event(s)"
Write-Host "API data contains $($global:events.Count) event(s)"

# Check each Purview event is present in API feed
$global:missingAPIEvents = @()
foreach ($purviewevent in $purviewevents) {	
	$foundAPIEvent = $false
	foreach ($apievent in $global:events) {
		if ($apievent.Id -eq $purviewevent.RecordId) {
			Write-Verbose "Found $($apievent.Id)"
			$foundAPIEvent = $true
			break;
		}
	}
	if (!$foundAPIEvent) {
		$global:missingAPIEvents += $purviewevent
		Write-Host "Missing from API: $($purviewevent.RecordId)" -ForegroundColor Red
	}
}

if ($CompareBothWays) {
	# Check each API event is present in Purview export
	$global:missingPurviewEvents = @()
	foreach ($apievent in $global:events) {
		$foundPurviewEvent = $false
		foreach ($purviewevent in $purviewevents) {
			if ($apievent.Id -eq $purviewevent.RecordId) {
				Write-Verbose "Found $($purviewevent.RecordId)"
				$foundPurviewEvent = $true
				break;
			}
		}
		if (!$foundPurviewEvent) {
			$global:missingPurviewEvents += $apievent
			Write-Host "Missing from Purview: $($apievent.Id)" -ForegroundColor Red
		}
	}
	
	if ($global:missingPurviewEvents.Count -gt 0) {
		Write-Host "$($global:missingPurviewEvents.Count) API event(s) missing from Purview export (stored in `$missingPurviewEvents)"
		Write-Host "It is usually expected that the imported API feed will contain many more audit events than a Purview export."
	} else {
		Write-Host "No events missing from Purview export" -ForegroundColor Green
	}
}

if ($global:missingAPIEvents.Count -gt 0) {
	Write-Host "$($global:missingAPIEvents.Count) Purview event(s) missing from API feed (stored in `$missingAPIEvents)" -ForegroundColor Red
} else {
	Write-Host "No events missing from API feed" -ForegroundColor Green
}

Write-Host "API events are available in `$events"
