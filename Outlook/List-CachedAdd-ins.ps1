<#
.SYNOPSIS
Lists Office add-in manifests cached for the current Windows user.

.DESCRIPTION
Recursively searches the following cache folders, including extensionless files:
  %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
  %LOCALAPPDATA%\Microsoft\Outlook\HubAppFileCache

Reads standalone XML OfficeApp manifests and Outlook MOS/TAOS JSON cache entries,
including XML manifests embedded in their xmlDefinition fields. Handles Outlook's
UTF-16 cache files without a byte order mark and unwraps both value and
taosCacheValue cache envelopes.

Displays the containing cache file's full path, display name, Id and version.
For embedded XML, values come from the XML manifest. For JSON add-ins, values
come from titleDefinition and manifestId (falling back to appId, then titleId).
Only JSON entries containing Office add-in extension definitions are listed;
unrelated Outlook apps and cache bookkeeping files are skipped.
Use -Verbose for scan totals, file encodings, JSON property names and skip reasons.
Diagnostics do not dump raw cache contents. Wef is checked only for XML;
JSON parsing is restricted to HubAppFileCache.
The Wef XML scan skips extensions listed in $xmlIgnoredExtensions below.
This list does not exclude files from Outlook's JSON cache scan.
The Wef scan prunes folders listed in $wefIgnoredFolders before enumeration;
by default, this excludes the root webview2 browser-storage subtree.
Empty files are skipped using file metadata. Other files are inspected using a
buffered text reader; only XML/JSON candidates are fully parsed.
Use -OutlookOnly to restrict Wef discovery to folders named Manifests and their
descendants. The default retains a broader XML scan for other Office hosts.
HubAppFileCache is always scanned recursively, regardless of folder names.

.PARAMETER WriteToFile
Optional CSV file path. Results are still displayed in the console.

.PARAMETER OutlookOnly
Restrict the Wef XML scan to Manifests folders and their descendants.
Does not restrict the Outlook HubAppFileCache scan.

.EXAMPLE
.\List-CachedAdd-ins.ps1

.EXAMPLE
.\List-CachedAdd-ins.ps1 -WriteToFile .\CachedAdd-ins.csv

.EXAMPLE
.\List-CachedAdd-ins.ps1 -OutlookOnly

.EXAMPLE
.\List-CachedAdd-ins.ps1 -Verbose 4>&1 | Tee-Object -FilePath .\List-CachedAdd-ins-diagnostics.log
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$False, HelpMessage="Optional CSV file path to export results to, in addition to the console")]
    [ValidateNotNullOrEmpty()]
    [Alias("ExportPath")]
    [string]$WriteToFile,

    [Parameter(Mandatory=$False, HelpMessage="Restrict Wef XML discovery to Manifests folders while retaining the complete Outlook cache scan")]
    [switch]$OutlookOnly
)

if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Write-Error "The LOCALAPPDATA environment variable is not set."
    return
}

$readerSettings = New-Object System.Xml.XmlReaderSettings
$readerSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
$readerSettings.XmlResolver = $null

$xmlIgnoredExtensions = @(
    ".baf"
    ".baj"
    ".db"
    ".db-journal"
    ".db-wal"
    ".dll"
    ".hyb"
    ".journal"
    ".js"
    ".json"
    ".ldb"
    ".log"
    ".pb"
    ".pma"
    ".sig"
)

# These names apply only to immediate children of Wef, not other cache trees.
$wefIgnoredFolders = @(
    "webview2"
)

Write-Verbose ("Cache scan started {0:o}; Script='{1}'; PowerShell={2}; 64-bit process={3}" -f [DateTime]::UtcNow, $PSCommandPath, $PSVersionTable.PSVersion, [Environment]::Is64BitProcess)

function Get-CacheFiles {
    param(
        [Parameter(Mandatory=$True, HelpMessage="Full path of the cache root to enumerate")]
        [ValidateNotNullOrEmpty()]
        [string]$CachePath,

        [Parameter(Mandatory=$True, HelpMessage="Whether this is the Wef cache, where folder exclusions and Outlook-only filtering apply")]
        [bool]$IsWefCache,

        [Parameter(Mandatory=$True, HelpMessage="Scan counters to update when directories or files are skipped")]
        [hashtable]$Statistics
    )

    $directories = [System.Collections.Generic.Stack[string]]::new()
    $directories.Push($CachePath)
    while ($directories.Count -gt 0) {
        $directory = $directories.Pop()
        $enumerationErrors = @()
        Get-ChildItem -LiteralPath $directory -Force -ErrorAction Continue -ErrorVariable enumerationErrors | ForEach-Object {
            if ($_.PSIsContainer) {
                if ($IsWefCache -and $directory -eq $CachePath -and $wefIgnoredFolders -contains $_.Name) {
                    $Statistics.IgnoredDirectories++
                    Write-Verbose "Skipping cache directory '$($_.FullName)': excluded from Wef discovery."
                }
                elseif (($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                    Write-Verbose "Skipping cache directory '$($_.FullName)': directory reparse points are not followed."
                }
                else {
                    $directories.Push($_.FullName)
                }
            }
            elseif ($IsWefCache -and $OutlookOnly -and $_.DirectoryName.Substring($CachePath.Length) -notmatch '\\Manifests(?:\\|$)') {
                $Statistics.OutsideManifestFolders++
            }
            else {
                $_
            }
        }
        $Statistics.EnumerationErrors += $enumerationErrors.Count
    }
}

function Get-CacheObjectDescription {
    param(
        [Parameter(Mandatory=$True, HelpMessage="Parsed cache object whose structure should be described without exposing values")]
        [AllowNull()]
        [object]$Data
    )

    if ($null -eq $Data) {
        return "null"
    }
    if ($Data -is [array]) {
        return "array ($($Data.Count) items)"
    }
    if ($Data -is [PSCustomObject]) {
        $names = @($Data.PSObject.Properties.Name | Select-Object -First 16)
        return "object properties=[$($names -join ', ')]"
    }
    return $Data.GetType().FullName
}

function Get-FileCachedAddIn {
    param(
        [Parameter(Mandatory=$True, HelpMessage="Full path of the cache file to read")]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter(Mandatory=$True, HelpMessage="Whether the file is eligible for XML discovery")]
        [bool]$CheckXml,

        [Parameter(Mandatory=$True, HelpMessage="Whether the cache supports JSON discovery")]
        [bool]$CheckJson,

        [Parameter(Mandatory=$True, HelpMessage="Scan counters to update after inspecting the file prefix")]
        [hashtable]$Statistics
    )

    $stream = [System.IO.File]::OpenRead($FilePath)
    $textReader = $null
    try {
        $firstByte = $stream.ReadByte()
        $secondByte = $stream.ReadByte()
        $stream.Position = 0
        $encoding = [System.Text.Encoding]::UTF8

        # HubAppFileCache writes UTF-16LE without a BOM; retain BOM detection for other files.
        if ($firstByte -gt 0 -and $secondByte -eq 0) {
            $encoding = [System.Text.Encoding]::Unicode
        }
        elseif ($firstByte -eq 0 -and $secondByte -gt 0) {
            $encoding = [System.Text.Encoding]::BigEndianUnicode
        }

        $textReader = [System.IO.StreamReader]::new($stream, $encoding, $true)
        while (($nextCharacter = $textReader.Peek()) -ge 0 -and [char]::IsWhiteSpace([char]$nextCharacter)) {
            $null = $textReader.Read()
        }
        Write-Verbose "Inspecting cache file '$FilePath': bytes=$($stream.Length); encoding=$($textReader.CurrentEncoding.WebName)."
        if ($nextCharacter -lt 0) {
            $Statistics.Empty++
            Write-Verbose "Skipping empty cache file '$FilePath'."
        }
        elseif ($CheckXml -and $nextCharacter -eq [int][char]"<") {
            $Statistics.XmlCandidates++
            # Preserve BOM-less UTF-16 detection; otherwise let XML declarations select the encoding.
            if ($textReader.CurrentEncoding.CodePage -in @(1200, 1201)) {
                Get-XmlCachedAddIn -TextReader $textReader -FilePath $FilePath
            }
            else {
                $stream.Position = 0
                Get-XmlCachedAddIn -Stream $stream -FilePath $FilePath
            }
        }
        elseif ($CheckJson -and ($nextCharacter -eq [int][char]"{" -or $nextCharacter -eq [int][char]"[")) {
            $Statistics.JsonCandidates++
            $data = ConvertFrom-Json -InputObject $textReader.ReadToEnd() -ErrorAction Stop
            if ($null -eq $data) {
                Write-Verbose "Skipping empty JSON collection in '$FilePath'."
            }
            else {
                Get-JsonCachedAddIn -Data $data -FilePath $FilePath
            }
        }
        else {
            $Statistics.Other++
            Write-Verbose "Skipping '$FilePath': unrecognized content for this cache's manifest scan; inspected bytes=$($stream.Position) of $($stream.Length)."
        }
    }
    finally {
        if ($null -ne $textReader) {
            $textReader.Dispose()
        }
        $stream.Dispose()
    }
}

function Get-XmlCachedAddIn {
    param(
        [Parameter(Mandatory=$False, HelpMessage="Embedded XML OfficeApp manifest")]
        [ValidateNotNullOrEmpty()]
        [string]$Xml,

        [Parameter(Mandatory=$False, HelpMessage="Already-open text reader for a standalone XML manifest")]
        [ValidateNotNull()]
        [System.IO.TextReader]$TextReader,

        [Parameter(Mandatory=$False, HelpMessage="Already-open stream for XML with its own encoding declaration")]
        [ValidateNotNull()]
        [System.IO.Stream]$Stream,

        [Parameter(Mandatory=$True, HelpMessage="Full path of the file containing the manifest")]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath
    )

    $stringReader = $null
    $reader = $null
    try {
        if ($PSBoundParameters.ContainsKey("Xml")) {
            $stringReader = [System.IO.StringReader]::new($Xml)
            $reader = [System.Xml.XmlReader]::Create($stringReader, $readerSettings)
        }
        elseif ($PSBoundParameters.ContainsKey("TextReader")) {
            $reader = [System.Xml.XmlReader]::Create($TextReader, $readerSettings)
        }
        elseif ($PSBoundParameters.ContainsKey("Stream")) {
            $reader = [System.Xml.XmlReader]::Create($Stream, $readerSettings)
        }
        else {
            $reader = [System.Xml.XmlReader]::Create($FilePath, $readerSettings)
        }
        if ($reader.MoveToContent() -ne [System.Xml.XmlNodeType]::Element -or $reader.LocalName -ne "OfficeApp") {
            Write-Verbose "Skipping XML in '$FilePath': root element '$($reader.LocalName)' is not OfficeApp."
            return
        }

        $manifest = New-Object System.Xml.XmlDocument
        $manifest.XmlResolver = $null
        $manifest.Load($reader)
        $root = $manifest.DocumentElement
        $namespaces = New-Object System.Xml.XmlNamespaceManager($manifest.NameTable)
        $namespaces.AddNamespace("manifest", $root.NamespaceURI)
        $id = $root.SelectSingleNode("manifest:Id", $namespaces)
        $displayName = $root.SelectSingleNode("manifest:DisplayName", $namespaces)
        $version = $root.SelectSingleNode("manifest:Version", $namespaces)

        if ($null -eq $id -or $null -eq $displayName -or $null -eq $version) {
            Write-Warning "Skipping incomplete Office add-in manifest: $FilePath"
            return
        }

        Write-Verbose "Matched XML OfficeApp manifest in '$FilePath'."
        [PSCustomObject]@{
            FilePath = $FilePath
            DisplayName = $displayName.GetAttribute("DefaultValue")
            Id = $id.InnerText.Trim()
            Version = $version.InnerText.Trim()
        }
    }
    catch [System.Xml.XmlException] {
        Write-Verbose "Skipping invalid XML in '$FilePath': $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        if ($null -ne $stringReader) {
            $stringReader.Dispose()
        }
    }
}

function Get-JsonCachedAddIn {
    param(
        [Parameter(Mandatory=$True, HelpMessage="Parsed Outlook MOS/TAOS JSON cache entry")]
        [ValidateNotNullOrEmpty()]
        [object]$Data,

        [Parameter(Mandatory=$True, HelpMessage="Full path of the file containing the cache entry")]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath
    )

    Write-Verbose "JSON root in '$FilePath': $(Get-CacheObjectDescription -Data $Data)."
    if ($null -ne $Data.PSObject.Properties["value"]) {
        $Data = $Data.value
        Write-Verbose "JSON payload in '$FilePath': $(Get-CacheObjectDescription -Data $Data)."
    }
    if ($null -ne $Data -and $null -ne $Data.PSObject.Properties["taosCacheValue"]) {
        $Data = $Data.taosCacheValue
        Write-Verbose "TAOS payload in '$FilePath': $(Get-CacheObjectDescription -Data $Data)."
    }
    if ($null -eq $Data -or $Data -isnot [PSCustomObject] -or $null -eq $Data.titleDefinition) {
        Write-Verbose "Skipping JSON in '$FilePath': no titleDefinition acquisition object."
        return
    }

    $definition = $Data.titleDefinition
    $elements = $definition.elementDefinitions
    Write-Verbose "JSON elementDefinitions in '$FilePath': $(Get-CacheObjectDescription -Data $elements)."
    $xmlAddIns = @($elements.exchangeAddIns) + @($elements.officeAddIns)
    $hasXml = $false
    foreach ($addIn in $xmlAddIns) {
        if (-not [string]::IsNullOrWhiteSpace($addIn.xmlDefinition)) {
            $hasXml = $true
            Write-Verbose "Parsing embedded XML definition in '$FilePath'."
            Get-XmlCachedAddIn -Xml $addIn.xmlDefinition -FilePath $FilePath
        }
    }

    if ($hasXml) {
        return
    }
    if (-not $elements.extensions) {
        Write-Verbose "Skipping JSON in '$FilePath': no embedded XML or extensions array."
        return
    }

    $id = $Data.manifestId
    if ([string]::IsNullOrWhiteSpace($id)) {
        $id = $Data.appId
    }
    if ([string]::IsNullOrWhiteSpace($id)) {
        $id = $Data.titleId
    }

    if ([string]::IsNullOrWhiteSpace($id) -or [string]::IsNullOrWhiteSpace($definition.name) -or [string]::IsNullOrWhiteSpace($definition.version)) {
        Write-Warning "Skipping incomplete Office add-in JSON entry: $FilePath"
        return
    }

    Write-Verbose "Matched JSON Office add-in in '$FilePath'."
    [PSCustomObject]@{
        FilePath = $FilePath
        DisplayName = $definition.name
        Id = $id
        Version = $definition.version
    }
}

$wefCachePath = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\Office\16.0\Wef"
$cachePaths = @(
    $wefCachePath
    (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\Outlook\HubAppFileCache")
)

$results = @(
    foreach ($cachePath in $cachePaths) {
        $isWefCache = $cachePath -eq $wefCachePath
        if ($isWefCache) {
            Write-Verbose "Scanning XML cache '$cachePath'."
            if ($OutlookOnly) {
                Write-Verbose "Outlook-only Wef discovery: checking files only within Manifests folders."
            }
        }
        else {
            Write-Verbose "Scanning Outlook cache '$cachePath'."
        }
        if (-not (Test-Path -LiteralPath $cachePath -PathType Container)) {
            Write-Verbose "Office add-in cache folder not found: $cachePath"
            continue
        }

        $scanStatistics = @{
            Files = 0
            Ignored = 0
            Empty = 0
            XmlCandidates = 0
            JsonCandidates = 0
            Other = 0
            ReadOrJsonFailures = 0
            AddIns = 0
            IgnoredDirectories = 0
            OutsideManifestFolders = 0
            EnumerationErrors = 0
        }
        Get-CacheFiles -CachePath $cachePath -IsWefCache $isWefCache -Statistics $scanStatistics | ForEach-Object {
            $file = $_
            $scanStatistics.Files++
            $checkXml = $xmlIgnoredExtensions -notcontains $file.Extension
            if (-not $checkXml -and $isWefCache) {
                $scanStatistics.Ignored++
                Write-Verbose "Skipping '$($file.FullName)': extension '$($file.Extension)' is excluded from XML discovery."
                return
            }

            if ($file.Length -eq 0) {
                $scanStatistics.Empty++
                Write-Verbose "Skipping empty cache file '$($file.FullName)' without opening it."
                return
            }

            try {
                $fileResults = @(Get-FileCachedAddIn -FilePath $file.FullName -CheckXml $checkXml -CheckJson (-not $isWefCache) -Statistics $scanStatistics)
                $scanStatistics.AddIns += $fileResults.Count
                $fileResults
            }
            catch {
                $scanStatistics.ReadOrJsonFailures++
                Write-Warning "Unable to parse cache file '$($file.FullName)': $($_.Exception.Message)"
            }
        }

        Write-Verbose ("Cache scan summary '{0}': files={1}; ignored extensions={2}; empty={3}; XML candidates={4}; JSON candidates={5}; other={6}; read/JSON failures={7}; enumeration errors={8}; add-ins={9}; ignored directories={10}; files outside Manifests folders={11}." -f $cachePath, $scanStatistics.Files, $scanStatistics.Ignored, $scanStatistics.Empty, $scanStatistics.XmlCandidates, $scanStatistics.JsonCandidates, $scanStatistics.Other, $scanStatistics.ReadOrJsonFailures, $scanStatistics.EnumerationErrors, $scanStatistics.AddIns, $scanStatistics.IgnoredDirectories, $scanStatistics.OutsideManifestFolders)
    }
)

if ($WriteToFile) {
    $results | Export-Csv -LiteralPath $WriteToFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
}

if ($results.Count -eq 0) {
    Write-Warning "No cached Office add-in manifests found in: $($cachePaths -join ', ')"
}
else {
    $results | Format-List -Property FilePath, DisplayName, Id, Version
}
