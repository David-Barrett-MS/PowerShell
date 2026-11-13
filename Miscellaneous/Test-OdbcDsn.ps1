<#
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

.SYNOPSIS
Tests a DSN's connection and schema retrieval in 32-bit and 64-bit Windows PowerShell.
.EXAMPLE
.\Test-OdbcDsn.ps1 -Dsn 'MyUserDsn'
.EXAMPLE
.\Test-OdbcDsn.ps1 -Dsn 'MyUserDsn' -UseDsnCredentials -Collection Tables
#>
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()][string]$Dsn,
    [Management.Automation.PSCredential]$Credential,
    [switch]$UseDsnCredentials,
    [ValidateSet('Tables', 'Columns', 'Views', 'MetaDataCollections')]
    [string]$Collection = 'Tables',
    [ValidateRange(0, 100)][int]$MaxSchemaRows = 10,
    [ValidateRange(5, 600)][int]$TimeoutSeconds = 90,
    [switch]$Worker
)

$ErrorActionPreference = 'Stop'

function New-DsnConnectionString {
    param([string]$Name, [Management.Automation.PSCredential]$DatabaseCredential)
    Add-Type -AssemblyName System.Data
    $builder = New-Object System.Data.Odbc.OdbcConnectionStringBuilder
    try {
        $builder['DSN'] = $Name
        if ($null -ne $DatabaseCredential) {
            $builder['UID'] = $DatabaseCredential.UserName
            $builder['PWD'] = $DatabaseCredential.GetNetworkCredential().Password
        }
        return $builder.ConnectionString
    }
    finally { $builder.Clear() }
}

function Invoke-DsnSchemaProbe {
    param(
        [string]$Name, [string]$SchemaCollection, [int]$SampleLimit,
        [Management.Automation.PSCredential]$DatabaseCredential,
        [scriptblock]$ConnectionFactory = { New-Object System.Data.Odbc.OdbcConnection }
    )
    $result = [ordered]@{
        Bitness = [IntPtr]::Size * 8; Dsn = $Name; Collection = $SchemaCollection
        StartedUtc = [DateTime]::UtcNow.ToString('o'); FinishedUtc = $null
        Connection = 'FAIL'; Schema = 'NOT RUN'; RowCount = $null
        Stage = 'Open'; Driver = $null; Sample = @(); Errors = @()
    }
    $connection = $null
    $schema = $null
    try {
        $connectionString = New-DsnConnectionString $Name $DatabaseCredential
        $connection = & $ConnectionFactory
        $connection.ConnectionString = $connectionString
        $connectionString = $null
        $connection.Open()
        $result.Connection = 'PASS'
        $result.Driver = $connection.Driver
        $result.Stage = 'GetSchema'
        $result.Schema = 'FAIL'
        $schema = $connection.GetSchema($SchemaCollection)
        $result.RowCount = $schema.Rows.Count
        $rows = @()
        for ($i = 0; $i -lt [Math]::Min($SampleLimit, $schema.Rows.Count); $i++) {
            $record = [ordered]@{}
            foreach ($column in $schema.Columns) {
                $value = $schema.Rows[$i][$column.ColumnName]
                $record[$column.ColumnName] = if ($value -is [DBNull]) { $null } else { [string]$value }
            }
            $rows += [pscustomobject]$record
        }
        $result.Sample = $rows
        $result.Schema = 'PASS'
        $result.Stage = 'Complete'
    }
    catch {
        $failure = $_.Exception
        while (($failure -isnot [System.Data.Odbc.OdbcException]) -and $failure.InnerException) {
            $failure = $failure.InnerException
        }
        if ($failure -is [System.Data.Odbc.OdbcException]) {
            $result.Errors = @($failure.Errors | ForEach-Object {
                [pscustomobject]@{
                    SQLState = $_.SQLState; NativeError = $_.NativeError
                    Source = $_.Source; Message = $_.Message
                }
            })
        }
        else {
            $result.Errors = @([pscustomobject]@{
                SQLState = $null; NativeError = $null
                Source = $failure.GetType().FullName; Message = $failure.Message
            })
        }
    }
    finally {
        if ($null -ne $schema) { $schema.Dispose() }
        if ($null -ne $connection) { $connection.Dispose() }
        $connectionString = $null
        $result.FinishedUtc = [DateTime]::UtcNow.ToString('o')
    }
    return [pscustomobject]$result
}

function Invoke-DsnArchitecture {
    param(
        [int]$Bits, [string]$ScriptPath, [string]$Name, [string]$SchemaCollection,
        [int]$SampleLimit, [int]$Timeout,
        [Management.Automation.PSCredential]$DatabaseCredential
    )
    $directory = if ($Bits -eq 32) { 'SysWOW64' }
        elseif ([Environment]::Is64BitProcess) { 'System32' } else { 'Sysnative' }
    $path = Join-Path $env:WINDIR "$directory\WindowsPowerShell\v1.0\powershell.exe"
    $scriptLiteral = $ScriptPath.Replace("'", "''")
    $command = @"
`$ErrorActionPreference = 'Stop'
try {
    `$request = [Management.Automation.PSSerializer]::Deserialize([Console]::In.ReadToEnd())
    & '$scriptLiteral' -Worker -Dsn `$request.Dsn -Collection `$request.Collection -MaxSchemaRows `$request.MaxSchemaRows -Credential `$request.Credential
}
catch { [Console]::Error.WriteLine(`$_.Exception.Message); exit 1 }
"@
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $path
    $start.Arguments = '-NoLogo -NoProfile -NonInteractive -EncodedCommand ' +
        [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    $bytes = $null
    $payload = $null
    try {
        $request = [pscustomobject]@{
            Dsn = $Name; Collection = $SchemaCollection; MaxSchemaRows = $SampleLimit
            Credential = $DatabaseCredential
        }
        $payload = [Management.Automation.PSSerializer]::Serialize($request)
        $null = $process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
        $process.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
        $process.StandardInput.Close()
        if (-not $process.WaitForExit($Timeout * 1000)) {
            $process.Kill()
            $process.WaitForExit()
            throw "$Bits-bit worker timed out after $Timeout seconds; only that worker was stopped."
        }
        $output = $stdout.GetAwaiter().GetResult().Trim()
        $errorOutput = $stderr.GetAwaiter().GetResult().Trim()
        if ($process.ExitCode -ne 0) { throw "Worker exit $($process.ExitCode): $errorOutput" }
        $result = $output | ConvertFrom-Json
        if ($null -eq $result -or $result.Bitness -ne $Bits -or
            $result.Dsn -cne $Name -or $result.Collection -cne $SchemaCollection -or
            $result.Connection -notin @('PASS', 'FAIL') -or
            $result.Schema -notin @('PASS', 'FAIL', 'NOT RUN') -or
            ($result.Schema -eq 'PASS' -and ($result.Connection -ne 'PASS' -or
                $null -eq $result.RowCount -or $result.RowCount -lt 0))) {
            throw 'Worker returned invalid or mismatched results.'
        }
        return $result
    }
    catch {
        return [pscustomobject]@{
            Bitness = $Bits; Dsn = $Name; Collection = $SchemaCollection
            Connection = 'NOT RUN'; Schema = 'NOT RUN'; RowCount = $null
            Stage = 'Worker'; Driver = $null; Sample = @()
            Errors = @([pscustomobject]@{
                SQLState = $null; NativeError = $null; Source = 'Worker'
                Message = $_.Exception.Message
            })
        }
    }
    finally {
        if ($null -ne $bytes) { [Array]::Clear($bytes, 0, $bytes.Length) }
        $payload = $null
        $process.Dispose()
    }
}

if ($MyInvocation.InvocationName -eq '.') { return }
if ([string]::IsNullOrWhiteSpace($Dsn)) { throw 'Specify -Dsn with the configured data source name.' }
if ($UseDsnCredentials -and $null -ne $Credential) { throw 'Use -Credential OR -UseDsnCredentials, not both.' }
if ($Worker) {
    Invoke-DsnSchemaProbe -Name $Dsn -SchemaCollection $Collection -SampleLimit $MaxSchemaRows `
        -DatabaseCredential $Credential | ConvertTo-Json -Depth 8 -Compress
    return
}
if (-not [Environment]::Is64BitOperatingSystem) { throw 'Dual-bitness testing requires 64-bit Windows.' }
if (-not $UseDsnCredentials -and $null -eq $Credential) {
    $Credential = Get-Credential -Message "Database credentials for DSN '$Dsn' (used for both architectures)"
    if ($null -eq $Credential) { throw 'Credential entry cancelled; no connections attempted.' }
}
$results = @()
foreach ($bits in @(32, 64)) {
    $result = Invoke-DsnArchitecture -Bits $bits -ScriptPath $PSCommandPath -Name $Dsn `
        -SchemaCollection $Collection -SampleLimit $MaxSchemaRows -Timeout $TimeoutSeconds `
        -DatabaseCredential $Credential
    $results += $result
    Write-Output "=== $bits-bit: Connection=$($result.Connection); Schema=$($result.Schema); Stage=$($result.Stage) ==="
    if ($result.Driver) { Write-Output "Driver: $($result.Driver)" }
    if ($result.Schema -eq 'PASS') {
        Write-Output "GetSchema('$Collection'): $($result.RowCount) metadata rows; displaying up to $MaxSchemaRows."
        if (@($result.Sample).Count -gt 0) { $result.Sample | Format-Table -AutoSize }
        if ($result.RowCount -eq 0) {
            Write-Output 'Retrieval succeeded but returned no visible objects; this does not prove access to any table.'
        }
    }
    foreach ($errorRecord in @($result.Errors)) {
        Write-Output ("SQLSTATE={0}; NativeError={1}; {2}" -f
            $errorRecord.SQLState, $errorRecord.NativeError, $errorRecord.Message)
    }
}
$results | Format-Table Bitness, Connection, Schema, RowCount, Stage -AutoSize
$successful = @($results | Where-Object { $_.Schema -eq 'PASS' })
if ($successful.Count -gt 0) {
    Write-Output ('Schema retrieval succeeded using: ' + (($successful | ForEach-Object { "$($_.Bitness)-bit" }) -join ', '))
    Write-Output 'Metadata retrieval does not verify SELECT permissions on the listed tables.'
}
else { Write-Output 'Schema retrieval did not succeed in either architecture.' }
$Credential = $null
if ($successful.Count -ne 2) { exit 1 }
