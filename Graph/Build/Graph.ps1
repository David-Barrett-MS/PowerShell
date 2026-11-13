#
# Graph.ps1
#
# By David Barrett, Microsoft Ltd. Use at your own risk.  No warranties are given.
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
Helper functions for interacting with Graph.

.DESCRIPTION
Helper functions for interacting with Graph.  Intended to be used as part of the build process for Graph scripts in this repository (requires Auth.ps1 and Logging.ps1).

.EXAMPLE

#>

params(
<# Graph.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="If specified, all Graph calls will be logged to file (same name as script, ending in .trace).")]
    [switch]$TraceGraphCalls,
<# Graph.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$False,HelpMessage="Not included in template.  Only included to prevent error in code editors.")]
    [switch]$Dummy
)

<# Graph.ps1 %FUNCTIONS_START% #>

$script:traceFile = "$($MyInvocation.InvocationName).trace"
function GetGraphErrorResponse([System.Management.Automation.ErrorRecord]$ErrorRecord)
{
    $response = $ErrorRecord.Exception.Response
    if ($null -eq $response) { return $null }
    $responseHeaders = @{}
    if ($response.Headers -is [System.Net.WebHeaderCollection])
    {
        foreach ($name in $response.Headers.AllKeys) { $responseHeaders[$name] = $response.Headers[$name] }
    }
    else
    {
        foreach ($header in $response.Headers.GetEnumerator()) { $responseHeaders[$header.Key] = $header.Value -join ", " }
    }
    if ($null -ne $response.Content -and $null -ne $response.Content.Headers)
    {
        foreach ($header in $response.Content.Headers.GetEnumerator()) { $responseHeaders[$header.Key] = $header.Value -join ", " }
    }
    $content = $ErrorRecord.ErrorDetails.Message
    if ([string]::IsNullOrEmpty($content))
    {
        if ($null -ne $response.Content)
        {
            try { $content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult() }
            catch { $content = "" }
        }
        elseif ($response -is [System.Net.HttpWebResponse])
        {
            try
            {
                $reader = [System.IO.StreamReader]::new($response.GetResponseStream())
                try { $content = $reader.ReadToEnd() }
                finally { $reader.Dispose() }
            }
            catch { $content = "" }
        }
    }
    $status = [int]$response.StatusCode
    $description = $response.ReasonPhrase
    if ([string]::IsNullOrEmpty($description)) { $description = $response.StatusDescription }
    $rawContent = "HTTP $status $description`r`n"
    foreach ($name in $responseHeaders.Keys) { $rawContent += "${name}: $($responseHeaders[$name])`r`n" }
    $rawContent += "`r`n$content"
    return [PSCustomObject]@{ StatusCode = $status; Headers = $responseHeaders; Content = $content; RawContent = $rawContent }
}

# Full responses and caller-managed redirects are opt-in; existing calls still return content.
function TraceInvokeRestMethod($method, $url, $body, [switch]$ReturnResponse, [switch]$NoRedirect)
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    if ($TraceGraphCalls)
    {
        Add-Content -Path $script:traceFile -Value "$method $url" -Encoding UTF8
    }
    try
    {
        $result = $null
        $invokeParameters = @{
            Method = $method
            Uri = $url
            Headers = $script:headers
            ContentType = "application/json"
            MaximumRedirection = 0  # Disable automatic redirects so we can handle them manually
            UseBasicParsing = $true
        }
        if ($PSVersionTable.PSVersion -ge [version]7.4)
        {
            # PowerShell 7.4+ supports ConnectionTimeoutSeconds
            $invokeParameters.ConnectionTimeoutSeconds = $ConnectionTimeout
        }
        elseif ($ReturnResponse)
        {
            $invokeParameters.TimeoutSec = $ConnectionTimeout
        }

        # Write request headers, excluding Authorization
        if ($TraceGraphCalls)
        {
            $invokeParameters.Headers.GetEnumerator() | Where-Object { $_.Key -ne "Authorization" } | ForEach-Object {
                "$($_.Key): $($_.Value)" | out-file $script:traceFile -Append -Encoding UTF8
            }
            "Date: $([datetime]::UtcNow.ToString('r'))" | out-file $script:traceFile -Append -Encoding UTF8
        }

        if ($method -eq "POST" -or $method -eq "PATCH")
        {            
            if ($TraceGraphCalls)
            {
                "`r`n$body" | out-file $script:traceFile -Append -Encoding UTF8
            }
            $invokeParameters.Body = $body
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = Invoke-WebRequest @invokeParameters -ErrorAction Stop
        $stopwatch.Stop()
        LogVerbose "Web request total time: $($stopwatch.Elapsed.TotalMilliseconds) ms"

        if ($TraceGraphCalls)
        {
            "" | out-file $script:traceFile -Append -Encoding UTF8
            $result.RawContent | out-file $script:traceFile -Append -Encoding UTF8
            "`r`n" | out-file $script:traceFile -Append -Encoding UTF8
        }
    }
    catch
    {
        if ($ReturnResponse -and ($NoRedirect -or !$_.Exception.Response -or $_.Exception.Response.StatusCode -ne 308))
        {
            $errorResponse = GetGraphErrorResponse $_
            if ($TraceGraphCalls)
            {
                "" | out-file $script:traceFile -Append -Encoding UTF8
                if ($null -ne $errorResponse) { $errorResponse.RawContent | out-file $script:traceFile -Append -Encoding UTF8 }
                else { $_ | out-file $script:traceFile -Append -Encoding UTF8 }
                "`r`n" | out-file $script:traceFile -Append -Encoding UTF8
            }
            if ($null -ne $errorResponse) { return $errorResponse }
            throw
        }

        # Check for HTTP 308 Permanent Redirect (archive mailbox redirect)
        if (!$NoRedirect -and $_.Exception.Response -and $_.Exception.Response.StatusCode -eq 308)
        {
            $redirectUrl = $_.Exception.Response.Headers['Location']
            if ($redirectUrl)
            {
                LogVerbose "Received HTTP 308 redirect to: $redirectUrl"
                
                if ($TraceGraphCalls)
                {
                    "" | out-file $script:traceFile -Append -Encoding UTF8
                    "HTTP 308 Permanent Redirect" | out-file $script:traceFile -Append -Encoding UTF8
                    "Location: $redirectUrl" | out-file $script:traceFile -Append -Encoding UTF8
                    "Following redirect..." | out-file $script:traceFile -Append -Encoding UTF8
                    "`r`n" | out-file $script:traceFile -Append -Encoding UTF8
                }
                
                # Follow the redirect by making a new request
                return TraceInvokeRestMethod $method $redirectUrl $body -ReturnResponse:$ReturnResponse
            }
        }
        
        if ($TraceGraphCalls)
        {
            # Check if there is a result object
            if ($result) {
                # If the result contains raw content, log it to the trace file
                if ($result.RawContent) {
                    "" | out-file $script:traceFile -Append -Encoding UTF8
                    $result.RawContent | out-file $script:traceFile -Append -Encoding UTF8
                    "`r`n" | out-file $script:traceFile -Append -Encoding UTF8
                }
                else {
                    # If raw content is not available, log the entire result object
                    $result | out-file $script:traceFile -Append -Encoding UTF8
                }                
            }
            else {
                # If no result object is available, check for exception response details
                if ($Error[0].Exception.Response)
                {
                    # Extract and log response headers if available
                    $headers = $Error[0].Exception.Response.Headers
                    if ($headers -and $headers.Key.Count -gt 0) {
                        for ($i=0; $i -lt $headers.Key.Count; $i++)
                        {
                            "$($headers.Key[$i]): $($headers.Value[$i])" | out-file $script:traceFile -Append -Encoding UTF8
                        }
                        "" | out-file $script:traceFile -Append -Encoding UTF8
                    }
                    else {
                        # Log a message if no response headers are returned
                        Log "No response headers returned" Red
                    }
                }
                else {
                    # Log a message if no further error details are available
                    Log "No further error details available" Red
                }                
            }
            # Log the error object to the trace file
            $Error[0] | out-file $script:traceFile -Append -Encoding UTF8
        }
        Log "$method failed to URL: $url" Red
        if ($Error[0].ErrorDetails.Message)
        {
            $errResponse = $null
            $errResponse = $Error[0].ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($errResponse)
            {
                if ($errResponse.error.details.message)
                {
                    Log "Error: $($errResponse.error.details.message)" Red
                }            
            }
        }
        return $null
    }
    if ($ReturnResponse) { return $result }
    return $result.Content
}

function ResolveGraphRequestUrl([string]$Url, [string]$BaseUrl, [scriptblock]$ValidateUrl)
{
    $target = $null
    if (![Uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$target))
    {
        if ([string]::IsNullOrEmpty($BaseUrl)) { throw "An absolute Graph URL is required." }
        $target = [Uri]::new([Uri]$BaseUrl, $Url)
    }
    if ($target.Scheme -ne "https" -or $target.Host -ne "graph.microsoft.com" -or
        !$target.IsDefaultPort -or $target.UserInfo -or $target.Fragment)
    {
        throw "Rejected an unsafe Graph redirect or pagination URL."
    }
    if ($null -ne $ValidateUrl) { $null = & $ValidateUrl $target.AbsoluteUri }
    return $target.AbsoluteUri
}

# RequestOptions enables strict JSON responses, bounded redirects and per-response observation.
# Supported keys: MaxRedirects, ValidateUrl, OnResponse, OnRedirect.
# ValidateUrl must throw to reject a URL.
function InvokeGraphRequest([string]$method, [string]$url, [object]$body, [hashtable]$RequestOptions)
{
    if ($null -eq $RequestOptions) { $RequestOptions = @{} }
    $redirectLimit = 10
    if ($RequestOptions.ContainsKey("MaxRedirects")) { $redirectLimit = [int]$RequestOptions.MaxRedirects }
    if ($redirectLimit -lt 0 -or $redirectLimit -gt 50) { throw "MaxRedirects must be between 0 and 50." }
    $urlToGet = ResolveGraphRequestUrl $url -ValidateUrl $RequestOptions.ValidateUrl
    $initialUrl = $urlToGet
    $visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    for ($hop = 0; $hop -le $redirectLimit; $hop++)
    {
        if (!$visited.Add($urlToGet)) { throw "Graph redirect loop at $urlToGet" }
        $clientRequestId = [Guid]::NewGuid().ToString()
        $started = [DateTime]::UtcNow
        $response = $null
        $responseHeaders = @{}
        $transportError = $null
        try
        {
            if ($null -eq $script:headers) { $script:headers = @{ "Content-Type" = "application/json" } }
            $script:headers["client-request-id"] = $clientRequestId
            $script:headers["return-client-request-id"] = "true"
            $response = TraceInvokeRestMethod $method $urlToGet $body -ReturnResponse -NoRedirect
            foreach ($name in $response.Headers.Keys) { $responseHeaders[$name] = $response.Headers[$name] -join ", " }
        }
        catch { $transportError = $_.Exception.Message }
        $record = [PSCustomObject]@{
            StartedUtc = $started.ToString("o")
            FinishedUtc = [DateTime]::UtcNow.ToString("o")
            Method = $method
            Url = $urlToGet
            ClientRequestId = $clientRequestId
            Status = $(if ($null -ne $response) { [int]$response.StatusCode } else { $null })
            Headers = $(if ($null -ne $response) { $responseHeaders } else { $null })
            Body = $(if ($null -ne $response) { [string]$response.Content } else { $null })
            TransportError = $transportError
        }
        if ($null -ne $RequestOptions.OnResponse) { $null = & $RequestOptions.OnResponse $record }
        if ($transportError) { throw "Transport failure for ${urlToGet}: $transportError" }
        $data = $null
        if ($record.Status -ne 308 -and ![string]::IsNullOrWhiteSpace($record.Body))
        {
            try { $data = ConvertFrom-Json -InputObject $record.Body -ErrorAction Stop }
            catch { throw "HTTP $($record.Status) returned invalid JSON at $urlToGet." }
        }
        $redirect = $null
        if ($record.Status -eq 308)
        {
            $redirect = $responseHeaders["Location"]
            if ([string]::IsNullOrWhiteSpace($redirect)) { throw "HTTP 308 without a Location header at $urlToGet" }
        }
        elseif ($data.error.code -eq "ErrorArchiveFolderMovedPermanently")
        {
            $redirect = $data.error.message
            if ([string]::IsNullOrWhiteSpace($redirect)) { throw "Archive-moved error without a redirect URL at $urlToGet" }
        }
        if ($redirect)
        {
            $nextUrl = ResolveGraphRequestUrl $redirect $urlToGet -ValidateUrl $RequestOptions.ValidateUrl
            if ($hop -ge $redirectLimit) { throw "Exceeded MaxRedirects ($redirectLimit) for $initialUrl" }
            if ($visited.Contains($nextUrl)) { throw "Graph redirect loop at $nextUrl" }
            if ($null -ne $RequestOptions.OnRedirect) { $null = & $RequestOptions.OnRedirect $urlToGet $nextUrl }
            LogVerbose "Following Graph redirect: $urlToGet -> $nextUrl"
            $urlToGet = $nextUrl
            continue
        }
        if ($record.Status -lt 200 -or $record.Status -ge 300 -or $null -ne $data.error -or
            ($record.Status -eq 200 -and $null -eq $data))
        {
            throw "$method failed: HTTP $($record.Status), code '$($data.error.code)', URL $urlToGet."
        }
        return [PSCustomObject]@{
            Data = $data; Url = $urlToGet; RequestUrl = $initialUrl
            Status = $record.Status; Headers = $responseHeaders; Body = $record.Body
        }
    }
}

function GET($url, [hashtable]$RequestOptions)
{
    if ($null -ne $RequestOptions) { return InvokeGraphRequest "GET" $url $null $RequestOptions }
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "GET" $url
    }
    catch
    {
        Write-Host "GET failed from URL: $url" -ForegroundColor Red
        return $null
    }
    return $result
}

function POST($url, $body, $contentType="application/json")
{
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "POST" $url $body $contentType
    }
    catch
    {
        Write-Host "POST failed to URL: $url" -ForegroundColor Red
        return $null
    }
    # Return result or empty string (HTTP 202 with no content is success)
    if ($null -eq $result) { return "" }
    return $result
}

function PATCH($url, $body)
{
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "PATCH" $url $body
    }
    catch
    {
        Write-Host "PATCH failed at URL: $url" -ForegroundColor Red
    }
    return $result
}

function DELETE($url)
{
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "DELETE" $url
    }
    catch
    {
        Write-Host "DELETE failed at URL: $url" -ForegroundColor Red
        return $null
    }
    # Return result or empty string (HTTP 204 with no content is success)
    if ($null -eq $result) { return "" }
    return $result
}

function rawGET($url)
{
    return TraceInvokeRestMethod "GET" $url $null
}

<# Graph.ps1 %FUNCTIONS_END% #>
