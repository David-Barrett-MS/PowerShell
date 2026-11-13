#
# Mailbox.ps1
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
Mailbox helper function for Microsoft Graph API.  Intended to be used as part of the build process for Graph scripts in this repository.

.DESCRIPTION
Contains functions for accessing mailbox folders and items via Microsoft Graph.

.EXAMPLE

#>

param (
<# Mailbox.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="Mailbox (me to access own mailbox when using delegate permissions).")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox = "me",
<# Mailbox.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$False,HelpMessage="Not included in template.  Only included to prevent error in code editors.")]
    [switch]$Dummy
)


<# Mailbox.ps1 %FUNCTIONS_START% #>
# RequestOptions returns routed response metadata; omitting it retains parsed JSON results.
function GetMailboxFolder($queryString, [string]$BaseUrl = $script:graphBaseUrl, [hashtable]$RequestOptions,
    [string]$FolderId, [switch]$ChildFolders)
{
    $url = $BaseUrl.TrimEnd('/') + "/mailFolders"
    if (![string]::IsNullOrEmpty($FolderId)) { $url += "/" + [Uri]::EscapeDataString($FolderId) }
    if ($ChildFolders)
    {
        if ([string]::IsNullOrEmpty($FolderId)) { throw "FolderId is required when listing child folders." }
        $url += "/childFolders"
    }
    $url += $queryString
    if ($null -ne $RequestOptions) { return GET $url -RequestOptions $RequestOptions }
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetFolderMessages($folderId, $queryString, [string]$BaseUrl = $script:graphBaseUrl, [hashtable]$RequestOptions)
{
    if ($null -ne $RequestOptions) { $folderId = [Uri]::EscapeDataString($folderId) }
    $url = $BaseUrl.TrimEnd('/') + "/mailFolders/$folderId/messages$queryString"
    if ($null -ne $RequestOptions) { return GET $url -RequestOptions $RequestOptions }
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetMessages($queryString)
{
    $url = $script:graphBaseUrl + "messages$queryString"
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetFolderMessageById($folderId, $messageId, $queryString, [string]$BaseUrl = $script:graphBaseUrl, [hashtable]$RequestOptions)
{
    if ($null -ne $RequestOptions)
    {
        $folderId = [Uri]::EscapeDataString($folderId)
        $messageId = [Uri]::EscapeDataString($messageId)
    }
    $url = $BaseUrl.TrimEnd('/') + "/mailFolders/$folderId/messages/$messageId$queryString"
    if ($null -ne $RequestOptions) { return GET $url -RequestOptions $RequestOptions }
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetFolderMessageByIdAttachments($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId/attachments"
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function DeleteFolderMessageById($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId"
    $response = DELETE $url
    if ($null -eq $response)
    {
        return $null
    }
    return $response
}

function CreateMailboxFolder($parentFolderId, $folderName)
{
    $url = $script:graphBaseUrl + "mailFolders/$parentFolderId/childFolders"
    $body = @{
        displayName = $folderName
    } | ConvertTo-Json
    $response = POST $url $body
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function CopyMessageToFolder($messageId, $destinationFolderId)
{
    $url = $script:graphBaseUrl + "messages/$messageId/copy"
    $body = @{
        destinationId = $destinationFolderId
    } | ConvertTo-Json
    $response = POST $url $body
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

<# Mailbox.ps1 %FUNCTIONS_END% #>