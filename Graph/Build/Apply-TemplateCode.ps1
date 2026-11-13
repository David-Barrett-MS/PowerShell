#
# Apply-TemplateCode.ps1
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
Applies shared code to PowerShell scripts in a folder.

.DESCRIPTION
This script is used to apply shared code to PowerShell scripts in a folder.  The shared code is injected into the target script at the location specified in the template.

The following template markers are used to identify the location of the shared code:
<# TemplatePath.ps1 %PARAMS_START% #>
<# TemplatePath.ps1 %PARAMS_END% #>
<# TemplatePath.ps1 %FUNCTIONS_START% #>
<# TemplatePath.ps1 %FUNCTIONS_END% #>
<#
TemplatePath.ps1 is the path to the template file.  The script will look for these markers in the target script and replace them with the shared code.
The script will also create a backup of the original script in the same folder as the target script, with a .bak extension.  If the -BackupFolder parameter is specified, the backup will be created in that folder instead.

.EXAMPLE
# Update existing scripts in a folder
$ScriptFolder = "C:\Github Repos\David-Barrett-MS\PowerShell\Graph"
$SharedFolder = "$ScriptFolder\Build"
$BackupFolder = "$ScriptFolder\Backup"
.\Apply-TemplateCode.ps1 -TargetScriptOrFolder "$ScriptFolder\Test-GraphSearch.ps1" -SharedCodeFolder $SharedFolder -SharedCode @("Auth.ps1", "Logging.ps1", "Graph.ps1") -BackupFolder $BackupFolder
.\Apply-TemplateCode.ps1 -TargetScriptOrFolder $ScriptFolder -SharedCodeFolder $SharedFolder -SharedCode @("Auth.ps1", "Logging.ps1", "Graph.ps1") -BackupFolder $BackupFolder

# Create a new script with shared code applied
$ScriptFolder = "C:\Github Repos\David-Barrett-MS\PowerShell\Graph"
$SharedFolder = "$ScriptFolder\Build"
$NewScript = "$ScriptFolder\New-Script.ps1"
.\Apply-TemplateCode.ps1 -TargetScriptOrFolder $NewScript -SharedCodeFolder $SharedFolder -SharedCode @("Auth.ps1", "Logging.ps1", "Graph.ps1") -Template "Template-Default.ps1"
.\Apply-TemplateCode.ps1 -TargetScriptOrFolder $NewScript -SharedCodeFolder $SharedFolder -SharedCode @("Auth.ps1", "Logging.ps1", "Graph.ps1", "Mailbox.ps1") -Template "Template-Mailbox.ps1"
#>

param (

    [Parameter(Mandatory=$True,HelpMessage="Scripts to which templates will be applied.  If a folder is specified, all .ps1 files in the folder will be processed.  If file does not exist, a new file will be created consisitent with the templates.")]
    $TargetScriptOrFolder,
	
    [Parameter(Mandatory=$True,HelpMessage="Folder where shared code files are located.")]
    $SharedCodeFolder,

    [Parameter(Mandatory=$True,HelpMessage="Shared code file(s) to be injected (any referenced files not in this list will be ignored).  If this parameter is not specified, all template references will be processed.")]
    [string[]]$SharedCode = @("Auth.ps1", "Logging.ps1", "Graph.ps1"),
	
    [Parameter(Mandatory=$False,HelpMessage="Folder where files will be backed up prior to update (if not specified, .ccbak file will be created)")]
    $BackupFolder,

    [Parameter(Mandatory=$False,HelpMessage="If creating a new script, the template to use.")]
    $Template,

    [Parameter(Mandatory=$False,HelpMessage="If specified, removes the shared code modules from the target PowerShell scripts.")]
    [switch]$StripSharedContent
)

$moduleCache = @{}

function ReadModuleFiles()
{
    foreach ($moduleFile in $SharedCode)
    {
        ReadModuleFile $moduleFile
    }
    return $moduleCache.Count
}

function ReadModuleFile()
{
    param (
        $moduleFile
    )
    $moduleCode = $null

    # Check if we have already cached (or failed to find) this module
    if ($moduleCache.ContainsKey($moduleFile))
    {
        return $moduleCache[$moduleFile]
    }
    $moduleCache[$moduleFile] = $false

    if (Test-Path $moduleFile)
    {
        $moduleCode = Get-Content $moduleFile -Raw
    }
    elseif (Test-Path "$SharedCodeFolder\$moduleFile")
    {
        $moduleCode = Get-Content "$SharedCodeFolder\$moduleFile" -Raw
    }
    else
    {
        Write-Host "Module file not found: $moduleFile" -ForegroundColor Red
        return $false
    }

    $templateParams = ""
    $templateFunctions = ""

    # Read $moduleCode and extract the parameters and functions
    $templateMarker = "<# $moduleFile %"
    
    $templateParamsStart = $moduleCode.IndexOf("$($templateMarker)PARAMS_START% #>")
    $templateParamsEnd = $moduleCode.IndexOf("$($templateMarker)PARAMS_END% #>")
    if ($templateParamsStart -gt 0 -and $templateParamsEnd -gt $templateParamsStart)
    {
        $templateParams = $moduleCode.Substring($templateParamsStart,$templateParamsEnd-$templateParamsStart+("$($templateMarker)PARAMS_END% #>").Length)
    }

    $templateFunctionsStart = $moduleCode.IndexOf("$($templateMarker)FUNCTIONS_START% #>")
    $templateFunctionsEnd = $moduleCode.IndexOf("$($templateMarker)FUNCTIONS_END% #>")
    if ($templateFunctionsStart -gt 0 -and $templateFunctionsEnd -gt $templateFunctionsStart)
    {
        $templateFunctions = $moduleCode.Substring($templateFunctionsStart,$templateFunctionsEnd-$templateFunctionsStart+("$($templateMarker)FUNCTIONS_END% #>").Length)
    }

    if ([String]::IsNullOrEmpty($templateParams) -and [String]::IsNullOrEmpty($templateFunctions))
    {
        Write-Host "No module code found in $moduleFile" -ForegroundColor Red
        return $false
    }

    # Add this data to the cache
    $moduleCache[$moduleFile] = @{
        "ParamsStart" = "$($templateMarker)PARAMS_START% #>"
        "ParamsEnd" = "$($templateMarker)PARAMS_END% #>"
        "Params" = $templateParams
        "FunctionsStart" = "$($templateMarker)FUNCTIONS_START% #>"
        "FunctionsEnd" = "$($templateMarker)FUNCTIONS_END% #>"
        "Functions" = $templateFunctions
    }
    return $moduleCache[$moduleFile]
}


# Main script

if ((ReadModuleFiles) -eq 0)
{
    Write-Host "No shared code modules found - exiting" -ForegroundColor Red
    exit
}
Write-Host "Loaded $($moduleCache.Count) shared code modules"
$global:mods = $moduleCache

$scriptIsNew = $false
if (-not (Test-Path $TargetScriptOrFolder))
{
    Write-Host "Creating new script: $TargetScriptOrFolder"

    if (Test-Path $Template)
    {
        Write-Host "Using template: $Template"
        Copy-Item -Path $Template -Destination $TargetScriptOrFolder
        $scriptIsNew = $true
    }
    else
    {
        Write-Host "Template file not found: $Template" -ForegroundColor Red
        exit
    }
}

# Create list of scripts to process, depending on whether TargetScripts is a folder or a file
$scriptSource = Get-Item $TargetScriptOrFolder
if ($scriptSource.PSIsContainer) {
    # Source is a folder - process all .ps1 files in the folder
    $psSourceFiles = Get-ChildItem -Path $TargetScripts -Include *.ps1 -Name
} else {
    # Source is a single file
    $psSourceFiles = @($scriptSource.Name)
    $ScriptFolder = $scriptSource.DirectoryName
}

if (![String]::IsNullOrEmpty($BackupFolder))
{
    # Test if backup folder exists, and attempt to create it if not
    if (-not (Test-Path $BackupFolder))
    {
        try {
            New-Item -ItemType Directory -Path $BackupFolder | Out-Null
            Write-Host "Created backup folder: $BackupFolder"    
        }
        catch {
            Write-Host "Failed to create backup folder: $BackupFolder" -ForegroundColor Red
            $BackupFolder = $null
        }
        
    }
}

foreach ($psSourceFile in $psSourceFiles)
{
    Write-Host "Processing $psSourceFile"
    $psCode = Get-Content "$ScriptFolder\$psSourceFile" -Raw
    $global:origCode = $psCode
    $updatedCode = $psCode

    $script:codeUpdated = $false

    if ($scriptIsNew)
    {
        # For new script we also update the script name in the comments at the top (if present)
        $scriptNameMarker = "<# %SCRIPTNAME% #>"
        $scriptNameIndex = $updatedCode.IndexOf($scriptNameMarker)
        if ($scriptNameIndex -gt 0)
        {
            $beforeScriptName = $updatedCode.Substring(0, $scriptNameIndex)
            $afterScriptName = $updatedCode.Substring($scriptNameIndex + $scriptNameMarker.Length)
            $updatedCode = "$beforeScriptName$psSourceFile$afterScriptName"
            $script:codeUpdated = $true
            Write-Host "Updated script name in comments to: $psSourceFile"
        }
    }

    foreach ($moduleFile in $SharedCode) {
        Write-Host "Applying shared code module: $moduleFile"
        if ($moduleCache.ContainsKey($moduleFile))
        {
            $moduleInfo = $moduleCache[$moduleFile]

            # Replace parameters
            $paramsStart = $updatedCode.IndexOf($moduleInfo.ParamsStart)
            $paramsEnd = $updatedCode.IndexOf($moduleInfo.ParamsEnd)+$moduleInfo.ParamsEnd.Length-1
            Write-Host "ParamsStart: $paramsStart, ParamsEnd: $paramsEnd"
            if ($paramsStart -gt 0 -and $paramsEnd -gt $paramsStart)
            {
                $beforeParams = $updatedCode.Substring(0, $paramsStart)
                $afterParams = $updatedCode.Substring($paramsEnd+1)
                $updatedCode = "$beforeParams$($moduleInfo.Params)$afterParams"
                $script:codeUpdated = $true
                Write-Host "Updated parameters from shared code module: $moduleFile"
            }

            # Replace functions
            $functionsStart = $updatedCode.IndexOf($moduleInfo.FunctionsStart)
            $functionsEnd = $updatedCode.IndexOf($moduleInfo.FunctionsEnd)+$moduleInfo.FunctionsEnd.Length-1
            Write-Host "FunctionsStart: $functionsStart, FunctionsEnd: $functionsEnd"
            if ($functionsStart -gt 0 -and $functionsEnd -gt $functionsStart)
            {
                $beforeFunctions = $updatedCode.Substring(0, $functionsStart)
                $afterFunctions = $updatedCode.Substring($functionsEnd+1)
                $updatedCode = $beforeFunctions + $moduleInfo.Functions + $afterFunctions
                $script:codeUpdated = $true
                Write-Host "Updated functions from shared code module: $moduleFile"
            }
        }
        else
        {
            Write-Host "Shared code module not found: $moduleFile" -ForegroundColor Yellow
        }
    }

    # Check whether the code has actually changed
    if ($script:codeUpdated)
    {
        if ($updatedCode -eq $psCode)
        {
            Write-Host "Template application resulted in no changes to $psSourceFile"
            $script:codeUpdated = $false
        }
    }

    $backupFileName = "$ScriptFolder\$psSourceFile.bak"
    if (![String]::IsNullOrEmpty($BackupFolder))
    {
        if (Test-Path $BackupFolder)
        {
            $backupFileName = "$BackupFolder\$psSourceFile"
        }
    }

    if ($script:codeUpdated)
    {
        if ($scriptIsNew)
        {
            Write-Verbose "Skipping backup of new script"
            $updatedCode | Out-File "$ScriptFolder\$psSourceFile" -Encoding utf8 -NoNewline
            Write-Host "Updated $psSourceFile" -ForegroundColor Green
        }
        else
        {
            if (Test-Path $backupFileName)
            {
                Remove-Item $backupFileName
            }
            Copy-Item -Path "$ScriptFolder\$psSourceFile" -Destination $backupFileName
            if (Test-Path $backupFileName)
            {
                Write-Host "Original backed up to: $backupFileName"
                $updatedCode | Out-File "$ScriptFolder\$psSourceFile" -Encoding utf8 -NoNewline
                Write-Host "Updated $psSourceFile" -ForegroundColor Green
            }
            else
            {
                Write-Host "Failed to backup $psSourceFile - changes not written" -ForegroundColor Red
            }
        }
    }
    else
    {
        Write-Host "No change to $psSourceFile" -ForegroundColor Yellow
    }
}