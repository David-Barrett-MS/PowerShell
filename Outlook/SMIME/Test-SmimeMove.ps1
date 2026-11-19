#requires -Version 5.1
<#
.SYNOPSIS
Tests Outlook Object Model transfers to a temporary PST and restores the lab.
.DESCRIPTION
Uses an already running classic Outlook instance. Logs the stage before each
relevant COM call so a modal dialog can be correlated with the blocking call.
Does not dismiss dialogs or change registry settings.
Source paths are relative to the default store root, for example Inbox\SMIME.
Transfers all direct items, not subfolders, to the same path in a new Unicode PST.
Snapshots IDs from the folder table; logs item open/transfer failures and continues.
Pauses for inspection, then restores moved items and removes the temporary store.
Quits Outlook gracefully, waits for exit, deletes the PST, and reopens the same profile.
Retains the PST if restoration cannot be verified. Cross-store moves change entry IDs.
Uses disposable lab data only. -Copy tests Copy().Move() instead of moving originals.
No pre-transfer confirmation; -LabOnly is required and -WhatIf is supported.
.PREREQUISITES
Windows PowerShell 5.1 or PowerShell 7 on Windows, classic Outlook, and an interactive console.
Save unrelated work first: Outlook will close all its windows during cleanup.
#>
[CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact='None')]
param (
    [Parameter(Mandatory=$True, HelpMessage='Source folder path relative to the default store root, for example Inbox\SMIME.')]
    [ValidateNotNullOrEmpty()]
    [string]$SourceFolderPath,

    [Parameter(Mandatory=$False, HelpMessage='Copy each item before moving the copy to the PST; leave original items in place.')]
    [switch]$Copy,

    [Parameter(Mandatory=$True, HelpMessage='Confirms that the source folder contains only disposable lab data.')]
    [switch]$LabOnly
)

$ErrorActionPreference = 'Stop'
if (-not $LabOnly) {
    throw 'Only disposable lab data is permitted.'
}
$sourcePath = $SourceFolderPath.Trim('\')
foreach ($segment in $sourcePath.Split('\')) {
    if ([string]::IsNullOrWhiteSpace($segment)) {
        throw 'Folder paths must contain nonempty folder names separated by backslashes.'
    }
}
$mode = 'Move'
if ($Copy) { $mode = 'Copy' }
if (-not $PSCmdlet.ShouldProcess(
    "All direct items in default store '$sourcePath'",
    "$mode to a new temporary PST, pause, restore, then close Outlook, delete the PST and reopen Outlook")) {
    return
}

$logPath = Join-Path $PSScriptRoot ('SmimeMove-{0}-{1}.jsonl' -f [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ'), [guid]::NewGuid().ToString('N'))
function Write-Stage {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Operation stage being recorded.')]
        [ValidateNotNullOrEmpty()]
        [string]$Stage,

        [Parameter(Mandatory=$False, HelpMessage='Additional operation and item identifiers.')]
        [hashtable]$Details = @{},

        [Parameter(Mandatory=$False, HelpMessage='Do not interrupt restoration if evidence cannot be written.')]
        [switch]$BestEffort
    )
    $record = [ordered]@{
        Utc = [DateTime]::UtcNow.ToString('o')
        Stage = $Stage
    }
    foreach ($key in $Details.Keys) { $record[$key] = $Details[$key] }
    try {
        [pscustomobject]$record | ConvertTo-Json -Compress | Add-Content -LiteralPath $logPath -Encoding UTF8 -Confirm:$False
    }
    catch {
        if (-not $BestEffort) { throw }
        Write-Warning "Could not write cleanup evidence: $($_.Exception.Message)"
    }
    Write-Host $Stage
}

function Get-LabItemIds {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Folder whose direct item IDs are required.')]
        [ValidateNotNull()]
        [object]$Folder
    )
    $table = $row = $null
    $index = 0
    try {
        # Read stored identities without opening/decrypting protected item objects.
        $table = $Folder.GetTable()
        $table.Columns.RemoveAll()
        $null = $table.Columns.Add('EntryID')
        while (-not $table.EndOfTable) {
            $index++
            $row = $table.GetNextRow()
            $entryId = [string]$row.Item('EntryID')
            if ([string]::IsNullOrEmpty($entryId)) {
                throw "Folder table row $index has no usable EntryID."
            }
            $entryId
        }
    }
    catch {
        Write-Stage 'SnapshotTableFailed' @{
            Row = $index; FolderEntryId = $Folder.EntryID; SourceStoreId = $Folder.StoreID
            Error = $_.Exception.Message; HResult = $_.Exception.HResult
        } -BestEffort
        throw
    }
    finally { $table = $row = $null }
}

function Find-LabEntryId {
    param (
        [Parameter(Mandatory=$True, HelpMessage='MAPI namespace used to compare provider-specific entry identifiers.')]
        [object]$Namespace,

        [Parameter(Mandatory=$True, HelpMessage='Entry identifier to locate, including alternate short/long representations.')]
        [ValidateNotNullOrEmpty()]
        [string]$EntryId,

        [Parameter(Mandatory=$True, HelpMessage='Known entry identifiers to search without opening items.')]
        [AllowEmptyCollection()]
        [Collections.Generic.HashSet[string]]$Ids
    )
    if ($Ids.Contains($EntryId)) { return $EntryId }
    foreach ($candidate in $Ids) {
        if ($Namespace.CompareEntryIDs($candidate, $EntryId)) { return $candidate }
    }
    return $null
}

function Test-LabEntryIdSet {
    param (
        [Parameter(Mandatory=$True, HelpMessage='MAPI namespace used for provider-aware identity comparison.')]
        [object]$Namespace,

        [Parameter(Mandatory=$True, HelpMessage='Expected item identities after restoration.')]
        [AllowEmptyCollection()]
        [Collections.Generic.HashSet[string]]$Expected,

        [Parameter(Mandatory=$True, HelpMessage='Actual folder-table entry identifiers.')]
        [AllowEmptyCollection()]
        [string[]]$Actual
    )
    if ($Expected.Count -ne $Actual.Count) { return $False }
    $remaining = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($id in $Expected) { [void]$remaining.Add($id) }
    foreach ($id in $Actual) {
        $matched = Find-LabEntryId -Namespace $Namespace -EntryId $id -Ids $remaining
        if ($null -eq $matched) { return $False }
        [void]$remaining.Remove($matched)
    }
    return $remaining.Count -eq 0
}

function Assert-LabParent {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Item returned by a transfer.')]
        [AllowNull()]
        [object]$Item,

        [Parameter(Mandatory=$True, HelpMessage='Expected destination folder.')]
        [ValidateNotNull()]
        [object]$Folder
    )
    if ($null -eq $Item) { throw 'Transfer returned no item.' }
    $parent = $Item.Parent
    if ($parent.EntryID -ne $Folder.EntryID -or $parent.StoreID -ne $Folder.StoreID) {
        throw 'Returned item is not in the requested destination folder.'
    }
}

function Find-LabPstStore {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Outlook MAPI namespace.')]
        [ValidateNotNull()]
        [object]$Namespace,

        [Parameter(Mandatory=$True, HelpMessage='Exact path of the PST owned by this run.')]
        [ValidateNotNullOrEmpty()]
        [string]$PstPath
    )
    $stores = $Namespace.Stores
    for ($index = 1; $index -le $stores.Count; $index++) {
        $candidate = $stores.Item($index)
        if ($candidate.FilePath -eq $PstPath) { return $candidate }
    }
}

function Get-LabStoreItemCount {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Root of the folder subtree to inspect.')]
        [ValidateNotNull()]
        [object]$Folder
    )
    $count = $Folder.Items.Count
    $folders = $Folder.Folders
    for ($index = 1; $index -le $folders.Count; $index++) {
        $count += Get-LabStoreItemCount -Folder $folders.Item($index)
    }
    return $count
}

function Wait-LabInspection {
    Write-Host "Inspect '$sourcePath' in the mailbox and temporary PST: $pstPath"
    Write-Host 'Inspect only: do not move, add, delete, or edit items. Close opened messages before continuing.'
    Write-Host 'Save unrelated Outlook work. Cleanup will close Outlook, delete the temporary PST, and reopen the same profile.'
    Write-Host 'Press any key to restore the source and begin cleanup.'
    try {
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
    catch {
        $null = Read-Host 'This host cannot read a single key. Press Enter to restore and clean up'
    }
}

function Resolve-LabFolder {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Root folder of the default Outlook store.')]
        [ValidateNotNull()]
        [object]$RootFolder,

        [Parameter(Mandatory=$True, HelpMessage='Folder path relative to the store root.')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderPath
    )
    $folder = $RootFolder
    foreach ($segment in $FolderPath.Split('\')) {
        try {
            $folder = $folder.Folders.Item($segment)
            if ($null -eq $folder) {
                throw 'Folder lookup returned no folder.'
            }
        }
        catch {
            throw "Cannot resolve '$FolderPath' in the default store at '$segment': $($_.Exception.Message)"
        }
    }
    return $folder
}

function Get-RunningOutlook {
    if (-not ('SmimeMove.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace SmimeMove
{
    public static class NativeMethods
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(
            ref Guid classId,
            IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object instance);

        public static object GetRunningObject(Guid classId)
        {
            object instance;
            GetActiveObject(ref classId, IntPtr.Zero, out instance);
            return instance;
        }
    }
}
'@
    }
    $outlookType = [type]::GetTypeFromProgID('Outlook.Application', $True)
    return [SmimeMove.NativeMethods]::GetRunningObject($outlookType.GUID)
}

function Get-LabOutlookSession {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Attached classic Outlook application whose exact process will be restarted.')]
        [object]$Application,

        [Parameter(Mandatory=$True, HelpMessage='Namespace containing the active Outlook profile name.')]
        [object]$Namespace
    )
    if (-not ('SmimeMove.WindowProcess' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace SmimeMove
{
    public static class WindowProcess
    {
        [DllImport("shlwapi.dll", PreserveSig = false)]
        private static extern void IUnknown_GetWindow(
            [MarshalAs(UnmanagedType.IUnknown)] object window, out IntPtr handle);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);

        public static int GetProcessId(object window)
        {
            IntPtr handle;
            uint processId;
            IUnknown_GetWindow(window, out handle);
            if (handle == IntPtr.Zero || GetWindowThreadProcessId(handle, out processId) == 0)
                throw new InvalidOperationException("Cannot identify the attached Outlook window process.");
            return checked((int)processId);
        }
    }
}
'@
    }
    $window = $Application.ActiveExplorer()
    if ($null -eq $window) { $window = $Application.ActiveInspector() }
    if ($null -eq $window) { throw 'Open an Outlook window before running the test; its process must be identified for restart.' }
    $process = $null
    try {
        $process = [Diagnostics.Process]::GetProcessById([SmimeMove.WindowProcess]::GetProcessId($window))
        $executable = $process.MainModule.FileName
        $profile = [string]$Namespace.CurrentProfileName
        if ([IO.Path]::GetFileName($executable) -ine 'OUTLOOK.EXE' -or -not [IO.File]::Exists($executable)) {
            throw 'Cannot determine the executable of the attached classic Outlook process.'
        }
        if ([string]::IsNullOrWhiteSpace($profile) -or $profile -match '[\x00-\x1f]') {
            throw 'Cannot determine a usable current Outlook profile name.'
        }
        # Acquire a handle now so exit waiting cannot target a reused process ID.
        $null = $process.Handle
        return [pscustomobject]@{ Process = $process; Executable = $executable; Profile = $profile }
    }
    catch {
        if ($null -ne $process) { $process.Dispose() }
        throw
    }
    finally { $window = $null }
}

function Start-LabOutlookSession {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Executable and profile captured from the attached Outlook instance.')]
        [object]$Session
    )
    $quotedProfile = [regex]::Replace($Session.Profile, '(\\*)"', '$1$1\"')
    $quotedProfile = '"' + [regex]::Replace($quotedProfile, '(\\+)$', '$1$1') + '"'
    $started = Start-Process -FilePath $Session.Executable -ArgumentList ('/profile ' + $quotedProfile) -PassThru -ErrorAction Stop
    try {
        if (-not $started.WaitForInputIdle(30000)) { throw 'Outlook was launched but did not become input-idle within 30 seconds.' }
        $started.Refresh()
        if ($started.HasExited -or -not $started.Responding) { throw 'The restarted Outlook process is not running and responsive.' }
    }
    finally { $started.Dispose() }
}

Write-Host "Evidence: $logPath"
$pstPath = Join-Path ([IO.Path]::GetTempPath()) ('SmimeMove-{0}.pst' -f [guid]::NewGuid().ToString('N'))
$namespace = $source = $destination = $pstStore = $pstRoot = $null
$item = $workingItem = $transferred = $restored = $null
$storeAttempted = $False
$outlookSession = $null
$operationFailure = $null
$itemFailures = New-Object 'System.Collections.Generic.List[string]'
$cleanupErrors = New-Object 'System.Collections.Generic.List[string]'
$originalIds = New-Object 'System.Collections.Generic.HashSet[string]'
$attemptedIds = New-Object 'System.Collections.Generic.HashSet[string]'
$movedIds = New-Object 'System.Collections.Generic.HashSet[string]'
$restoredIds = New-Object 'System.Collections.Generic.HashSet[string]'
$sourceCopyIds = New-Object 'System.Collections.Generic.HashSet[string]'
$pstCopyIds = New-Object 'System.Collections.Generic.HashSet[string]'
try {
    Write-Stage 'BeforeAttachToRunningOutlook'
    $outlook = Get-RunningOutlook
    $namespace = $outlook.GetNamespace('MAPI')
    Write-Stage 'BeforeDefaultStoreRoot'
    $root = $namespace.DefaultStore.GetRootFolder()
    Write-Stage 'BeforeOpenSourceFolder'
    $source = Resolve-LabFolder -RootFolder $root -FolderPath $sourcePath
    Write-Stage 'BeforeSnapshotSource'
    $snapshot = @(Get-LabItemIds -Folder $source)
    foreach ($id in $snapshot) {
        if (-not $originalIds.Add($id)) { throw 'Source snapshot contains duplicate EntryIDs.' }
    }
    Write-Stage 'SourceSnapshot' @{ SourcePath = $sourcePath; SourceStoreId = $source.StoreID; ItemCount = $snapshot.Count; Mode = $mode; PstPath = $pstPath }
    if ($snapshot.Count -eq 0) {
        Write-Stage 'SourceEmpty'
        return
    }
    Write-Stage 'BeforeCaptureOutlookSession'
    $outlookSession = Get-LabOutlookSession -Application $outlook -Namespace $namespace
    if (Test-Path -LiteralPath $pstPath) { throw "Refusing to reuse an existing PST: $pstPath" }
    Write-Stage 'BeforeAddStore' @{ PstPath = $pstPath }
    $storeAttempted = $True
    $namespace.AddStoreEx($pstPath, 2)
    $pstStore = Find-LabPstStore -Namespace $namespace -PstPath $pstPath
    if ($null -eq $pstStore) { throw 'The newly attached PST could not be located by its file path.' }
    $pstRoot = $pstStore.GetRootFolder()
    if ($pstRoot.StoreID -eq $source.StoreID) { throw 'Temporary PST unexpectedly matches the source store.' }
    if ((Get-LabStoreItemCount -Folder $pstRoot) -ne 0) { throw 'New PST unexpectedly contains items.' }
    Write-Stage 'BeforeCreateTargetFolder'
    $destination = $pstRoot
    foreach ($segment in $sourcePath.Split('\')) {
        $folders = $destination.Folders
        $nextFolder = $null
        for ($index = 1; $index -le $folders.Count; $index++) {
            $candidate = $folders.Item($index)
            if ($candidate.Name -eq $segment) {
                $nextFolder = $candidate
                break
            }
        }
        if ($null -eq $nextFolder) { $nextFolder = $folders.Add($segment, 6) }
        $destination = $nextFolder
    }
    $folders = $candidate = $nextFolder = $null
    Write-Stage 'TargetReady' @{ PstPath = $pstPath; TargetStoreId = $destination.StoreID; TargetPath = $sourcePath }
    $ordinal = 0
    foreach ($id in $snapshot) {
        $ordinal++
        try {
            Write-Stage 'BeforeGetSourceItem' @{ Item = $ordinal; OriginalEntryId = $id }
            $item = $namespace.GetItemFromID($id, $source.StoreID)
            Assert-LabParent -Item $item -Folder $source
            $workingItem = $item
            if ($Copy) {
                Write-Stage 'BeforeCopy' @{ Item = $ordinal; OriginalEntryId = $id }
                $workingItem = $item.Copy()
                Assert-LabParent -Item $workingItem -Folder $source
                $copyId = [string]$workingItem.EntryID
                if (-not [string]::IsNullOrEmpty($copyId) -and
                    $null -eq (Find-LabEntryId -Namespace $namespace -EntryId $copyId -Ids $originalIds)) {
                    [void]$sourceCopyIds.Add($copyId)
                }
                Write-Stage 'CopyCreated' @{ Item = $ordinal; CopyEntryId = $copyId }
            }
            else {
                [void]$attemptedIds.Add($id)
            }
            Write-Stage 'BeforeMoveToPst' @{ Item = $ordinal; OriginalEntryId = $id; Mode = $mode }
            $transferred = $workingItem.Move($destination)
            Write-Stage 'MoveToPstReturned' @{ Item = $ordinal }
            Assert-LabParent -Item $transferred -Folder $destination
            if ($Copy) { [void]$pstCopyIds.Add([string]$transferred.EntryID) }
            else { [void]$movedIds.Add($id) }
            Write-Stage 'DestinationVerified' @{ Item = $ordinal; OriginalEntryId = $id; TargetEntryId = $transferred.EntryID }
        }
        catch {
            $itemFailures.Add("Item $ordinal ($id): $($_.Exception.Message)")
            Write-Stage 'ItemFailed' @{ Item = $ordinal; OriginalEntryId = $id; Error = $_.Exception.Message; HResult = $_.Exception.HResult } -BestEffort
        }
        finally {
            $item = $workingItem = $transferred = $null
        }
    }
    Write-Stage 'TransferPhaseComplete' @{ Attempted = $snapshot.Count; Failed = $itemFailures.Count; TargetItemCount = $destination.Items.Count }
    Wait-LabInspection
    Write-Stage 'InspectionComplete'
}
catch {
    $operationFailure = $_
    Write-Stage 'Failed' @{ Error = $_.Exception.Message; HResult = $_.Exception.HResult } -BestEffort
}
finally {
    if ($storeAttempted) {
        $safeToDelete = $False
        try {
            Write-Stage 'BeforeRestore' @{ PstPath = $pstPath; Mode = $mode } -BestEffort
            $pstStore = Find-LabPstStore -Namespace $namespace -PstPath $pstPath
            if ($null -eq $pstStore) {
                if (Test-Path -LiteralPath $pstPath) { throw 'PST exists but is not attached; retain it for recovery.' }
            }
            else {
                $pstRoot = $pstStore.GetRootFolder()
                if ($null -ne $destination) {
                    if ($Copy) {
                        # Recover only copies created by this run, never an original.
                        foreach ($id in @(Get-LabItemIds -Folder $source)) {
                            if ($null -ne (Find-LabEntryId -Namespace $namespace -EntryId $id -Ids $sourceCopyIds)) {
                                Write-Stage 'BeforeRecoverSourceCopy' @{ CopyEntryId = $id } -BestEffort
                                $workingItem = $namespace.GetItemFromID($id, $source.StoreID)
                                $transferred = $workingItem.Move($destination)
                                Assert-LabParent -Item $transferred -Folder $destination
                                [void]$pstCopyIds.Add([string]$transferred.EntryID)
                            }
                        }
                    }
                    else {
                        # Snapshot again: EntryIDs change when crossing stores.
                        foreach ($id in @(Get-LabItemIds -Folder $destination)) {
                            try {
                                Write-Stage 'BeforeRestoreItem' @{ TargetEntryId = $id } -BestEffort
                                $workingItem = $namespace.GetItemFromID($id, $destination.StoreID)
                                $restored = $workingItem.Move($source)
                                Assert-LabParent -Item $restored -Folder $source
                                [void]$restoredIds.Add([string]$restored.EntryID)
                                Write-Stage 'ItemRestored' @{ TargetEntryId = $id; RestoredEntryId = $restored.EntryID } -BestEffort
                            }
                            catch {
                                $cleanupErrors.Add($_.Exception.Message)
                                Write-Stage 'RestoreItemFailed' @{ TargetEntryId = $id; Error = $_.Exception.Message } -BestEffort
                            }
                        }
                    }
                }
                $actualSourceIds = @(Get-LabItemIds -Folder $source)
                $actualSourceIdSet = New-Object 'System.Collections.Generic.HashSet[string]'
                foreach ($id in $actualSourceIds) { [void]$actualSourceIdSet.Add($id) }
                $expectedSourceIds = New-Object 'System.Collections.Generic.HashSet[string]'
                foreach ($id in $snapshot) {
                    if ($Copy -or -not $attemptedIds.Contains($id) -or
                        (-not $movedIds.Contains($id) -and
                         $null -ne (Find-LabEntryId -Namespace $namespace -EntryId $id -Ids $actualSourceIdSet))) {
                        [void]$expectedSourceIds.Add($id)
                    }
                }
                foreach ($id in $restoredIds) { [void]$expectedSourceIds.Add($id) }
                $sourceIdsMatch = Test-LabEntryIdSet -Namespace $namespace -Expected $expectedSourceIds -Actual $actualSourceIds
                Write-Stage 'SourceReconciliation' @{
                    OriginalCount = $snapshot.Count; ExpectedCount = $expectedSourceIds.Count
                    ActualCount = $actualSourceIds.Count; IdentitiesMatch = $sourceIdsMatch
                } -BestEffort
                if ($actualSourceIds.Count -ne $snapshot.Count -or -not $sourceIdsMatch) {
                    throw 'Source items could not be reconciled with the original snapshot; retain the PST.'
                }
                $pstItemCount = Get-LabStoreItemCount -Folder $pstRoot
                if ($Copy) {
                    $remainingIds = @()
                    if ($null -ne $destination) { $remainingIds = @(Get-LabItemIds -Folder $destination) }
                    if ($pstItemCount -ne $remainingIds.Count) { throw 'Unexpected items elsewhere in the PST; retain it.' }
                    foreach ($id in $remainingIds) {
                        if ($null -eq (Find-LabEntryId -Namespace $namespace -EntryId $id -Ids $pstCopyIds)) {
                            throw 'Unidentified item in the PST; retain it.'
                        }
                    }
                }
                elseif ($pstItemCount -ne 0) {
                    throw 'Items remain in the PST; retain it for recovery.'
                }
                if ($cleanupErrors.Count -ne 0) { throw 'One or more items could not be restored; retain the PST.' }
                Write-Stage 'SourceRestored' @{ ItemCount = $actualSourceIds.Count } -BestEffort
                Write-Stage 'BeforeRemoveStore' @{ PstPath = $pstPath } -BestEffort
                $namespace.RemoveStore($pstRoot)
                if ($null -ne (Find-LabPstStore -Namespace $namespace -PstPath $pstPath)) {
                    throw 'PST is still attached; it will not be deleted.'
                }
                $safeToDelete = $True
                Write-Stage 'StoreDetached' @{ PstPath = $pstPath } -BestEffort
            }
        }
        catch {
            $cleanupErrors.Add($_.Exception.Message)
            Write-Stage 'CleanupFailed' @{ Error = $_.Exception.Message; PstPath = $pstPath } -BestEffort
        }
        finally {
            $item = $workingItem = $transferred = $restored = $destination = $pstRoot = $pstStore = $folders = $candidate = $nextFolder = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
        }
        if ($safeToDelete) {
            $outlookExited = $False
            try {
                Write-Stage 'BeforeQuitOutlook' @{ ProcessId = $outlookSession.Process.Id } -BestEffort
                try { $outlook.Quit() }
                finally {
                    $outlook = $namespace = $source = $root = $null
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                    [GC]::Collect()
                }
                if (-not $outlookSession.Process.WaitForExit(60000)) {
                    throw 'Outlook did not exit within 60 seconds. It was not force-terminated; the PST is retained.'
                }
                $outlookExited = $True
                Write-Stage 'OutlookExited' @{ ProcessId = $outlookSession.Process.Id } -BestEffort
            }
            catch {
                $cleanupErrors.Add($_.Exception.Message)
                Write-Stage 'OutlookShutdownFailed' @{ Error = $_.Exception.Message; PstPath = $pstPath } -BestEffort
                $outlookExited = $outlookSession.Process.HasExited
            }
            if ($outlookExited) {
                try {
                    for ($attempt = 1; $attempt -le 20; $attempt++) {
                        try {
                            if (Test-Path -LiteralPath $pstPath) { Remove-Item -LiteralPath $pstPath -Force -Confirm:$False }
                            Write-Stage 'PstDeleted' @{ PstPath = $pstPath } -BestEffort
                            break
                        }
                        catch {
                            if ($attempt -eq 20) {
                                $cleanupErrors.Add("Detached PST could not be deleted: $($_.Exception.Message)")
                                Write-Stage 'PstDeleteFailed' @{ Error = $_.Exception.Message; PstPath = $pstPath } -BestEffort
                            }
                            else { Start-Sleep -Milliseconds 250 }
                        }
                    }
                }
                finally {
                    try {
                        Write-Stage 'BeforeRestartOutlook' -BestEffort
                        Start-LabOutlookSession -Session $outlookSession
                        Write-Stage 'OutlookRestarted' -BestEffort
                    }
                    catch {
                        $cleanupErrors.Add("Outlook restart failed: $($_.Exception.Message)")
                        Write-Stage 'OutlookRestartFailed' @{ Error = $_.Exception.Message } -BestEffort
                    }
                }
            }
        }
        if ($cleanupErrors.Count -ne 0) {
            Write-Warning "Cleanup incomplete; see the log. Retain any existing PST until all originals are recovered: $pstPath"
        }
    }
    if ($null -ne $outlookSession) { $outlookSession.Process.Dispose() }
}
if ($cleanupErrors.Count -ne 0) { throw ("Cleanup incomplete: " + ($cleanupErrors -join '; ') + ". PST: $pstPath") }
if ($null -ne $operationFailure) { throw $operationFailure }
if ($itemFailures.Count -ne 0) { throw ("Test completed with item failures (cleanup completed): " + ($itemFailures -join '; ')) }
Write-Stage 'Completed'
