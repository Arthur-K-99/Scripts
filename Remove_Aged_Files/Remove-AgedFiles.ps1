<#
.SYNOPSIS
    Removes files older than a retention period, with optional ACL repair for stubborn SMB files.

.DESCRIPTION
    Remove-AgedFiles.ps1 is designed for scheduled cleanup jobs such as PDQ Deploy tasks.
    By default, it only deletes aged files. Permission repair and ownership takeover are
    explicit opt-in switches.

    When a deletion fails, the script can optionally:
      1. Enable inherited ACLs from the parent folder.
      2. Retry deletion.
      3. Take ownership of the item.
      4. Enable inherited ACLs again.
      5. Retry deletion one final time.

    Ownership takeover is useful for old application-generated files that were created
    with protected or incomplete ACLs. Use it carefully.

.PARAMETER Path
    One or more root folders to scan. The configured root folders themselves are never deleted.

.PARAMETER OlderThanDays
    Files with LastWriteTime earlier than this many days ago are cleanup candidates.
    The default is 30.

.PARAMETER RemoveEmptyDirectories
    Removes empty subdirectories after old files have been processed.

.PARAMETER RepairInheritance
    If a deletion fails, try to enable inherited ACLs on the item and delete it again.

.PARAMETER TakeOwnershipOnFailure
    If deletion and ACL inheritance repair both fail, take ownership, enable inheritance,
    and retry deletion. Requires -RepairInheritance.

.PARAMETER IncludeExtension
    Optional list of file extensions to delete, such as .pdf or .log. If omitted, all
    file extensions are eligible.

.PARAMETER ExcludePath
    Optional list of files or folders to skip. If a folder is excluded, everything under
    that folder is skipped.

.PARAMETER MaxFiles
    Safety limit for the number of aged files processed per root path. The default is 10000.

.PARAMETER PassThru
    Writes structured result objects after processing.

.EXAMPLE
    .\Remove-AgedFiles.ps1 -Path "\\server\share\Archive" -OlderThanDays 30 -WhatIf

    Shows what would be deleted without changing anything.

.EXAMPLE
    .\Remove-AgedFiles.ps1 -Path "\\nas01\Operations\Faxes\Archived", "\\nas01\Operations\Faxes\Sent" -OlderThanDays 30 -RemoveEmptyDirectories -RepairInheritance -TakeOwnershipOnFailure -IncludeExtension .pdf

    Cleans old PDF files, repairs broken inheritance only when needed, takes ownership only
    after repair fails, removes empty folders, and returns a PDQ-friendly exit code.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [Alias('LiteralPath')]
    [string[]]$Path,

    [ValidateRange(1, 36500)]
    [int]$OlderThanDays = 30,

    [switch]$RemoveEmptyDirectories,

    [switch]$RepairInheritance,

    [switch]$TakeOwnershipOnFailure,

    [ValidateNotNullOrEmpty()]
    [string[]]$IncludeExtension,

    [ValidateNotNullOrEmpty()]
    [string[]]$ExcludePath,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxFiles = 10000,

    [switch]$PassThru
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:LogPrefix = '[AGED-FILE-CLEANUP]'
$script:Cmdlet = $PSCmdlet
$script:Failures = [System.Collections.Generic.List[string]]::new()
$script:Results = [System.Collections.Generic.List[object]]::new()

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "$script:LogPrefix $Message"
}

function Add-CleanupResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ItemType,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [string]$Message
    )

    $script:Results.Add([pscustomobject]@{
        Path = $Path
        ItemType = $ItemType
        Action = $Action
        Status = $Status
        Message = $Message
    })
}

function Add-CleanupFailure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $script:Failures.Add($Message)
    Write-Warning "$script:LogPrefix $Message"
}

function Test-DangerousRootPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    $trimmedPath = $InputPath.TrimEnd('\')

    if ($trimmedPath -match '^[A-Za-z]:$') {
        return $true
    }

    if ($trimmedPath -match '^\\\\[^\\]+\\[^\\]+$') {
        return $true
    }

    return $false
}

function Resolve-CleanupRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    if (Test-DangerousRootPath -InputPath $InputPath) {
        throw "Refusing to scan dangerous root path: $InputPath"
    }

    $item = Get-Item -LiteralPath $InputPath -ErrorAction Stop

    if (-not $item.PSIsContainer) {
        throw "Cleanup path must be a directory: $InputPath"
    }

    $resolvedPath = $item.FullName

    if (Test-DangerousRootPath -InputPath $resolvedPath) {
        throw "Refusing to scan dangerous root path: $resolvedPath"
    }

    return $item
}

function ConvertTo-NormalizedExtensionSet {
    param(
        [string[]]$Extension
    )

    $set = @{}

    foreach ($entry in @($Extension)) {
        if ([string]::IsNullOrWhiteSpace($entry)) {
            continue
        }

        $normalized = $entry.Trim().ToLowerInvariant()

        if (-not $normalized.StartsWith('.')) {
            $normalized = ".$normalized"
        }

        $set[$normalized] = $true
    }

    return $set
}

function ConvertTo-NormalizedPathList {
    param(
        [string[]]$InputPath
    )

    $paths = [System.Collections.Generic.List[string]]::new()

    foreach ($entry in @($InputPath)) {
        if ([string]::IsNullOrWhiteSpace($entry)) {
            continue
        }

        try {
            $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($entry)
        }
        catch {
            $resolved = $entry
        }

        $paths.Add($resolved.TrimEnd('\'))
    }

    # Prevent PowerShell from unrolling an empty generic list into $null.
    return ,$paths
}

function Test-ExcludedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidatePath,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$NormalizedExcludePath
    )

    $trimmedCandidate = $CandidatePath.TrimEnd('\')

    foreach ($excluded in $NormalizedExcludePath) {
        if ($trimmedCandidate.Equals($excluded, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }

        if ($trimmedCandidate.StartsWith("$excluded\", [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Invoke-IcaclsInheritanceRepair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemPath
    )

    $output = & icacls.exe $ItemPath /inheritancelevel:e /Q 2>&1
    $exitCode = $LASTEXITCODE

    foreach ($line in @($output)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            Write-Host "$script:LogPrefix icacls: $line"
        }
    }

    return ($exitCode -eq 0)
}

function Invoke-OwnershipTakeover {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemPath
    )

    $output = & takeown.exe /F $ItemPath 2>&1
    $exitCode = $LASTEXITCODE

    foreach ($line in @($output)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            Write-Host "$script:LogPrefix takeown: $line"
        }
    }

    return ($exitCode -eq 0)
}

function Remove-CleanupItem {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo]$Item,

        [Parameter(Mandatory = $true)]
        [ValidateSet('File', 'Directory')]
        [string]$ItemType
    )

    $itemPath = $Item.FullName

    if (-not $script:Cmdlet.ShouldProcess($itemPath, "Delete $ItemType")) {
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'Delete' -Status 'WhatIf' -Message 'Would delete item.'
        return
    }

    try {
        Remove-Item -LiteralPath $itemPath -Force -ErrorAction Stop
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'Delete' -Status 'Success' -Message 'Deleted item.'
        Write-Log "Deleted $($ItemType.ToLowerInvariant()): $itemPath"
        return
    }
    catch {
        Write-Warning "$script:LogPrefix Normal deletion failed: $itemPath"
        Write-Warning "$script:LogPrefix Reason: $($_.Exception.Message)"
    }

    if (-not $RepairInheritance) {
        $message = "Could not delete $($ItemType.ToLowerInvariant()) and ACL repair is disabled: $itemPath"
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'Delete' -Status 'Failed' -Message $message
        Add-CleanupFailure -Message $message
        return
    }

    if (-not $script:Cmdlet.ShouldProcess($itemPath, 'Enable ACL inheritance')) {
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'RepairInheritance' -Status 'Skipped' -Message 'ACL inheritance repair was declined.'
        return
    }

    Write-Warning "$script:LogPrefix Attempting to enable inheritance: $itemPath"

    if (Invoke-IcaclsInheritanceRepair -ItemPath $itemPath) {
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'RepairInheritance' -Status 'Success' -Message 'Enabled inherited ACLs.'

        try {
            Remove-Item -LiteralPath $itemPath -Force -ErrorAction Stop
            Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'DeleteAfterInheritanceRepair' -Status 'Success' -Message 'Deleted item after inheritance repair.'
            Write-Log "Deleted $($ItemType.ToLowerInvariant()) after inheritance repair: $itemPath"
            return
        }
        catch {
            Write-Warning "$script:LogPrefix Deletion still failed after inheritance repair."
            Write-Warning "$script:LogPrefix Reason: $($_.Exception.Message)"
        }
    }
    else {
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'RepairInheritance' -Status 'Failed' -Message 'Could not enable inherited ACLs.'
        Write-Warning "$script:LogPrefix Could not enable inheritance using the current permissions."
    }

    if (-not $TakeOwnershipOnFailure) {
        $message = "Could not delete $($ItemType.ToLowerInvariant()) after inheritance repair and ownership takeover is disabled: $itemPath"
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'DeleteAfterInheritanceRepair' -Status 'Failed' -Message $message
        Add-CleanupFailure -Message $message
        return
    }

    if (-not $script:Cmdlet.ShouldProcess($itemPath, 'Take ownership')) {
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'TakeOwnership' -Status 'Skipped' -Message 'Ownership takeover was declined.'
        return
    }

    Write-Warning "$script:LogPrefix Attempting to take ownership: $itemPath"

    if (-not (Invoke-OwnershipTakeover -ItemPath $itemPath)) {
        $message = "Could not take ownership: $itemPath"
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'TakeOwnership' -Status 'Failed' -Message $message
        Add-CleanupFailure -Message $message
        return
    }

    Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'TakeOwnership' -Status 'Success' -Message 'Took ownership of item.'

    if (-not $script:Cmdlet.ShouldProcess($itemPath, 'Enable ACL inheritance after ownership takeover')) {
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'RepairInheritanceAfterOwnership' -Status 'Skipped' -Message 'ACL inheritance repair after ownership takeover was declined.'
        return
    }

    Write-Warning "$script:LogPrefix Ownership changed. Enabling inheritance: $itemPath"

    if (-not (Invoke-IcaclsInheritanceRepair -ItemPath $itemPath)) {
        $message = "Took ownership but could not enable inheritance: $itemPath"
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'RepairInheritanceAfterOwnership' -Status 'Failed' -Message $message
        Add-CleanupFailure -Message $message
        return
    }

    Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'RepairInheritanceAfterOwnership' -Status 'Success' -Message 'Enabled inherited ACLs after ownership takeover.'

    try {
        Remove-Item -LiteralPath $itemPath -Force -ErrorAction Stop
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'DeleteAfterOwnership' -Status 'Success' -Message 'Deleted item after ownership takeover.'
        Write-Log "Deleted $($ItemType.ToLowerInvariant()) after taking ownership: $itemPath"
    }
    catch {
        $message = "Could not delete $($ItemType.ToLowerInvariant()) after taking ownership: $itemPath - $($_.Exception.Message)"
        Add-CleanupResult -Path $itemPath -ItemType $ItemType -Action 'DeleteAfterOwnership' -Status 'Failed' -Message $message
        Add-CleanupFailure -Message $message
    }
}

if ($TakeOwnershipOnFailure -and -not $RepairInheritance) {
    Write-Error "$script:LogPrefix -TakeOwnershipOnFailure requires -RepairInheritance." -ErrorAction Continue
    exit 1
}

$cutoffDate = (Get-Date).AddDays(-$OlderThanDays)
$extensionSet = ConvertTo-NormalizedExtensionSet -Extension $IncludeExtension
$normalizedExcludePaths = ConvertTo-NormalizedPathList -InputPath $ExcludePath

Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Log "Deleting files older than: $cutoffDate"
Write-Log "Maximum aged files per root path: $MaxFiles"

foreach ($inputPath in $Path) {
    try {
        $root = Resolve-CleanupRoot -InputPath $inputPath
    }
    catch {
        Add-CleanupFailure -Message $_.Exception.Message
        continue
    }

    Write-Log "Scanning: $($root.FullName)"

    try {
        $oldFiles = @(Get-ChildItem -LiteralPath $root.FullName -Recurse -File -ErrorAction Stop |
            Where-Object {
                $_.LastWriteTime -lt $cutoffDate -and
                ($extensionSet.Count -eq 0 -or $extensionSet.ContainsKey($_.Extension.ToLowerInvariant())) -and
                (-not (Test-ExcludedPath -CandidatePath $_.FullName -NormalizedExcludePath $normalizedExcludePaths))
            })

        if ($oldFiles.Count -gt $MaxFiles) {
            Add-CleanupFailure -Message "Found $($oldFiles.Count) aged files under $($root.FullName), which exceeds -MaxFiles $MaxFiles. No files were deleted from this root."
            continue
        }

        foreach ($file in $oldFiles) {
            Remove-CleanupItem -Item $file -ItemType File
        }
    }
    catch {
        Add-CleanupFailure -Message "Could not scan files under $($root.FullName) - $($_.Exception.Message)"
    }

    if (-not $RemoveEmptyDirectories) {
        continue
    }

    try {
        $folders = @(Get-ChildItem -LiteralPath $root.FullName -Recurse -Directory -ErrorAction Stop |
            Where-Object {
                -not (Test-ExcludedPath -CandidatePath $_.FullName -NormalizedExcludePath $normalizedExcludePaths)
            } |
            Sort-Object -Property FullName -Descending)

        foreach ($folder in $folders) {
            try {
                $contents = @(Get-ChildItem -LiteralPath $folder.FullName -Force -ErrorAction Stop)

                if ($contents.Count -eq 0) {
                    Remove-CleanupItem -Item $folder -ItemType Directory
                }
            }
            catch {
                Add-CleanupFailure -Message "Could not process folder: $($folder.FullName) - $($_.Exception.Message)"
            }
        }
    }
    catch {
        Add-CleanupFailure -Message "Could not scan folders under $($root.FullName) - $($_.Exception.Message)"
    }
}

$deletedCount = @($script:Results | Where-Object { $_.Action -like 'Delete*' -and $_.Status -eq 'Success' }).Count
$whatIfCount = @($script:Results | Where-Object { $_.Status -eq 'WhatIf' }).Count
$repairCount = @($script:Results | Where-Object { $_.Action -like 'RepairInheritance*' -and $_.Status -eq 'Success' }).Count
$takeOwnershipCount = @($script:Results | Where-Object { $_.Action -eq 'TakeOwnership' -and $_.Status -eq 'Success' }).Count

Write-Host '--------------------------------------------------'
Write-Log 'SUMMARY COMPLETE'
Write-Host "Deleted: $deletedCount | WhatIf: $whatIfCount | ACL Repairs: $repairCount | Ownership Takeovers: $takeOwnershipCount | Failures: $($script:Failures.Count)"
Write-Host '--------------------------------------------------'

if ($PassThru) {
    $script:Results
}

if ($script:Failures.Count -gt 0) {
    Write-Error "$script:LogPrefix Cleanup completed with $($script:Failures.Count) failure(s)." -ErrorAction Continue

    foreach ($failure in $script:Failures) {
        Write-Error "$script:LogPrefix $failure" -ErrorAction Continue
    }

    exit 1
}

Write-Log 'Cleanup completed successfully.'
exit 0
