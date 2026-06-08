<#
.SYNOPSIS
    Lightweight validation harness for Remove-AgedFiles.ps1.

.DESCRIPTION
    This intentionally avoids external test dependencies such as Pester. It creates temporary
    folders, runs Remove-AgedFiles.ps1 in child PowerShell processes, and verifies the
    expected file-system state afterward.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:TestCount = 0
$script:FailedCount = 0

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Remove-AgedFiles.ps1'

if ($PSVersionTable.PSEdition -eq 'Core') {
    $powerShell = [System.IO.Path]::Combine($PSHOME, 'pwsh.exe')
}
else {
    $powerShell = [System.IO.Path]::Combine($PSHOME, 'powershell.exe')
}

function New-TestRoot {
    $root = Join-Path ([System.IO.Path]::GetTempPath()) ("RemoveAgedFilesTest_{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -Path $root -ItemType Directory -Force | Out-Null
    return $root
}

function Remove-TestRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function Invoke-CleanupScript {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList
    )

    $previousErrorActionPreference = $ErrorActionPreference

    try {
        # Windows PowerShell 5.1 can surface native stderr as NativeCommandError.
        # Some tests intentionally expect the child script to exit 1, so capture
        # that output without letting it abort the harness.
        $ErrorActionPreference = 'Continue'
        $output = & $powerShell -NoProfile -ExecutionPolicy Bypass -File $scriptPath @ArgumentList 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = @($output)
    }
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Condition,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Invoke-TestCase {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $script:TestCount++
    Write-Host "[TEST] $Name"

    try {
        & $ScriptBlock
        Write-Host "[PASS] $Name"
    }
    catch {
        $script:FailedCount++
        Write-Host "[FAIL] $Name"
        Write-Host "       $($_.Exception.Message)"
    }
}

Invoke-TestCase -Name 'WhatIf keeps old files in place' -ScriptBlock {
    $root = New-TestRoot

    try {
        $oldFile = Join-Path $root 'old.log'
        Set-Content -LiteralPath $oldFile -Value 'old'
        (Get-Item -LiteralPath $oldFile).LastWriteTime = (Get-Date).AddDays(-10)

        $result = Invoke-CleanupScript -ArgumentList @('-Path', $root, '-OlderThanDays', '1', '-WhatIf')

        Assert-True -Condition ($result.ExitCode -eq 0) -Message "Expected exit code 0, got $($result.ExitCode)."
        Assert-True -Condition (Test-Path -LiteralPath $oldFile) -Message 'Old file should remain during WhatIf.'
    }
    finally {
        Remove-TestRoot -Path $root
    }
}

Invoke-TestCase -Name 'Deletes only files older than retention period' -ScriptBlock {
    $root = New-TestRoot

    try {
        $oldFile = Join-Path $root 'old.log'
        $newFile = Join-Path $root 'new.log'

        Set-Content -LiteralPath $oldFile -Value 'old'
        Set-Content -LiteralPath $newFile -Value 'new'
        (Get-Item -LiteralPath $oldFile).LastWriteTime = (Get-Date).AddDays(-10)

        $result = Invoke-CleanupScript -ArgumentList @('-Path', $root, '-OlderThanDays', '1')

        Assert-True -Condition ($result.ExitCode -eq 0) -Message "Expected exit code 0, got $($result.ExitCode)."
        Assert-True -Condition (-not (Test-Path -LiteralPath $oldFile)) -Message 'Old file should be deleted.'
        Assert-True -Condition (Test-Path -LiteralPath $newFile) -Message 'New file should remain.'
    }
    finally {
        Remove-TestRoot -Path $root
    }
}

Invoke-TestCase -Name 'Extension filter limits deletion candidates' -ScriptBlock {
    $root = New-TestRoot

    try {
        $oldLog = Join-Path $root 'old.log'
        $oldPdf = Join-Path $root 'old.pdf'

        Set-Content -LiteralPath $oldLog -Value 'old log'
        Set-Content -LiteralPath $oldPdf -Value 'old pdf'
        (Get-Item -LiteralPath $oldLog).LastWriteTime = (Get-Date).AddDays(-10)
        (Get-Item -LiteralPath $oldPdf).LastWriteTime = (Get-Date).AddDays(-10)

        $result = Invoke-CleanupScript -ArgumentList @('-Path', $root, '-OlderThanDays', '1', '-IncludeExtension', '.pdf')

        Assert-True -Condition ($result.ExitCode -eq 0) -Message "Expected exit code 0, got $($result.ExitCode)."
        Assert-True -Condition (Test-Path -LiteralPath $oldLog) -Message 'Old .log file should remain.'
        Assert-True -Condition (-not (Test-Path -LiteralPath $oldPdf)) -Message 'Old .pdf file should be deleted.'
    }
    finally {
        Remove-TestRoot -Path $root
    }
}

Invoke-TestCase -Name 'Excluded paths are not deleted' -ScriptBlock {
    $root = New-TestRoot

    try {
        $excludedFolder = Join-Path $root 'excluded'
        $includedFolder = Join-Path $root 'included'
        $excludedFile = Join-Path $excludedFolder 'old.log'
        $includedFile = Join-Path $includedFolder 'old.log'

        New-Item -Path $excludedFolder -ItemType Directory -Force | Out-Null
        New-Item -Path $includedFolder -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath $excludedFile -Value 'old'
        Set-Content -LiteralPath $includedFile -Value 'old'
        (Get-Item -LiteralPath $excludedFile).LastWriteTime = (Get-Date).AddDays(-10)
        (Get-Item -LiteralPath $includedFile).LastWriteTime = (Get-Date).AddDays(-10)

        $result = Invoke-CleanupScript -ArgumentList @('-Path', $root, '-OlderThanDays', '1', '-ExcludePath', $excludedFolder)

        Assert-True -Condition ($result.ExitCode -eq 0) -Message "Expected exit code 0, got $($result.ExitCode)."
        Assert-True -Condition (Test-Path -LiteralPath $excludedFile) -Message 'File under excluded folder should remain.'
        Assert-True -Condition (-not (Test-Path -LiteralPath $includedFile)) -Message 'Non-excluded old file should be deleted.'
    }
    finally {
        Remove-TestRoot -Path $root
    }
}

Invoke-TestCase -Name 'MaxFiles guard prevents unexpected large cleanup' -ScriptBlock {
    $root = New-TestRoot

    try {
        $firstFile = Join-Path $root 'first.log'
        $secondFile = Join-Path $root 'second.log'

        Set-Content -LiteralPath $firstFile -Value 'old'
        Set-Content -LiteralPath $secondFile -Value 'old'
        (Get-Item -LiteralPath $firstFile).LastWriteTime = (Get-Date).AddDays(-10)
        (Get-Item -LiteralPath $secondFile).LastWriteTime = (Get-Date).AddDays(-10)

        $result = Invoke-CleanupScript -ArgumentList @('-Path', $root, '-OlderThanDays', '1', '-MaxFiles', '1')

        Assert-True -Condition ($result.ExitCode -eq 1) -Message "Expected exit code 1, got $($result.ExitCode)."
        Assert-True -Condition (Test-Path -LiteralPath $firstFile) -Message 'First file should remain when MaxFiles guard trips.'
        Assert-True -Condition (Test-Path -LiteralPath $secondFile) -Message 'Second file should remain when MaxFiles guard trips.'
    }
    finally {
        Remove-TestRoot -Path $root
    }
}

Invoke-TestCase -Name 'Ownership takeover requires inheritance repair' -ScriptBlock {
    $root = New-TestRoot

    try {
        $oldFile = Join-Path $root 'old.log'
        Set-Content -LiteralPath $oldFile -Value 'old'
        (Get-Item -LiteralPath $oldFile).LastWriteTime = (Get-Date).AddDays(-10)

        $result = Invoke-CleanupScript -ArgumentList @('-Path', $root, '-OlderThanDays', '1', '-TakeOwnershipOnFailure')

        Assert-True -Condition ($result.ExitCode -eq 1) -Message "Expected exit code 1, got $($result.ExitCode)."
        Assert-True -Condition (Test-Path -LiteralPath $oldFile) -Message 'File should remain when parameter validation fails.'
    }
    finally {
        Remove-TestRoot -Path $root
    }
}

Invoke-TestCase -Name 'SMB share roots are refused before scanning' -ScriptBlock {
    $result = Invoke-CleanupScript -ArgumentList @('-Path', '\\example-server\example-share', '-OlderThanDays', '1', '-WhatIf')
    $hasDangerousRootWarning = @($result.Output | Where-Object { $_ -match 'Refusing to scan dangerous root path' }).Count -gt 0

    Assert-True -Condition ($result.ExitCode -eq 1) -Message "Expected exit code 1, got $($result.ExitCode)."
    Assert-True -Condition $hasDangerousRootWarning -Message 'Expected dangerous root path warning.'
}

Invoke-TestCase -Name 'RemoveEmptyDirectories deletes only empty child folders' -ScriptBlock {
    $root = New-TestRoot

    try {
        $emptyFolder = Join-Path $root 'empty'
        $nonEmptyFolder = Join-Path $root 'non-empty'
        $newFile = Join-Path $nonEmptyFolder 'new.log'

        New-Item -Path $emptyFolder -ItemType Directory -Force | Out-Null
        New-Item -Path $nonEmptyFolder -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath $newFile -Value 'new'

        $result = Invoke-CleanupScript -ArgumentList @('-Path', $root, '-OlderThanDays', '1', '-RemoveEmptyDirectories')

        Assert-True -Condition ($result.ExitCode -eq 0) -Message "Expected exit code 0, got $($result.ExitCode)."
        Assert-True -Condition (-not (Test-Path -LiteralPath $emptyFolder)) -Message 'Empty folder should be deleted.'
        Assert-True -Condition (Test-Path -LiteralPath $nonEmptyFolder) -Message 'Non-empty folder should remain.'
        Assert-True -Condition (Test-Path -LiteralPath $newFile) -Message 'New file should remain.'
    }
    finally {
        Remove-TestRoot -Path $root
    }
}

Write-Host '--------------------------------------------------'
Write-Host "Tests: $script:TestCount | Failed: $script:FailedCount"
Write-Host '--------------------------------------------------'

if ($script:FailedCount -gt 0) {
    exit 1
}

exit 0
