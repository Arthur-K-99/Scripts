<#
.SYNOPSIS
    Takes ownership of a folder tree and grants the running user Full Control over it.

.DESCRIPTION
    This script forcefully seizes ownership of one or more target folders and every file and
    subfolder beneath them, then grants Full Control to the account that runs the script (or to
    an alternate identity supplied with -Identity). It is intended for recovering access to
    directory trees whose permissions have been locked down, orphaned, or corrupted - for
    example, folders left behind by a deleted user account, profile directories, or data
    inherited from a decommissioned server where the current administrator has no rights at all.

    Recovery is performed in three stages, in the order that survives even a completely
    inaccessible ACL:

        1. takeown.exe seizes ownership recursively. Becoming the owner grants the implicit
           READ_CONTROL and WRITE_DAC rights needed to rewrite the permissions in the next
           stages, even when the current DACL denies everything.
        2. icacls /setowner sets the final owner to the requested identity. This is normally
           the running user and therefore redundant, but it matters when -Identity names a
           different principal.
        3. icacls /grant applies an inheritable Full Control entry so the identity has full
           access to the tree and to anything created in it later.

    Because seizing ownership requires the SeTakeOwnershipPrivilege, the script must be run
    from an elevated (Run as administrator) session. It is enforced with #requires.

    This is a deliberately destructive operation: it changes the owner and permissions of every
    item in the tree and cannot be automatically undone. The script therefore supports -WhatIf
    and -Confirm and prompts before acting unless -Force is supplied. Existing permission
    entries are preserved (the grant is additive) unless -Reset is specified, which first
    removes explicit ACLs and restores inheritance before the Full Control grant is applied.

.PARAMETER Path
    One or more folders to take ownership of. Accepts pipeline input. Each path must exist and
    resolve to a directory.

.PARAMETER Identity
    The account to make the owner and grant Full Control to. Defaults to the user running the
    script. Accepts any form icacls understands, such as "DOMAIN\User", "MACHINE\User",
    "user@domain.com", or a well-known name like "Administrators".

.PARAMETER NoRecurse
    Limits the operation to the top-level folder only. By default the script applies ownership
    and the permission grant to all files and subfolders in the tree.

.PARAMETER Reset
    Removes explicit ACL entries and restores inheritance on the tree before granting Full
    Control. Use this to clear out stale or conflicting permissions rather than adding to them.

.PARAMETER Force
    Suppresses the confirmation prompt and performs the change. Equivalent to -Confirm:$false.

.PARAMETER YesToken
    The affirmative response character that takeown.exe expects for its /D prompt on folders
    where the current user lacks list permission. This is localized: 'Y' on English Windows,
    but a different letter on other language installations (for example 'J' on German,
    'O' on French). Defaults to 'Y'.

.INPUTS
    System.String. Folder paths can be piped in.

.OUTPUTS
    PSCustomObject, one per path, reporting the resulting Owner, Identity, whether each stage
    succeeded, the number of items icacls failed to process, and any error text.

.EXAMPLE
    .\Grant-FolderFullControl.ps1 -Path "D:\Data\OldProfile"

    Takes ownership of D:\Data\OldProfile and everything under it and grants the running
    administrator Full Control, after a confirmation prompt.

.EXAMPLE
    .\Grant-FolderFullControl.ps1 -Path "E:\Share\Finance" -Force

    Recovers the folder tree without prompting - suitable for use inside another script.

.EXAMPLE
    .\Grant-FolderFullControl.ps1 -Path "C:\Users\jdoe" -Identity "CONTOSO\Domain Admins" -Reset -Force

    Clears the existing permissions on the orphaned profile, restores inheritance, and grants
    Domain Admins Full Control as the new owner.

.EXAMPLE
    Get-ChildItem D:\Departed -Directory | .\Grant-FolderFullControl.ps1 -Force

    Takes ownership of every immediate subfolder of D:\Departed and grants the running user
    Full Control on each.

.EXAMPLE
    .\Grant-FolderFullControl.ps1 -Path "F:\Archive" -WhatIf

    Shows what would happen without changing anything.

.NOTES
    Run from an elevated PowerShell session. The script wraps the built-in takeown.exe and
    icacls.exe tools, so it works on Windows PowerShell 5.1 and PowerShell 7+.
#>

#requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('FullName', 'PSPath')]
    [ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [ValidateNotNullOrEmpty()]
    [string]$Identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,

    [switch]$NoRecurse,

    [switch]$Reset,

    [switch]$Force,

    [ValidatePattern('^\S$')]
    [string]$YesToken = 'Y'
)

begin {
    # Recursive by default; -NoRecurse limits the operation to the top-level folder only.
    $Recurse = -not $NoRecurse

    # Runs a native console tool and returns its stdout/stderr and exit code without letting
    # a non-zero exit terminate the script. icacls uses /C to continue past individual items,
    # so a non-zero code means "some items failed", not "nothing happened".
    function Invoke-NativeTool {
        param(
            [Parameter(Mandatory)][string]$FilePath,
            [Parameter(Mandatory)][string[]]$Arguments
        )

        Write-Verbose "$FilePath $($Arguments -join ' ')"
        $Output = & $FilePath @Arguments 2>&1
        [PSCustomObject]@{
            ExitCode = $LASTEXITCODE
            Output   = ($Output | Out-String).Trim()
        }
    }

    # icacls reports "Successfully processed N files; Failed processing M files" - pull M out
    # so callers can distinguish a clean run from a partial one.
    function Get-IcaclsFailureCount {
        param([string]$Output)
        if ($Output -match 'Failed processing (\d+)') { return [int]$Matches[1] }
        return 0
    }
}

process {
    foreach ($TargetPath in $Path) {
        $Result = [PSCustomObject]@{
            Path           = $TargetPath
            Identity       = $Identity
            Owner          = $null
            OwnershipTaken = $false
            OwnerSet       = $false
            Reset          = $false
            Granted        = $false
            FailedItems    = 0
            Status         = 'NotProcessed'
            Error          = $null
        }

        # Resolve and validate the target before touching anything.
        $Resolved = $null
        try {
            $Resolved = (Resolve-Path -LiteralPath $TargetPath -ErrorAction Stop).ProviderPath
        }
        catch {
            $Result.Status = 'NotFound'
            $Result.Error = "Path could not be resolved: $($_.Exception.Message)"
            Write-Error -Message $Result.Error -TargetObject $TargetPath
            $Result
            continue
        }

        if (-not (Test-Path -LiteralPath $Resolved -PathType Container)) {
            $Result.Status = 'NotADirectory'
            $Result.Error = "Path is not a directory: $Resolved"
            Write-Error -Message $Result.Error -TargetObject $TargetPath
            $Result
            continue
        }

        $Result.Path = $Resolved

        $Action = "Take ownership and grant '$Identity' Full Control" +
        $(if ($Recurse) { ' (recursive)' } else { '' }) +
        $(if ($Reset) { ', resetting existing permissions' } else { '' })
        if (-not ($Force -or $PSCmdlet.ShouldProcess($Resolved, $Action))) {
            $Result.Status = 'Skipped'
            $Result
            continue
        }

        try {
            # Stage 1: seize ownership. takeown assigns to the current user; /R recurses and
            # /D answers the localized default prompt on folders we cannot currently list.
            $TakeownArgs = @('/F', $Resolved)
            if ($Recurse) { $TakeownArgs += @('/R', '/D', $YesToken) }
            $Takeown = Invoke-NativeTool -FilePath 'takeown.exe' -Arguments $TakeownArgs
            $Result.OwnershipTaken = ($Takeown.ExitCode -eq 0)
            if (-not $Result.OwnershipTaken) {
                Write-Warning "takeown reported errors on '$Resolved' (exit $($Takeown.ExitCode)). Continuing."
                Write-Verbose $Takeown.Output
            }

            # Stage 2: set the final owner. Only current user or Administrators can be an owner
            # via takeown, so use icacls /setowner to honour a custom -Identity.
            $SetOwnerArgs = @($Resolved, '/setowner', $Identity, '/C', '/Q')
            if ($Recurse) { $SetOwnerArgs += '/T' }
            $SetOwner = Invoke-NativeTool -FilePath 'icacls.exe' -Arguments $SetOwnerArgs
            $Result.OwnerSet = ($SetOwner.ExitCode -eq 0)
            $Result.FailedItems += Get-IcaclsFailureCount -Output $SetOwner.Output

            # Optional stage: strip explicit ACLs and restore inheritance for a clean slate.
            if ($Reset) {
                $ResetArgs = @($Resolved, '/reset', '/C', '/Q')
                if ($Recurse) { $ResetArgs += '/T' }
                $ResetRun = Invoke-NativeTool -FilePath 'icacls.exe' -Arguments $ResetArgs
                $Result.Reset = ($ResetRun.ExitCode -eq 0)
                $Result.FailedItems += Get-IcaclsFailureCount -Output $ResetRun.Output
            }

            # Stage 3: grant inheritable Full Control. (OI)(CI) makes the entry propagate to
            # files and subfolders created in the tree from now on.
            $Grantee = if ($Recurse) { "${Identity}:(OI)(CI)F" } else { "${Identity}:F" }
            $GrantArgs = @($Resolved, '/grant', $Grantee, '/C', '/Q')
            if ($Recurse) { $GrantArgs += '/T' }
            $Grant = Invoke-NativeTool -FilePath 'icacls.exe' -Arguments $GrantArgs
            $Result.Granted = ($Grant.ExitCode -eq 0)
            $Result.FailedItems += Get-IcaclsFailureCount -Output $Grant.Output
            if (-not $Result.Granted) {
                Write-Warning "icacls /grant reported errors on '$Resolved' (exit $($Grant.ExitCode))."
                Write-Verbose $Grant.Output
            }

            # Read back the effective owner for confirmation in the output object.
            try {
                $Result.Owner = (Get-Acl -LiteralPath $Resolved -ErrorAction Stop).Owner
            }
            catch {
                Write-Verbose "Could not read back owner on '$Resolved': $($_.Exception.Message)"
            }

            $Result.Status = if ($Result.Granted -and $Result.FailedItems -eq 0) { 'Success' }
            elseif ($Result.Granted) { 'CompletedWithErrors' }
            else { 'Failed' }
        }
        catch {
            $Result.Status = 'Failed'
            $Result.Error = $_.Exception.Message
            Write-Error -Message $_.Exception.Message -TargetObject $Resolved
        }

        $Result
    }
}
