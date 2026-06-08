# Remove Aged Files

Reusable PowerShell cleanup script for scheduled removal of old files. It is built for Windows administration workflows such as PDQ Deploy jobs, SMB archive folders, application export folders, and log retention tasks.

The script is intentionally conservative:

- It never deletes the configured root folders.
- It refuses drive roots such as `C:\` and SMB share roots such as `\\server\share`.
- ACL repair is disabled unless `-RepairInheritance` is supplied.
- Ownership takeover is disabled unless `-TakeOwnershipOnFailure` is supplied.
- Ownership takeover requires `-RepairInheritance`.
- `-WhatIf` is supported for preview runs.
- A failed delete, ACL repair, or ownership takeover returns exit code `1`.

## Basic Usage

Preview a cleanup without changing anything:

```powershell
.\Remove-AgedFiles.ps1 -Path "\\server\share\Archive" -OlderThanDays 30 -WhatIf
```

Delete files older than 30 days:

```powershell
.\Remove-AgedFiles.ps1 -Path "\\server\share\Archive" -OlderThanDays 30
```

Delete old files and remove empty subfolders:

```powershell
.\Remove-AgedFiles.ps1 -Path "\\server\share\Archive" -OlderThanDays 30 -RemoveEmptyDirectories
```

## Fax Cleanup Example

This example mirrors the TrueNAS/PDQ fax cleanup scenario. Normal deletion is attempted first. If deletion fails, the script tries to enable inherited permissions. If that still does not work, it takes ownership, enables inherited permissions again, and retries deletion.

```powershell
.\Remove-AgedFiles.ps1 `
    -Path "\\nas01\Operations\Faxes\Archived", "\\nas01\Operations\Faxes\Sent" `
    -OlderThanDays 30 `
    -IncludeExtension .pdf `
    -RemoveEmptyDirectories `
    -RepairInheritance `
    -TakeOwnershipOnFailure
```

## Parameters

| Parameter | Purpose |
| --- | --- |
| `-Path` | One or more root folders to scan. |
| `-OlderThanDays` | Retention period. Files older than this are candidates. Default: `30`. |
| `-RemoveEmptyDirectories` | Removes empty subfolders after old files are processed. |
| `-RepairInheritance` | On deletion failure, runs `icacls /inheritancelevel:e` and retries. |
| `-TakeOwnershipOnFailure` | On continued failure, runs `takeown`, repairs inheritance again, and retries. |
| `-IncludeExtension` | Optional extension filter, such as `.pdf`, `.log`, or `.bak`. |
| `-ExcludePath` | Optional list of files or folders to skip. |
| `-MaxFiles` | Maximum aged files processed per root path. Default: `10000`. |
| `-PassThru` | Emits structured result objects after processing. |
| `-WhatIf` | Shows intended destructive actions without changing anything. |
| `-Confirm` | Uses PowerShell's confirmation behavior for destructive actions. |

When launching through `powershell.exe -File` on Windows PowerShell 5.1, avoid passing `-Confirm:$false` on the command line. The script does not prompt by default, and Windows PowerShell 5.1 has awkward switch-parameter parsing for explicit `$false` values in `-File` calls.

## Exit Codes

| Exit Code | Meaning |
| --- | --- |
| `0` | Cleanup completed successfully. |
| `1` | One or more paths, deletes, ACL repairs, or ownership takeovers failed. |

This makes the script suitable for PDQ Deploy, scheduled tasks, and other automation tools that rely on process exit codes.

## Permission Repair Notes

`-RepairInheritance` uses:

```powershell
icacls.exe <path> /inheritancelevel:e /Q
```

This enables inherited permissions from the parent folder. It does not remove existing explicit permissions from the file.

`-TakeOwnershipOnFailure` uses:

```powershell
takeown.exe /F <path>
```

Taking ownership by itself does not grant delete permissions. The script enables inheritance again after ownership takeover so the file can inherit permissions from the parent folder.

For TrueNAS SMB shares, the account running the script may need to be in the SMB service's configured administrators group before `takeown.exe` can succeed.

## Testing

Run the included validation harness:

```powershell
.\Tests\Test-RemoveAgedFiles.ps1
```

The test script creates temporary folders under the current user's temp path, invokes `Remove-AgedFiles.ps1` in child PowerShell processes, and checks common safety behavior.
