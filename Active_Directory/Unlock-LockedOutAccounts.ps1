#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Reports locked-out and expired-password Active Directory accounts, and interactively
    unlocks the locked-out ones.

.DESCRIPTION
    - Uses Search-ADAccount to find every account in the domain that is currently locked out.
    - Prompts whether to unlock them: all at once, one-by-one, or none.
    - Also reports accounts whose password has expired (Search-ADAccount -PasswordExpired).

    Expired passwords are report-only: there is no "unlock" equivalent — the user (or an
    admin reset) must set a new password — so the script surfaces them for follow-up rather
    than taking an action. Use -SkipPasswordCheck to report locked-out accounts only.

    By default the script is interactive. Use -Unlock to unlock every locked-out account
    without prompting, or use -WhatIf to preview which accounts would be unlocked.

.PARAMETER SearchBase
    Optional AD OU distinguished name to narrow the search scope. If omitted, the entire
    current domain is searched.

.PARAMETER Unlock
    Unlocks every locked-out account found, without prompting. Honors -WhatIf / -Confirm.

.PARAMETER SkipPasswordCheck
    Skips the expired-password report and only handles locked-out accounts.

.PARAMETER ReportPath
    Optional path to export a CSV report. Both locked-out and expired-password accounts are
    included, distinguished by a Category column.

.EXAMPLE
    # Show locked-out and expired-password accounts, and prompt for what to unlock
    .\Unlock-LockedOutAccounts.ps1

.EXAMPLE
    # Unlock everything that is locked out, no prompt
    .\Unlock-LockedOutAccounts.ps1 -Unlock

.EXAMPLE
    # Preview without changing anything
    .\Unlock-LockedOutAccounts.ps1 -Unlock -WhatIf

.EXAMPLE
    # Locked-out accounts only, no password report
    .\Unlock-LockedOutAccounts.ps1 -SkipPasswordCheck

.EXAMPLE
    # Scope to an OU and export a report
    .\Unlock-LockedOutAccounts.ps1 -SearchBase "OU=Employees,DC=contoso,DC=com" -ReportPath "C:\Reports\accounts.csv"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$SearchBase,

    [switch]$Unlock,

    [switch]$SkipPasswordCheck,

    [string]$ReportPath
)

# Common parameters for the Search-ADAccount calls below.
$searchBaseParam = @{}
if ($SearchBase) { $searchBaseParam['SearchBase'] = $SearchBase }

# ── 1. Find locked-out accounts ──────────────────────────────────────────────
Write-Host "🔍 Searching for locked-out accounts..." -ForegroundColor Cyan

try {
    $lockedAccounts = Search-ADAccount -LockedOut @searchBaseParam |
        Where-Object { $_.ObjectClass -eq 'user' } |
        Sort-Object SamAccountName
}
catch {
    Write-Error "Failed to query Active Directory: $($_.Exception.Message)"
    return
}

# ── 2. Build the locked-out report ───────────────────────────────────────────
$lockedResults = foreach ($account in $lockedAccounts) {
    $detail = $null
    try {
        $detail = Get-ADUser -Identity $account.SamAccountName `
            -Properties DisplayName, LastBadPasswordAttempt, lockoutTime, Enabled
    }
    catch {
        Write-Warning "Could not read details for $($account.SamAccountName): $_"
    }

    [PSCustomObject]@{
        Category               = 'LockedOut'
        SamAccountName         = $account.SamAccountName
        DisplayName            = $detail.DisplayName
        Enabled                = $detail.Enabled
        LastBadPasswordAttempt = $detail.LastBadPasswordAttempt
        LockoutTime            = if ($detail.lockoutTime) { [DateTime]::FromFileTime($detail.lockoutTime) } else { $null }
        PasswordExpiry         = $null
        DistinguishedName      = $account.DistinguishedName
        Action                 = 'Pending'
    }
}

if ($lockedResults) {
    Write-Host "⚠️  Found $($lockedResults.Count) locked-out account(s):`n" -ForegroundColor Yellow
    $lockedResults | Format-Table SamAccountName, DisplayName, Enabled, LockoutTime, LastBadPasswordAttempt -AutoSize
}
else {
    Write-Host "✅ No locked-out accounts found." -ForegroundColor Green
}

# ── 3. Decide which accounts to unlock ───────────────────────────────────────
# Build the set of accounts to unlock, either from -Unlock or an interactive prompt.
$toUnlock = @()

if ($lockedResults) {
    if ($Unlock) {
        $toUnlock = $lockedResults
    }
    else {
        Write-Host "Unlock options:" -ForegroundColor Cyan
        Write-Host "  [A] Unlock ALL listed accounts"
        Write-Host "  [S] Select accounts one-by-one"
        Write-Host "  [N] Do nothing (default)"
        $choice = (Read-Host "Your choice (A/S/N)").Trim().ToUpper()

        switch ($choice) {
            'A' { $toUnlock = $lockedResults }
            'S' {
                $toUnlock = foreach ($entry in $lockedResults) {
                    $answer = (Read-Host "Unlock $($entry.SamAccountName)? (y/N)").Trim().ToUpper()
                    if ($answer -eq 'Y') { $entry }
                }
            }
            default {
                Write-Host "ℹ️  No accounts unlocked." -ForegroundColor DarkYellow
            }
        }
    }
}

# ── 4. Unlock the chosen accounts ────────────────────────────────────────────
if ($toUnlock) {
    Write-Host "`n🔓 Unlocking $($toUnlock.Count) account(s)..." -ForegroundColor Cyan

    foreach ($entry in $toUnlock) {
        try {
            if ($PSCmdlet.ShouldProcess($entry.SamAccountName, "Unlock account")) {
                Unlock-ADAccount -Identity $entry.SamAccountName
                $entry.Action = 'Unlocked ✅'
                Write-Host "  ✅ $($entry.SamAccountName)" -ForegroundColor Green
            }
            else {
                $entry.Action = 'WhatIf'
            }
        }
        catch {
            $entry.Action = "Failed ❌: $_"
            Write-Warning "  ❌ $($entry.SamAccountName): $_"
        }
    }
}

# Mark any locked-out account not chosen (interactive 'S' mode or 'N') as skipped.
foreach ($entry in $lockedResults) {
    if ($entry.Action -eq 'Pending') { $entry.Action = 'Skipped' }
}

# ── 5. Find accounts with expired passwords (report-only) ────────────────────
$expiredResults = @()
if (-not $SkipPasswordCheck) {
    Write-Host "`n🔍 Searching for accounts with expired passwords..." -ForegroundColor Cyan

    try {
        $expiredAccounts = Search-ADAccount -PasswordExpired @searchBaseParam |
            Where-Object { $_.ObjectClass -eq 'user' -and $_.Enabled } |
            Sort-Object SamAccountName
    }
    catch {
        Write-Warning "Failed to query expired passwords: $($_.Exception.Message)"
        $expiredAccounts = @()
    }

    $expiredResults = foreach ($account in $expiredAccounts) {
        $detail = $null
        try {
            $detail = Get-ADUser -Identity $account.SamAccountName `
                -Properties DisplayName, PasswordLastSet, Enabled, 'msDS-UserPasswordExpiryTimeComputed'
        }
        catch {
            Write-Warning "Could not read details for $($account.SamAccountName): $_"
        }

        $expiry = $null
        $expiryRaw = $detail.'msDS-UserPasswordExpiryTimeComputed'
        # 0 = "must change at next logon"; 0x7FFFFFFFFFFFFFFF = "never expires".
        if ($expiryRaw -and $expiryRaw -ne 0 -and $expiryRaw -ne [Int64]::MaxValue) {
            $expiry = [DateTime]::FromFileTime($expiryRaw)
        }

        [PSCustomObject]@{
            Category               = 'PasswordExpired'
            SamAccountName         = $account.SamAccountName
            DisplayName            = $detail.DisplayName
            Enabled                = $detail.Enabled
            LastBadPasswordAttempt = $null
            LockoutTime            = $null
            PasswordExpiry         = $expiry
            DistinguishedName      = $account.DistinguishedName
            Action                 = 'Report only'
        }
    }

    if ($expiredResults) {
        Write-Host "⚠️  Found $($expiredResults.Count) enabled account(s) with an expired password:`n" -ForegroundColor Yellow
        $expiredResults | Format-Table SamAccountName, DisplayName, PasswordExpiry -AutoSize
        Write-Host "ℹ️  Expired passwords must be reset by the user or an admin; no action is taken here." -ForegroundColor DarkYellow
    }
    else {
        Write-Host "✅ No enabled accounts with expired passwords found." -ForegroundColor Green
    }
}

# ── 6. Export CSV report (optional) ──────────────────────────────────────────
$allResults = @($lockedResults) + @($expiredResults)

if ($ReportPath -and $allResults) {
    $allResults | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n📄 Report saved to: $ReportPath" -ForegroundColor Cyan
}

# ── 7. Summary ───────────────────────────────────────────────────────────────
if ($allResults) {
    Write-Host "`n── Summary ──" -ForegroundColor Magenta
    $allResults | Group-Object Category, Action | ForEach-Object {
        Write-Host "  $($_.Name): $($_.Count) account(s)"
    }
}
