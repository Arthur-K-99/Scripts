#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Finds active AD users expiring soon and optionally extends their expiration by one year.

.DESCRIPTION
    - Queries AD for enabled user accounts with an expiration date within a specified window.
    - Outputs a report of soon-to-expire accounts.
    - Optionally extends each account's expiration date by 1 year.

.PARAMETER DaysUntilExpiry
    Number of days from today to look ahead for expiring accounts. Default: 30

.PARAMETER ExtendExpiration
    If specified, extends matching accounts' expiration by 1 year.

.PARAMETER ReportPath
    Optional path to export a CSV report of affected users.

.PARAMETER SearchBase
    Optional AD OU distinguished name to narrow the search scope.

.EXAMPLE
    # Report only (no changes)
    .\Extend-ExpiringUsers.ps1 -DaysUntilExpiry 60

    # Extend expiration dates
    .\Extend-ExpiringUsers.ps1 -DaysUntilExpiry 30 -ExtendExpiration

    # With CSV export and scoped OU
    .\Extend-ExpiringUsers.ps1 -DaysUntilExpiry 30 -ExtendExpiration -ReportPath "C:\Reports\expiring.csv" -SearchBase "OU=Employees,DC=contoso,DC=com"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [int]$DaysUntilExpiry = 30,

    [switch]$ExtendExpiration,

    [string]$ReportPath,

    [string]$SearchBase
)

# ── 1. Build the search filter ──────────────────────────────────────────────
$today = (Get-Date).Date
$horizon = $today.AddDays($DaysUntilExpiry)

Write-Host "🔍 Searching for enabled accounts expiring between $($today.ToString('yyyy-MM-dd')) and $($horizon.ToString('yyyy-MM-dd'))..." -ForegroundColor Cyan

# Common parameters for Get-ADUser
$adParams = @{
    Filter     = 'Enabled -eq $true -and AccountExpirationDate -like "*"'
    Properties = @('AccountExpirationDate', 'DisplayName', 'EmailAddress',
        'Description', 'Title', 'Manager', 'SamAccountName')
}
if ($SearchBase) { $adParams['SearchBase'] = $SearchBase }

# ── 2. Retrieve and filter users ────────────────────────────────────────────
$expiringUsers = Get-ADUser @adParams | Where-Object {
    $_.AccountExpirationDate -and
    $_.AccountExpirationDate -ge $today -and
    $_.AccountExpirationDate -le $horizon
} | Sort-Object AccountExpirationDate

if (-not $expiringUsers) {
    Write-Host "✅ No active accounts found expiring within the next $DaysUntilExpiry days." -ForegroundColor Green
    return
}

Write-Host "⚠️  Found $($expiringUsers.Count) account(s) expiring soon:`n" -ForegroundColor Yellow

# ── 3. Build a results collection ───────────────────────────────────────────
$results = foreach ($user in $expiringUsers) {
    $currentExpiry = $user.AccountExpirationDate
    $newExpiry = $currentExpiry.AddYears(1)
    $daysLeft = ($currentExpiry - $today).Days

    [PSCustomObject]@{
        SamAccountName  = $user.SamAccountName
        DisplayName     = $user.DisplayName
        Email           = $user.EmailAddress
        Description     = $user.Description
        Title           = $user.Title
        CurrentExpiry   = $currentExpiry.ToString('yyyy-MM-dd')
        DaysRemaining   = $daysLeft
        NewExpiry       = $newExpiry.ToString('yyyy-MM-dd')
        ExtensionStatus = 'Pending'
    }
}

# Display the table
$results | Format-Table SamAccountName, DisplayName, Description,
CurrentExpiry, DaysRemaining, NewExpiry -AutoSize

# ── 4. Optionally extend expiration dates ───────────────────────────────────
if ($ExtendExpiration) {
    Write-Host "`n🔄 Extending expiration dates by 1 year..." -ForegroundColor Cyan

    foreach ($entry in $results) {
        $newDate = [DateTime]::ParseExact($entry.NewExpiry, 'yyyy-MM-dd', $null)

        try {
            if ($PSCmdlet.ShouldProcess($entry.SamAccountName,
                    "Extend expiration to $($entry.NewExpiry)")) {

                Set-ADAccountExpiration -Identity $entry.SamAccountName -DateTime $newDate
                $entry.ExtensionStatus = 'Extended ✅'
                Write-Host "  ✅ $($entry.SamAccountName) → $($entry.NewExpiry)" -ForegroundColor Green
            }
        }
        catch {
            $entry.ExtensionStatus = "Failed ❌: $_"
            Write-Warning "  ❌ $($entry.SamAccountName): $_"
        }
    }
}
else {
    Write-Host "ℹ️  Run with -ExtendExpiration to actually extend these accounts." -ForegroundColor DarkYellow
    foreach ($entry in $results) { $entry.ExtensionStatus = 'Report Only' }
}

# ── 5. Export CSV report (optional) ─────────────────────────────────────────
if ($ReportPath) {
    $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n📄 Report saved to: $ReportPath" -ForegroundColor Cyan
}

# ── 6. Summary ──────────────────────────────────────────────────────────────
Write-Host "`n── Summary ──" -ForegroundColor Magenta
$results | Group-Object ExtensionStatus | ForEach-Object {
    Write-Host "  $($_.Name): $($_.Count) account(s)"
}