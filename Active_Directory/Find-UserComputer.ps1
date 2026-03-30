<#
.SYNOPSIS
    Finds which computer(s) a target user is logged into across AD.
.DESCRIPTION
    Queries AD for recently-active computers, then scans them in parallel
    using quser to locate a specific user's session(s).
.PARAMETER Username
    The SAMAccountName to search for (required).
.PARAMETER DaysInactive
    How far back to consider a computer "recently active". Default: 30.
.PARAMETER ThrottleLimit
    Max concurrent threads for the parallel scan. Default: 50.
.PARAMETER PingTimeoutSeconds
    Seconds to wait for each ping reply. Default: 1.
.EXAMPLE
    .\Find-UserComputer.ps1 -Username "JDoe"
.EXAMPLE
    .\Find-UserComputer.ps1 -Username "JDoe" -DaysInactive 15 -ThrottleLimit 100
.EXAMPLE
    # Pipe results to CSV
    .\Find-UserComputer.ps1 -Username "JDoe" | Export-Csv -Path .\results.csv -NoTypeInformation
#>

#requires -Version 7.0
#requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0,
        HelpMessage = "SAMAccountName of the user to locate.")]
    [ValidateNotNullOrEmpty()]
    [string]$Username,

    [ValidateRange(1, 365)]
    [int]$DaysInactive = 30,

    [ValidateRange(1, 500)]
    [int]$ThrottleLimit = 50,

    [ValidateRange(1, 5)]
    [int]$PingTimeoutSeconds = 1
)

# ── 1. Validate the user exists in AD before wasting time scanning ──
try {
    $ADUser = Get-ADUser -Identity $Username -ErrorAction Stop
    Write-Host "Confirmed AD user: $($ADUser.SamAccountName) ($($ADUser.Name))" -ForegroundColor Cyan
}
catch {
    Write-Error "User '$Username' not found in Active Directory. Verify the SAMAccountName."
    exit 1
}

# ── 2. Pull recently-active computer list ──
$DateCutoff = (Get-Date).AddDays(-$DaysInactive).ToFileTime()

Write-Host "Retrieving computers active in the last $DaysInactive days..." -ForegroundColor Cyan

$Computers = Get-ADComputer `
    -Filter "Enabled -eq 'true' -and LastLogonTimestamp -gt $DateCutoff" `
    -Properties Name |
    Select-Object -ExpandProperty Name

if (-not $Computers) {
    Write-Warning "No active computers found. Try increasing -DaysInactive."
    exit 0
}

Write-Host "Scanning $($Computers.Count) computers (ThrottleLimit: $ThrottleLimit)..." -ForegroundColor Cyan

# ── 3. Parallel scan ──
$Results = $Computers | ForEach-Object -Parallel {

    $Computer       = $_
    $User           = $using:Username
    $PingTimeout    = $using:PingTimeoutSeconds

    # Fast-fail: skip unreachable machines
    if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet -TimeoutSeconds $PingTimeout)) {
        return   # next iteration — cleaner than deep nesting
    }

    try {
        $SessionLines = quser /server:$Computer 2>&1

        # quser returns strings; filter to only the line(s) matching our user
        $MatchedLines = $SessionLines | Where-Object { $_ -match "\b$User\b" }

        foreach ($Line in $MatchedLines) {
            # ── Parse quser's fixed-width columns ──
            # Two formats depending on whether SESSIONNAME is present:
            #   USER  SESSION  ID  STATE  IDLE   LOGON
            #   USER           ID  STATE  IDLE   LOGON   (disconnected — no session name)
            $Session  = 'N/A'
            $State    = 'Unknown'
            $IdleTime = 'Unknown'
            $Logon    = 'Unknown'

            if ($Line -match '^\s*>?(\S+)\s+(console|rdp-tcp\S*)\s+(\d+)\s+(\S+)\s+(\S+)\s+(.+)$') {
                $Session  = $Matches[2]
                $State    = $Matches[4]
                $IdleTime = $Matches[5]
                $Logon    = $Matches[6].Trim()
            }
            elseif ($Line -match '^\s*>?(\S+)\s+(\d+)\s+(Disc)\s+(\S+)\s+(.+)$') {
                $State    = 'Disconnected'
                $IdleTime = $Matches[4]
                $Logon    = $Matches[5].Trim()
            }

            # Emit a structured object (NOT Write-Host) so output is pipeline-friendly
            [PSCustomObject]@{
                Computer  = $Computer
                Username  = $User
                Session   = $Session
                State     = $State
                IdleTime  = $IdleTime
                LogonTime = $Logon
            }
        }
    }
    catch {
        # RPC unavailable / Access Denied — silently skip
    }

} -ThrottleLimit $ThrottleLimit

# ── 4. Report ──
if ($Results) {
    Write-Host "`n✅ '$Username' found on $($Results.Count) machine(s):" -ForegroundColor Green
    $Results | Format-Table -AutoSize
}
else {
    Write-Host "`n❌ '$Username' was not found on any of the $($Computers.Count) scanned computers." -ForegroundColor Yellow
}

Write-Host "Scan complete." -ForegroundColor Cyan

# Pass objects to the pipeline so callers can pipe to Export-Csv, etc.
return $Results