[CmdletBinding()]
param (
    [string]$OutputPath = "C:\temp\DomainPowerStateAudit.csv",
    [string]$SearchBase # Optional: Scope to a specific OU
)

# 1. Ensure Output Path directory exists
$parentDir = Split-Path $OutputPath
if (-not (Test-Path $parentDir)) {
    New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
}

# 2. Fetch active Windows workstations from Active Directory (filtering out disabled accounts)
$filter = "Enabled -eq '$true' -and (OperatingSystem -like '*Windows 10*' -or OperatingSystem -like '*Windows 11*')"
$adParams = @{
    Filter = $filter
}
if ($SearchBase) { $adParams.SearchBase = $SearchBase }

Write-Host "Fetching computer list from Active Directory..." -ForegroundColor Cyan
$computers = Get-ADComputer @adParams | Select-Object -ExpandProperty Name

if ($computers.Count -eq 0) {
    Write-Warning "No active Windows 10/11 computers found."
    return
}

Write-Host "Auditing $($computers.Count) computers in parallel..." -ForegroundColor Cyan

# 3. Query the computers remotely
$errorList = @()
$powerAudit = Invoke-Command -ComputerName $computers -ErrorVariable errorList -ErrorAction SilentlyContinue -ScriptBlock {
    # Capture powercfg output as a single multiline string
    $powercfgOutput = (powercfg /a) -join "`n"

    $delimiter = "The following sleep states are not available on this system"
    # Use -split operator instead of .Split() method to avoid splitting by individual characters
    $availableSection = ($powercfgOutput -split $delimiter)[0]

    # Evaluate against the available section (case-insensitive regex matches)
    $supportsS3 = [bool]($availableSection -match "Standby \(S3\)")
    $supportsModernStandby = [bool]($availableSection -match "Standby \(S0 Low Power Idle\)")
    $supportsHibernate = [bool]($availableSection -match "Hibernate")

    [PSCustomObject]@{
        ComputerName     = $env:COMPUTERNAME
        Status           = "Online"
        S3_LegacySleep   = $supportsS3
        S0_ModernStandby = $supportsModernStandby
        Hibernate        = $supportsHibernate
        Error            = $null
    }
}

# 4. Identify and append unreachable hosts to the report
$respondedComputers = $powerAudit.ComputerName
$failedComputers = $computers | Where-Object { $_ -notin $respondedComputers }

$failedResults = foreach ($comp in $failedComputers) {
    # Find matching error message
    $errMessage = $errorList | Where-Object { $_.TargetObject -eq $comp } | Select-Object -ExpandProperty Exception -First 1
    if (-not $errMessage) { $errMessage = "Connection failed / WinRM unreachable" }

    [PSCustomObject]@{
        ComputerName     = $comp
        Status           = "Offline/Error"
        S3_LegacySleep   = $false
        S0_ModernStandby = $false
        Hibernate        = $false
        Error            = $errMessage.Message
    }
}

# Combine and Export Results
$finalResults = $powerAudit + $failedResults
$finalResults | Export-Csv -Path $OutputPath -NoTypeInformation

Write-Host "Audit complete. Results saved to $OutputPath" -ForegroundColor Green