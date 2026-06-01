<#
.SYNOPSIS
    Pings enabled Active Directory computer objects in a specified Organizational Unit (OU).

.DESCRIPTION
    This script queries Active Directory for enabled computer objects within a specified OU,
    then tests whether each computer responds to a single ICMP ping. It also includes the
    computer object's approximate last AD logon date and password set date to help identify
    stale or obsolete computer objects.

    LastADLogonDate comes from Active Directory's replicated lastLogonTimestamp value, so it
    can lag behind the exact last logon time by several days. It is useful for stale computer
    cleanup, but it should not be treated as a real-time "last online" value.

.PARAMETER TargetOU
    The Distinguished Name (DN) of the Active Directory OU you want to query.

.PARAMETER TimeoutSeconds
    The number of seconds to wait for each ping response. Defaults to 1 second.
    Valid values are 1 through 10.

.PARAMETER StaleAfterDays
    The number of days after which an offline computer with no recent AD logon is flagged as
    possibly obsolete. Defaults to 90 days.

.EXAMPLE
    .\Ping-ComputersInOU.ps1 -TargetOU "OU=Desktops,DC=uhc-nyc,DC=org"

.EXAMPLE
    .\Ping-ComputersInOU.ps1 -TargetOU "OU=Workstations,DC=uhc-nyc,DC=org" -TimeoutSeconds 3

.EXAMPLE
    .\Ping-ComputersInOU.ps1 -TargetOU "OU=Workstations,DC=uhc-nyc,DC=org" -StaleAfterDays 120 |
        Where-Object PossiblyObsolete

.EXAMPLE
    .\Ping-ComputersInOU.ps1 -TargetOU "OU=Workstations,DC=uhc-nyc,DC=org" |
        Select-Object ComputerName, Status, IPAddress, LastADLogonDate, DaysSinceLastADLogon |
        Format-Table -AutoSize

.EXAMPLE
    .\Ping-ComputersInOU.ps1 -TargetOU "OU=Workstations,DC=uhc-nyc,DC=org" |
        Where-Object { $_.PossiblyObsolete } |
        Sort-Object DaysSinceLastADLogon -Descending |
        Format-Table ComputerName, Status, DaysSinceLastADLogon, LastADLogonDate -AutoSize

.EXAMPLE
    .\Ping-ComputersInOU.ps1 -TargetOU "OU=Workstations,DC=uhc-nyc,DC=org" |
        Where-Object { $_.Status -eq 'Offline' } |
        Format-List ComputerName, IPAddress, LastADLogonDate, PasswordLastSet, PossiblyObsolete

.EXAMPLE
    .\Ping-ComputersInOU.ps1 -TargetOU "OU=Workstations,DC=uhc-nyc,DC=org" |
        Out-GridView -Title "Computer Ping and Stale Object Audit"

.EXAMPLE
    .\Ping-ComputersInOU.ps1 -TargetOU "OU=Workstations,DC=uhc-nyc,DC=org" |
        Select-Object ComputerName, Status, LastADLogonDate, DaysSinceLastADLogon, PasswordLastSet |
        Export-Csv .\ComputerPingAudit.csv -NoTypeInformation
#>

#requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetOU,

    [ValidateRange(1, 10)]
    [int]$TimeoutSeconds = 1,

    [ValidateRange(1, 3650)]
    [int]$StaleAfterDays = 90
)

Write-Verbose "Fetching enabled computers from: $TargetOU"

try {
    $Now = Get-Date
    $StaleCutoff = $Now.AddDays(-$StaleAfterDays)

    $Computers = Get-ADComputer `
        -Filter "Enabled -eq 'true'" `
        -SearchBase $TargetOU `
        -Properties IPv4Address, DNSHostName, LastLogonDate, PasswordLastSet `
        -ErrorAction Stop

    if (-not $Computers) {
        Write-Warning "No enabled computers found in OU '$TargetOU'."
        return
    }

    foreach ($Computer in $Computers) {
        $Name = if ($Computer.DNSHostName) {
            $Computer.DNSHostName
        }
        else {
            $Computer.Name
        }

        Write-Verbose "Pinging $Name..."

        try {
            $Online = Test-Connection `
                -ComputerName $Name `
                -Count 1 `
                -Quiet `
                -TimeoutSeconds $TimeoutSeconds `
                -ErrorAction Stop
        }
        catch {
            $Online = $false
        }

        $LastADLogonDate = $Computer.LastLogonDate
        $DaysSinceLastADLogon = if ($LastADLogonDate) {
            [int]($Now - $LastADLogonDate).TotalDays
        }
        else {
            $null
        }

        $ADLogonStale = if ($LastADLogonDate) {
            $LastADLogonDate -lt $StaleCutoff
        }
        else {
            $true
        }

        [PSCustomObject]@{
            ComputerName         = $Name
            Status               = if ($Online) { 'Online' } else { 'Offline' }
            IPAddress            = $Computer.IPv4Address
            LastADLogonDate      = $LastADLogonDate
            DaysSinceLastADLogon = $DaysSinceLastADLogon
            PasswordLastSet      = $Computer.PasswordLastSet
            ADLogonStale         = $ADLogonStale
            PossiblyObsolete     = (-not $Online) -and $ADLogonStale
        }
    }
}
catch {
    Write-Error "Failed to query Active Directory for OU '$TargetOU': $($_.Exception.Message)"
    exit 1
}
