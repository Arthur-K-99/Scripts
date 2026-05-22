<#
.SYNOPSIS
    Audits Active Directory computer objects for their chassis type and hidden battery/UPS states.

.DESCRIPTION
    This script queries Active Directory for active computer objects within a specified Organizational Unit (OU). 
    It utilizes concurrent connections (multi-threading) to rapidly ping and query each endpoint via CIM/WMI. 
    This quickly identifies desktops that Windows is treating as battery-powered devices (often due to a connected UPS), 
    which would cause them to bypass strictly "Plugged In" GPO power policies.

.PARAMETER SearchBase
    The Distinguished Name (DN) of the Active Directory OU you want to query.

.PARAMETER ThrottleLimit
    The maximum number of concurrent connections to run at once. Defaults to 20. 
    Increase this for larger OUs if your network and CPU can handle it.

.EXAMPLE
    .\Get-ChassisBatteryAudit.ps1 -SearchBase "OU=Desktops,DC=uhc-nyc,DC=org"

.EXAMPLE
    .\Get-ChassisBatteryAudit.ps1 -SearchBase "OU=Workstations,DC=uhc-nyc,DC=org" -ThrottleLimit 50
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "The Distinguished Name of the OU to scan.")]
    [string]$SearchBase,

    [Parameter(Mandatory = $false, HelpMessage = "Maximum number of concurrent connections.")]
    [int]$ThrottleLimit = 15 # Lowered slightly to respect default WinRM concurrency limits
)

#Requires -Version 7.0

Write-Host "Querying Active Directory for active computers in: $SearchBase" -ForegroundColor Cyan
try {
    $Computers = Get-ADComputer -SearchBase $SearchBase -Filter { Enabled -eq $true } -ErrorAction Stop
}
catch {
    Write-Error "Failed to query AD. Ensure the RSAT ActiveDirectory module is installed and the SearchBase is correct."
    return
}

if (-not $Computers) {
    Write-Warning "No active computers found in the specified SearchBase."
    return
}

Write-Host "Found $($Computers.Count) computers. Starting native CIM audit..." -ForegroundColor Cyan

$Results = $Computers | ForEach-Object -Parallel {
    $PC = $_.Name
    
    try {
        # Using the native command to avoid CIM Session negotiation quotas
        # If no battery exists, it returns null without throwing an error.
        $Battery = Get-CimInstance -ClassName Win32_Battery -ComputerName $PC -ErrorAction Stop
        $HasBattery = if ($null -ne $Battery) { $true } else { $false }

        # Query Chassis Class
        $Enclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ComputerName $PC -ErrorAction Stop
        $Chassis = $Enclosure.ChassisTypes -join ','

        [PSCustomObject]@{
            ComputerName = $PC
            Status       = "Success"
            HasBattery   = $HasBattery
            ChassisType  = $Chassis
            ErrorDetails = ""
        }
    }
    catch {
        [PSCustomObject]@{
            ComputerName = $PC
            Status       = "Failed"
            HasBattery   = "Unknown"
            ChassisType  = "Unknown"
            ErrorDetails = $_.Exception.Message
        }
    }
} -ThrottleLimit $ThrottleLimit

$Results | Out-GridView -Title "Chassis & Battery Audit Results"