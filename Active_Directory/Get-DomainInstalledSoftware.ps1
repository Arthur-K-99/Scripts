<#
.SYNOPSIS
    Inventories machine-wide (admin-installed) software across Active Directory computers.
    Per-user installations are not detected.

.DESCRIPTION
    This script queries Active Directory for computer objects and remotely reads each
    machine's "Uninstall" registry keys to build an inventory of installed software. It
    reports the display name, version, publisher, install date, and source registry scope
    for every product found, and records an entry for each computer that could not be
    reached so the report distinguishes "no matching software" from "host unreachable".

    Both the native 64-bit and the 32-bit (WOW6432Node) uninstall hives are read, so 32-bit
    products on 64-bit systems are not missed. Entries without a DisplayName (typically
    update or component fragments) are skipped.

    NOTE: Only machine-wide (admin / "for all users") installations are detected. These live
    under the HKEY_LOCAL_MACHINE uninstall hives. Per-user installations performed by a
    standard user (which register under that user's HKEY_CURRENT_USER hive, for example
    user-scoped Chrome, Teams, or Zoom) are NOT detected, because this script reads only the
    HKLM hives and does not load remote per-user profile hives.

    Use NameFilter to limit the inventory to products whose display name matches a wildcard
    pattern (for example '*Adobe*'). Omit it to inventory everything.

    Remote queries run through Invoke-Command (WinRM), so the querying account needs remote
    PowerShell access to the targets and WinRM must be enabled on them. Targets are queried
    in parallel by Invoke-Command's built-in fan-out.

.PARAMETER NameFilter
    Optional wildcard pattern matched against each product's DisplayName (for example
    '*Adobe*' or 'Google Chrome'). If omitted, all installed products are returned.

.PARAMETER SearchBase
    Optional Distinguished Name (DN) of an OU to scan. If omitted, the entire current
    Active Directory domain is scanned.

.PARAMETER IncludeDisabled
    Includes disabled computer objects. Disabled objects are excluded by default.

.PARAMETER OperatingSystemFilter
    Optional wildcard pattern matched against each computer object's operatingSystem
    attribute (for example '*Windows 10*' or '*Server*'). If omitted, all computer objects
    are targeted.

.PARAMETER OutputPath
    Optional path to a CSV file. When supplied, results are written to the file (and its
    parent directory is created if needed) in addition to being returned to the pipeline.

.PARAMETER ThrottleLimit
    Maximum number of computers Invoke-Command contacts at the same time. Defaults to 32.

.EXAMPLE
    .\Get-DomainInstalledSoftware.ps1 -NameFilter '*Adobe*' |
        Sort-Object ComputerName, DisplayName |
        Format-Table ComputerName, DisplayName, DisplayVersion, InstallDate -AutoSize

    Inventories every Adobe product across the domain and displays a sorted table.

.EXAMPLE
    .\Get-DomainInstalledSoftware.ps1 -SearchBase "OU=Workstations,DC=contoso,DC=com" `
        -OutputPath C:\Reports\WorkstationSoftware.csv

    Inventories all software on workstations in the specified OU and exports it to CSV.

.EXAMPLE
    .\Get-DomainInstalledSoftware.ps1 -NameFilter '*Chrome*' -OperatingSystemFilter '*Windows 10*' |
        Group-Object DisplayVersion |
        Sort-Object Name |
        Format-Table Name, Count -AutoSize

    Shows how many Windows 10 machines run each version of Google Chrome.
#>

#requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [string]$NameFilter,

    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,

    [switch]$IncludeDisabled,

    [string]$OperatingSystemFilter,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [ValidateRange(1, 128)]
    [int]$ThrottleLimit = 32
)

$ADQueryParameters = @{
    Filter      = 'Name -like "*"'
    Properties  = @('Enabled', 'OperatingSystem')
    ErrorAction = 'Stop'
}

if ($OperatingSystemFilter) {
    $ADQueryParameters.Filter = "OperatingSystem -like '$OperatingSystemFilter'"
}

if ($SearchBase) {
    $ADQueryParameters.SearchBase = $SearchBase
}

Write-Verbose 'Querying Active Directory for computer objects.'

try {
    $Computers = Get-ADComputer @ADQueryParameters |
        Where-Object { $IncludeDisabled -or $_.Enabled } |
        Sort-Object Name
}
catch {
    Write-Error "Failed to query Active Directory: $($_.Exception.Message)"
    return
}

if (-not $Computers) {
    Write-Warning 'No matching computer objects were found.'
    return
}

$ComputerNames = $Computers | Select-Object -ExpandProperty Name

Write-Verbose "Querying $($ComputerNames.Count) computer(s) for installed software."

$ErrorList = @()
$Found = Invoke-Command -ComputerName $ComputerNames -ThrottleLimit $ThrottleLimit `
    -ErrorVariable ErrorList -ErrorAction SilentlyContinue -ArgumentList $NameFilter -ScriptBlock {
        param($NameFilter)

        $Scopes = @(
            [PSCustomObject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'HKLM64' }
            [PSCustomObject]@{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'HKLM32' }
        )

        foreach ($Scope in $Scopes) {
            Get-ItemProperty -Path $Scope.Path -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.DisplayName -and (-not $NameFilter -or $_.DisplayName -like $NameFilter)
                } |
                ForEach-Object {
                    [PSCustomObject]@{
                        DisplayName    = $_.DisplayName
                        DisplayVersion = $_.DisplayVersion
                        Publisher      = $_.Publisher
                        InstallDate    = $_.InstallDate
                        RegistryScope  = $Scope.Scope
                    }
                }
        }
    }

$Results = foreach ($Item in $Found) {
    [PSCustomObject]@{
        ComputerName   = $Item.PSComputerName
        Status         = 'Online'
        DisplayName    = $Item.DisplayName
        DisplayVersion = $Item.DisplayVersion
        Publisher      = $Item.Publisher
        InstallDate    = $Item.InstallDate
        RegistryScope  = $Item.RegistryScope
        Error          = $null
    }
}

$RespondedComputers = $Found.PSComputerName | Sort-Object -Unique
$FailedComputers = $ComputerNames | Where-Object { $_ -notin $RespondedComputers }

$FailedResults = foreach ($Computer in $FailedComputers) {
    $ErrorMessage = $ErrorList |
        Where-Object { $_.TargetObject -eq $Computer } |
        Select-Object -ExpandProperty Exception -First 1

    [PSCustomObject]@{
        ComputerName   = $Computer
        Status         = 'Unreachable'
        DisplayName    = $null
        DisplayVersion = $null
        Publisher      = $null
        InstallDate    = $null
        RegistryScope  = $null
        Error          = if ($ErrorMessage) { $ErrorMessage.Message } else { 'Connection failed / WinRM unreachable' }
    }
}

$AllResults = @($Results) + @($FailedResults) | Sort-Object ComputerName, DisplayName

if ($OutputPath) {
    $ParentDir = Split-Path -Path $OutputPath -Parent
    if ($ParentDir -and -not (Test-Path -Path $ParentDir)) {
        New-Item -ItemType Directory -Path $ParentDir -Force | Out-Null
    }

    $AllResults | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Verbose "Results written to $OutputPath"
}

$AllResults
