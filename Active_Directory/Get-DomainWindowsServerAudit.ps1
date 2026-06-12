<#
.SYNOPSIS
    Audits Windows Server computer objects across Active Directory.

.DESCRIPTION
    This script queries Active Directory for Windows Server computer objects and returns
    their operating system, version, enabled state, approximate last AD logon date, ping
    status, and Windows Server extended-support date.

    By default, the script scans the entire current domain and uses Active Directory data.
    Use QueryLiveDetails to also query each server through CIM for its current OS caption,
    version, build number, and last boot time. Live queries connect over WSMan first and
    fall back to DCOM if the WSMan connection fails.

    Under PowerShell 7 or later, servers are audited in parallel (see ThrottleLimit) and
    output order is not guaranteed; pipe to Sort-Object if you need a stable order. Under
    Windows PowerShell 5.1, servers are audited sequentially.

    LastADLogonDate comes from Active Directory's replicated lastLogonTimestamp value, so it
    can lag behind the exact last logon time by several days. A failed ping does not prove a
    server is offline because ICMP might be blocked. A PingStatus of NameResolutionFailed
    usually indicates a stale computer object whose DNS record no longer exists.

.PARAMETER SearchBase
    Optional Distinguished Name (DN) of an OU to scan. If omitted, the entire current
    Active Directory domain is scanned.

.PARAMETER IncludeDisabled
    Includes disabled Windows Server computer objects. Disabled objects are excluded by
    default.

.PARAMETER SkipPing
    Skips the ICMP ping test. PingStatus is reported as NotTested.

.PARAMETER PingTimeoutSeconds
    Number of seconds to wait for each ping response. Defaults to 1 second.

.PARAMETER QueryLiveDetails
    Queries each server through CIM for its current OS caption, version, build number,
    and last boot time. The querying account must have remote-management access.

.PARAMETER CimTimeoutSeconds
    Number of seconds to wait for each live CIM operation. Defaults to 10 seconds.

.PARAMETER ThrottleLimit
    Maximum number of servers audited at the same time when running under PowerShell 7 or
    later. Defaults to 16. Ignored under Windows PowerShell 5.1.

.EXAMPLE
    .\Get-DomainWindowsServerAudit.ps1

    Scans enabled Windows Server computer objects across the current domain.

.EXAMPLE
    .\Get-DomainWindowsServerAudit.ps1 -IncludeDisabled |
        Sort-Object ExtendedSupportEnd, ComputerName |
        Format-Table ComputerName, Enabled, ADOperatingSystem, ADOperatingSystemVersion,
            PingStatus, ExtendedSupportEnd, SupportStatus -AutoSize

    Includes disabled objects and displays a lifecycle-focused report.

.EXAMPLE
    .\Get-DomainWindowsServerAudit.ps1 -QueryLiveDetails |
        Where-Object LiveQueryStatus -eq 'Success' |
        Format-Table ComputerName, LiveCaption, LiveVersion, LiveBuild,
            LiveLastBootTime -AutoSize

    Queries live OS details and displays servers that responded successfully.

.EXAMPLE
    .\Get-DomainWindowsServerAudit.ps1 -SearchBase "OU=Servers,DC=uhc-nyc,DC=org" |
        Out-GridView -Title "Windows Server Domain Audit"

    Scans only the specified OU and displays the results in a searchable grid.

.EXAMPLE
    .\Get-DomainWindowsServerAudit.ps1 -IncludeDisabled -QueryLiveDetails |
        Export-Csv .\WindowsServerDomainAudit.csv -NoTypeInformation

    Exports a complete server inventory to CSV.
#>

#requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,

    [switch]$IncludeDisabled,

    [switch]$SkipPing,

    [ValidateRange(1, 10)]
    [int]$PingTimeoutSeconds = 1,

    [switch]$QueryLiveDetails,

    [ValidateRange(1, 120)]
    [int]$CimTimeoutSeconds = 10,

    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 16
)

function Get-ServerLifecycle {
    param(
        [string]$OperatingSystem
    )

    $ExtendedSupportEnd = switch -Regex ($OperatingSystem) {
        'Windows Server 2025'    { [datetime]'2034-10-10'; break }
        'Windows Server 2022'    { [datetime]'2031-10-14'; break }
        'Windows Server 2019'    { [datetime]'2029-01-09'; break }
        'Windows Server 2016'    { [datetime]'2027-01-12'; break }
        'Windows Server 2012 R2' { [datetime]'2023-10-10'; break }
        'Windows Server 2012'    { [datetime]'2023-10-10'; break }
        'Windows Server 2008 R2' { [datetime]'2020-01-14'; break }
        'Windows Server 2008'    { [datetime]'2020-01-14'; break }
        default                  { $null }
    }

    $SupportStatus = if (-not $ExtendedSupportEnd) {
        'Unknown'
    }
    elseif ((Get-Date).Date -le $ExtendedSupportEnd) {
        'Supported'
    }
    else {
        'EndOfSupport'
    }

    [PSCustomObject]@{
        ExtendedSupportEnd = $ExtendedSupportEnd
        SupportStatus      = $SupportStatus
    }
}

function Get-ServerAuditResult {
    param(
        [Parameter(Mandatory)]
        $Server,

        [bool]$SkipPing,
        [int]$PingTimeoutSeconds,
        [bool]$QueryLiveDetails,
        [int]$CimTimeoutSeconds,
        [datetime]$Now
    )

    $Target = if ($Server.DNSHostName) {
        $Server.DNSHostName
    }
    else {
        $Server.Name
    }

    Write-Verbose "Auditing $Target"

    $PingStatus = if ($SkipPing) {
        'NotTested'
    }
    else {
        $Pinger = New-Object System.Net.NetworkInformation.Ping

        try {
            $PingReply = $Pinger.Send($Target, ($PingTimeoutSeconds * 1000))

            if ($PingReply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                'Online'
            }
            else {
                'NoReply'
            }
        }
        catch {
            $InnerException = $_.Exception
            while ($InnerException.InnerException) {
                $InnerException = $InnerException.InnerException
            }

            if ($InnerException -is [System.Net.Sockets.SocketException] -and
                $InnerException.SocketErrorCode -eq [System.Net.Sockets.SocketError]::HostNotFound) {
                'NameResolutionFailed'
            }
            else {
                'NoReply'
            }
        }
        finally {
            $Pinger.Dispose()
        }
    }

    $LiveCaption = $null
    $LiveVersion = $null
    $LiveBuild = $null
    $LiveLastBootTime = $null
    $LiveQueryStatus = if ($QueryLiveDetails) { 'Failed' } else { 'NotRequested' }
    $LiveQueryError = $null

    if ($QueryLiveDetails) {
        $CimSession = $null

        try {
            try {
                $CimSession = New-CimSession `
                    -ComputerName $Target `
                    -OperationTimeoutSec $CimTimeoutSeconds `
                    -ErrorAction Stop
            }
            catch {
                $DcomSessionOption = New-CimSessionOption -Protocol Dcom
                $CimSession = New-CimSession `
                    -ComputerName $Target `
                    -SessionOption $DcomSessionOption `
                    -OperationTimeoutSec $CimTimeoutSeconds `
                    -ErrorAction Stop
            }

            $LiveOS = Get-CimInstance `
                -ClassName Win32_OperatingSystem `
                -CimSession $CimSession `
                -OperationTimeoutSec $CimTimeoutSeconds `
                -ErrorAction Stop

            $LiveCaption = $LiveOS.Caption
            $LiveVersion = $LiveOS.Version
            $LiveBuild = $LiveOS.BuildNumber
            $LiveLastBootTime = $LiveOS.LastBootUpTime
            $LiveQueryStatus = 'Success'
        }
        catch {
            $LiveQueryError = $_.Exception.Message
        }
        finally {
            if ($CimSession) {
                Remove-CimSession -CimSession $CimSession -ErrorAction SilentlyContinue
            }
        }
    }

    $LifecycleOS = if ($LiveCaption) {
        $LiveCaption
    }
    else {
        $Server.OperatingSystem
    }

    $Lifecycle = Get-ServerLifecycle -OperatingSystem $LifecycleOS
    $LastADLogonDate = $Server.LastLogonDate
    $DaysSinceLastADLogon = if ($LastADLogonDate) {
        [int]($Now - $LastADLogonDate).TotalDays
    }
    else {
        $null
    }

    [PSCustomObject]@{
        ComputerName             = $Server.Name
        DNSHostName              = $Server.DNSHostName
        Enabled                  = $Server.Enabled
        IPAddress                = $Server.IPv4Address
        PingStatus               = $PingStatus
        ADOperatingSystem        = $Server.OperatingSystem
        ADOperatingSystemVersion = $Server.OperatingSystemVersion
        ADServicePack            = $Server.OperatingSystemServicePack
        ExtendedSupportEnd       = $Lifecycle.ExtendedSupportEnd
        SupportStatus            = $Lifecycle.SupportStatus
        LastADLogonDate          = $LastADLogonDate
        DaysSinceLastADLogon     = $DaysSinceLastADLogon
        PasswordLastSet          = $Server.PasswordLastSet
        LiveCaption              = $LiveCaption
        LiveVersion              = $LiveVersion
        LiveBuild                = $LiveBuild
        LiveLastBootTime         = $LiveLastBootTime
        LiveQueryStatus          = $LiveQueryStatus
        LiveQueryError           = $LiveQueryError
        DistinguishedName        = $Server.DistinguishedName
    }
}

$ADQueryParameters = @{
    Filter      = "OperatingSystem -like '*Windows Server*'"
    Properties  = @(
        'DNSHostName'
        'Enabled'
        'IPv4Address'
        'LastLogonDate'
        'OperatingSystem'
        'OperatingSystemServicePack'
        'OperatingSystemVersion'
        'PasswordLastSet'
    )
    ErrorAction = 'Stop'
}

if ($SearchBase) {
    $ADQueryParameters.SearchBase = $SearchBase
}

Write-Verbose "Querying Active Directory for Windows Server computer objects."

try {
    $Servers = Get-ADComputer @ADQueryParameters |
        Where-Object { $IncludeDisabled -or $_.Enabled } |
        Sort-Object Name
}
catch {
    Write-Error "Failed to query Active Directory: $($_.Exception.Message)"
    return
}

if (-not $Servers) {
    Write-Warning 'No matching Windows Server computer objects were found.'
    return
}

$Now = Get-Date

if ($PSVersionTable.PSVersion.Major -ge 7) {
    $GetServerLifecycleDefinition = ${function:Get-ServerLifecycle}.ToString()
    $GetServerAuditResultDefinition = ${function:Get-ServerAuditResult}.ToString()

    $Servers | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
        ${function:Get-ServerLifecycle} = $using:GetServerLifecycleDefinition
        ${function:Get-ServerAuditResult} = $using:GetServerAuditResultDefinition

        Get-ServerAuditResult `
            -Server $_ `
            -SkipPing ([bool]$using:SkipPing) `
            -PingTimeoutSeconds $using:PingTimeoutSeconds `
            -QueryLiveDetails ([bool]$using:QueryLiveDetails) `
            -CimTimeoutSeconds $using:CimTimeoutSeconds `
            -Now $using:Now
    }
}
else {
    foreach ($Server in $Servers) {
        Get-ServerAuditResult `
            -Server $Server `
            -SkipPing ([bool]$SkipPing) `
            -PingTimeoutSeconds $PingTimeoutSeconds `
            -QueryLiveDetails ([bool]$QueryLiveDetails) `
            -CimTimeoutSeconds $CimTimeoutSeconds `
            -Now $Now
    }
}
