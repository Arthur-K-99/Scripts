<#
.SYNOPSIS
    Audits WinRM (PowerShell Remoting) accessibility across Active Directory computers.

.DESCRIPTION
    This script queries Active Directory for computer objects and tests each one for WinRM
    connectivity. For every reachable computer it reports whether WinRM responds on the
    standard HTTP port (5985) and the HTTPS port (5986), and runs a brief remote session
    to collect the PowerShell version and WinRM service state.

    Computers that do not respond to ping are marked Offline. Computers that respond to
    ping but refuse WinRM connections are marked Unreachable and the error is recorded.
    Computers where the remote session is established but the caller lacks permissions are
    marked AccessDenied.

    Under PowerShell 7 or later, computers are audited in parallel (see ThrottleLimit).
    Under Windows PowerShell 5.1, computers are audited sequentially.

.PARAMETER SearchBase
    Optional Distinguished Name (DN) of an OU to scan. If omitted, the entire current
    Active Directory domain is scanned.

.PARAMETER IncludeDisabled
    Includes disabled computer objects. Disabled objects are excluded by default.

.PARAMETER ComputerName
    Audits specific computer name(s) directly, bypassing the Active Directory domain-wide
    search. Supports pipeline input.

.PARAMETER OperatingSystemFilter
    Optional wildcard pattern matched against each computer object's operatingSystem
    attribute (for example '*Windows 10*' or '*Server*'). If omitted, all computer objects
    are targeted.

.PARAMETER SkipPing
    Skips the ICMP ping check. When specified, the script attempts WinRM on all computers.

.PARAMETER PingTimeoutSeconds
    Number of seconds to wait for each ping response. Defaults to 1 second.

.PARAMETER Credential
    Optional alternate credentials used for both Test-WSMan and the remote session.

.PARAMETER ThrottleLimit
    Maximum number of concurrent computer audits to run at once (PowerShell 7+ only).
    Defaults to 16.

.PARAMETER OutputPath
    Optional path to a CSV file. When supplied, results are written to the file (and its
    parent directory is created if needed) in addition to being returned to the pipeline.

.EXAMPLE
    .\Get-DomainWinRMAudit.ps1 | Where-Object WinRMStatus -ne 'Accessible' | Format-Table ComputerName, WinRMStatus, ErrorDetails -AutoSize

    Finds all computers in the domain that are not reachable via WinRM.

.EXAMPLE
    .\Get-DomainWinRMAudit.ps1 -SearchBase "OU=Workstations,DC=contoso,DC=com" -OutputPath C:\Reports\WinRMAudit.csv

    Audits all computers in the Workstations OU and exports the results to CSV.

.EXAMPLE
    "PC-001", "PC-002" | .\Get-DomainWinRMAudit.ps1 -SkipPing | Select-Object ComputerName, WinRMStatus, PSVersion

    Audits specific computers, skipping the ping check.

.EXAMPLE
    .\Get-DomainWinRMAudit.ps1 -OperatingSystemFilter '*Server*' | Group-Object WinRMStatus | Format-Table Name, Count -AutoSize

    Summarises WinRM accessibility across all domain servers.
#>

#requires -Modules ActiveDirectory

[CmdletBinding(DefaultParameterSetName = 'ADSearch')]
param(
    [Parameter(ParameterSetName = 'ADSearch', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,

    [Parameter(ParameterSetName = 'ADSearch')]
    [switch]$IncludeDisabled,

    [Parameter(ParameterSetName = 'ManualList', Mandatory, ValueFromPipeline)]
    [string[]]$ComputerName,

    [Parameter(ParameterSetName = 'ADSearch')]
    [string]$OperatingSystemFilter,

    [switch]$SkipPing,

    [ValidateRange(1, 10)]
    [int]$PingTimeoutSeconds = 1,

    [System.Management.Automation.PSCredential]$Credential,

    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 16,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    function Test-TcpPort {
        param([string]$HostName, [int]$Port, [int]$TimeoutMs = 2000)
        $Tcp = [System.Net.Sockets.TcpClient]::new()
        try {
            $Connect = $Tcp.BeginConnect($HostName, $Port, $null, $null)
            $Connected = $Connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
            if ($Connected -and $Tcp.Connected) { 'Open' } else { 'Closed' }
        }
        catch { 'Closed' }
        finally { $Tcp.Dispose() }
    }

    function Get-WinRMAuditResult {
        param(
            [Parameter(Mandatory)]
            $Computer,

            [bool]$SkipPing,
            [int]$PingTimeoutSeconds,
            [System.Management.Automation.PSCredential]$Credential
        )

        $Target = if ($Computer.DNSHostName) { $Computer.DNSHostName } else { $Computer.Name }

        Write-Verbose "Auditing $Target"

        # Ping check
        $PingStatus = if ($SkipPing) {
            'NotTested'
        }
        else {
            $Pinger = [System.Net.NetworkInformation.Ping]::new()
            try {
                $Reply = $Pinger.Send($Target, $PingTimeoutSeconds * 1000)
                if ($Reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                    'Online'
                }
                else { 'NoReply' }
            }
            catch {
                $Inner = $_.Exception
                while ($Inner.InnerException) { $Inner = $Inner.InnerException }
                if ($Inner -is [System.Net.Sockets.SocketException] -and
                    $Inner.SocketErrorCode -eq [System.Net.Sockets.SocketError]::HostNotFound) {
                    'NameResolutionFailed'
                }
                else { 'NoReply' }
            }
            finally { $Pinger.Dispose() }
        }

        $WinRMStatus = 'Unknown'
        $WinRMHTTP = 'Unknown'
        $WinRMHTTPS = 'Unknown'
        $PSVersion = 'Unknown'
        $ServiceState = 'Unknown'
        $ErrorDetails = ''

        if ($PingStatus -eq 'Online' -or $SkipPing) {
            # Port reachability
            $WinRMHTTP = Test-TcpPort -HostName $Target -Port 5985
            $WinRMHTTPS = Test-TcpPort -HostName $Target -Port 5986

            # WinRM protocol test
            $WSManParams = @{ ComputerName = $Target; ErrorAction = 'Stop' }
            if ($Credential) { $WSManParams.Credential = $Credential }

            try {
                $null = Test-WSMan @WSManParams
            }
            catch {
                $Msg = $_.Exception.Message
                if ($Msg -match 'Access is denied' -or $Msg -match '401') {
                    $WinRMStatus = 'AccessDenied'
                    $ErrorDetails = $Msg
                }
                else {
                    $WinRMStatus = 'Unreachable'
                    $ErrorDetails = $Msg
                }
            }

            # Remote session for extra detail
            if (-not $ErrorDetails) {
                $InvokeParams = @{
                    ComputerName = $Target
                    ErrorAction  = 'Stop'
                    ScriptBlock  = {
                        [PSCustomObject]@{
                            PSVersion    = $PSVersionTable.PSVersion.ToString()
                            ServiceState = (Get-Service -Name WinRM -ErrorAction SilentlyContinue).Status
                        }
                    }
                }
                if ($Credential) { $InvokeParams.Credential = $Credential }

                try {
                    $Remote = Invoke-Command @InvokeParams
                    $PSVersion = $Remote.PSVersion
                    $ServiceState = $Remote.ServiceState
                    $WinRMStatus = 'Accessible'
                }
                catch {
                    $Msg = $_.Exception.Message
                    if ($Msg -match 'Access is denied' -or $Msg -match '401') {
                        $WinRMStatus = 'AccessDenied'
                    }
                    else {
                        $WinRMStatus = 'Unreachable'
                    }
                    $ErrorDetails = $Msg
                }
            }
        }
        else {
            $WinRMStatus = 'Offline'
            $WinRMHTTP = 'N/A'
            $WinRMHTTPS = 'N/A'
            $PSVersion = 'N/A'
            $ServiceState = 'N/A'
        }

        [PSCustomObject]@{
            ComputerName      = $Computer.Name
            DNSHostName       = $Computer.DNSHostName
            Enabled           = $Computer.Enabled
            IPAddress         = $Computer.IPv4Address
            PingStatus        = $PingStatus
            ADOperatingSystem = $Computer.OperatingSystem
            WinRMStatus       = $WinRMStatus
            WinRMHTTP         = $WinRMHTTP
            WinRMHTTPS        = $WinRMHTTPS
            PSVersion         = $PSVersion
            ServiceState      = $ServiceState
            ErrorDetails      = $ErrorDetails
        }
    }

    $script:ComputersToAudit = [System.Collections.Generic.List[PSCustomObject]]::new()
}

process {
    if ($PSBoundParameters.ContainsKey('ComputerName')) {
        foreach ($Name in $ComputerName) {
            if (-not $Name) { continue }
            try {
                $AdComp = Get-ADComputer -Identity $Name `
                    -Properties DNSHostName, Enabled, IPv4Address, OperatingSystem `
                    -ErrorAction Stop
                $script:ComputersToAudit.Add($AdComp)
            }
            catch {
                $script:ComputersToAudit.Add([PSCustomObject]@{
                        Name            = $Name
                        DNSHostName     = $Name
                        Enabled         = $true
                        IPv4Address     = $null
                        OperatingSystem = 'Unknown'
                    })
            }
        }
    }
}

end {
    if ($PSCmdlet.ParameterSetName -eq 'ADSearch') {
        $ADQueryParameters = @{
            Filter      = if ($OperatingSystemFilter) {
                "OperatingSystem -like '$OperatingSystemFilter'"
            }
            else { 'Name -like "*"' }
            Properties  = @('DNSHostName', 'Enabled', 'IPv4Address', 'OperatingSystem')
            ErrorAction = 'Stop'
        }

        if ($SearchBase) { $ADQueryParameters.SearchBase = $SearchBase }

        Write-Verbose 'Querying Active Directory for computer objects.'
        try {
            $AdComputers = Get-ADComputer @ADQueryParameters
            foreach ($Comp in $AdComputers) {
                if ($IncludeDisabled -or $Comp.Enabled) {
                    $script:ComputersToAudit.Add($Comp)
                }
            }
        }
        catch {
            Write-Error "Failed to query Active Directory: $($_.Exception.Message)"
            return
        }
    }

    if ($script:ComputersToAudit.Count -eq 0) {
        Write-Warning 'No computer objects were found to audit.'
        return
    }

    Write-Verbose "Starting WinRM audit of $($script:ComputersToAudit.Count) computer(s)."

    $Results = if ($PSVersionTable.PSVersion.Major -ge 7) {
        $FunctionDef = ${function:Get-WinRMAuditResult}.ToString()
        $TcpFunctionDef = ${function:Test-TcpPort}.ToString()

        $script:ComputersToAudit | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            ${function:Test-TcpPort} = $using:TcpFunctionDef
            ${function:Get-WinRMAuditResult} = $using:FunctionDef

            Get-WinRMAuditResult `
                -Computer $_ `
                -SkipPing ([bool]$using:SkipPing) `
                -PingTimeoutSeconds $using:PingTimeoutSeconds `
                -Credential $using:Credential
        }
    }
    else {
        foreach ($Computer in $script:ComputersToAudit) {
            Get-WinRMAuditResult `
                -Computer $Computer `
                -SkipPing ([bool]$SkipPing) `
                -PingTimeoutSeconds $PingTimeoutSeconds `
                -Credential $Credential
        }
    }

    $Results = $Results | Sort-Object ComputerName

    if ($OutputPath) {
        $ParentDir = Split-Path -Path $OutputPath -Parent
        if ($ParentDir -and -not (Test-Path -Path $ParentDir)) {
            New-Item -ItemType Directory -Path $ParentDir -Force | Out-Null
        }
        $Results | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Verbose "Results written to $OutputPath"
    }

    $Results
}
