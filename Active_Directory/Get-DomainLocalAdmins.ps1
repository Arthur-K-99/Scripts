<#
.SYNOPSIS
    Reports the members of the local Administrators group on Active Directory computers.

.DESCRIPTION
    This script queries Active Directory for computer objects and remotely enumerates the
    membership of each machine's built-in local Administrators group. It returns one row per
    group member, reporting the member's name, the domain/machine it belongs to, whether it
    is a user or a group, whether the principal is local or an Active Directory object, its
    SID, and (for local accounts) whether the account is enabled.

    Group membership by itself does not reveal whether an account is enabled, so for local
    user accounts the enabled state is resolved from Get-LocalUser on the target. The Enabled
    column is left blank for domain principals and groups, whose state is not stored locally.

    The Administrators group is located by its well-known SID (S-1-5-32-544) rather than by
    name, so the script works correctly on non-English installations of Windows where the
    group is named differently.

    Enumeration is performed through Invoke-Command (WinRM), so the querying account needs
    remote PowerShell access to the targets and WinRM must be enabled on them. On each target
    the script first tries Get-LocalGroupMember; if that fails (a known issue when the group
    contains orphaned SIDs from deleted accounts) it falls back to parsing "net localgroup".

    Computers that cannot be contacted are still represented in the output with a status of
    Unreachable, so the report distinguishes "no extra admins" from "host not reached".

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

.PARAMETER ExcludeBuiltIn
    Excludes the expected default members of the local Administrators group (the built-in
    local Administrator account and the Domain Admins group) so the output highlights only
    the principals that have been added on top of the defaults.

.PARAMETER SkipPing
    Skips the ICMP ping check. When specified, the script attempts to query all computers
    regardless of ping response.

.PARAMETER PingTimeoutSeconds
    Number of seconds to wait for each ping response. Defaults to 1 second.

.PARAMETER Credential
    Optional alternate credentials used for the remote enumeration.

.PARAMETER ThrottleLimit
    Maximum number of concurrent computer queries to run at once (PowerShell 7+ only).
    Defaults to 16.

.PARAMETER OutputPath
    Optional path to a CSV file. When supplied, results are written to the file (and its
    parent directory is created if needed) in addition to being returned to the pipeline.

.EXAMPLE
    .\Get-DomainLocalAdmins.ps1 | Format-Table ComputerName, MemberName, ObjectClass, PrincipalSource -AutoSize

    Lists every local Administrators member on every computer in the domain.

.EXAMPLE
    .\Get-DomainLocalAdmins.ps1 -SearchBase "OU=Workstations,DC=contoso,DC=com" -ExcludeBuiltIn `
        -OutputPath C:\Reports\LocalAdmins.csv

    Reports only the non-default local admins on workstations in the given OU and exports to CSV.

.EXAMPLE
    "PC-001", "PC-002" | .\Get-DomainLocalAdmins.ps1 | Where-Object ObjectClass -eq 'User'

    Audits specific computers and shows only user (not group) members.

.EXAMPLE
    .\Get-DomainLocalAdmins.ps1 -SearchBase "OU=Laptops,DC=contoso,DC=com" |
        Format-Table ComputerName, MemberName, ObjectClass, PrincipalSource, Enabled -AutoSize

    Lists local Administrators members and shows whether each local account is enabled. The
    Enabled column is blank for domain principals and groups (their state is held in AD, not
    on the machine).

.EXAMPLE
    .\Get-DomainLocalAdmins.ps1 -SearchBase "OU=Laptops,DC=contoso,DC=com" |
        Where-Object { $_.PrincipalSource -eq 'Local' -and $_.Enabled } |
        Format-Table ComputerName, MemberName, Enabled -AutoSize

    Shows only the enabled local admin accounts - the ones that are actually usable and
    therefore matter most for risk. Disabled accounts (such as the default built-in
    Administrator) and domain principals are filtered out.

.EXAMPLE
    .\Get-DomainLocalAdmins.ps1 -ExcludeBuiltIn |
        Group-Object MemberName |
        Sort-Object Count -Descending |
        Format-Table Name, Count -AutoSize

    Ranks which non-default principals appear as local admins on the most machines.
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

    [switch]$ExcludeBuiltIn,

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
    function Get-LocalAdminResult {
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

        if ($PingStatus -ne 'Online' -and -not $SkipPing) {
            return [PSCustomObject]@{
                ComputerName    = $Computer.Name
                Status          = 'Offline'
                MemberName      = $null
                MemberDomain    = $null
                ObjectClass     = $null
                PrincipalSource = $null
                Enabled         = $null
                SID             = $null
                Error           = "Host did not respond to ping ($PingStatus)"
            }
        }

        # Remote enumeration of the local Administrators group (well-known SID S-1-5-32-544).
        $InvokeParams = @{
            ComputerName = $Target
            ErrorAction  = 'Stop'
            ScriptBlock  = {
                $AdminsSid = 'S-1-5-32-544'

                # Build a lookup of local accounts and their enabled state. Group membership alone
                # does not reveal whether an account is enabled, so we resolve it from Get-LocalUser.
                # Keyed by both SID (used by the primary method) and name (used by the fallback).
                $LocalEnabledBySid = @{}
                $LocalEnabledByName = @{}
                try {
                    Get-LocalUser -ErrorAction Stop | ForEach-Object {
                        $LocalEnabledBySid[$_.SID.Value] = $_.Enabled
                        $LocalEnabledByName[$_.Name] = $_.Enabled
                    }
                }
                catch {
                    Write-Verbose "Get-LocalUser failed: $($_.Exception.Message). Enabled state for local accounts will be unknown."
                }

                # Primary method: Get-LocalGroupMember by SID (locale-independent).
                try {
                    Get-LocalGroupMember -SID $AdminsSid -ErrorAction Stop | ForEach-Object {
                        $Parts = $_.Name -split '\\', 2
                        if ($Parts.Count -eq 2) {
                            $Domain = $Parts[0]
                            $Account = $Parts[1]
                        }
                        else {
                            $Domain = $null
                            $Account = $_.Name
                        }
                        # Enabled state is only meaningful (and resolvable) for local accounts.
                        $Enabled = if ($_.PrincipalSource.ToString() -eq 'Local' -and
                            $LocalEnabledBySid.ContainsKey($_.SID.Value)) {
                            $LocalEnabledBySid[$_.SID.Value]
                        }
                        else { $null }
                        [PSCustomObject]@{
                            MemberName      = $Account
                            MemberDomain    = $Domain
                            ObjectClass     = $_.ObjectClass
                            PrincipalSource = $_.PrincipalSource.ToString()
                            SID             = $_.SID.Value
                            Enabled         = $Enabled
                            Method          = 'Get-LocalGroupMember'
                        }
                    }
                    return
                }
                catch {
                    Write-Verbose "Get-LocalGroupMember failed: $($_.Exception.Message). Falling back to net localgroup."
                }

                # Fallback: resolve the group's localized name from its SID, then parse net localgroup.
                $GroupName = (New-Object System.Security.Principal.SecurityIdentifier($AdminsSid)
                    ).Translate([System.Security.Principal.NTAccount]).Value -replace '^.*\\', ''

                $Lines = net localgroup $GroupName 2>$null
                # Members are listed between the dashed separator line and the "command completed" line.
                $Start = ($Lines | Select-String -SimpleMatch '----').LineNumber
                if (-not $Start) { return }

                $Lines | Select-Object -Skip $Start | Where-Object {
                    $_ -and $_ -notmatch 'The command completed'
                } | ForEach-Object {
                    $Entry = $_.Trim()
                    if (-not $Entry) { return }
                    $Parts = $Entry -split '\\', 2
                    if ($Parts.Count -eq 2) {
                        $Domain = $Parts[0]
                        $Account = $Parts[1]
                    }
                    else {
                        $Domain = $null
                        $Account = $Entry
                    }
                    # The local machine name as the "domain" indicates a local account; resolve its
                    # enabled state by name. Anything else is a domain principal (unknown here).
                    $IsLocal = (-not $Domain) -or ($Domain -eq $env:COMPUTERNAME)
                    $Enabled = if ($IsLocal -and $LocalEnabledByName.ContainsKey($Account)) {
                        $LocalEnabledByName[$Account]
                    }
                    else { $null }
                    [PSCustomObject]@{
                        MemberName      = $Account
                        MemberDomain    = $Domain
                        ObjectClass     = 'Unknown'
                        PrincipalSource = 'Unknown'
                        SID             = $null
                        Enabled         = $Enabled
                        Method          = 'net localgroup'
                    }
                }
            }
        }
        if ($Credential) { $InvokeParams.Credential = $Credential }

        try {
            $Members = Invoke-Command @InvokeParams

            if (-not $Members) {
                return [PSCustomObject]@{
                    ComputerName    = $Computer.Name
                    Status          = 'Online'
                    MemberName      = $null
                    MemberDomain    = $null
                    ObjectClass     = $null
                    PrincipalSource = $null
                    Enabled         = $null
                    SID             = $null
                    Error           = 'No members enumerated'
                }
            }

            foreach ($Member in $Members) {
                [PSCustomObject]@{
                    ComputerName    = $Computer.Name
                    Status          = 'Online'
                    MemberName      = $Member.MemberName
                    MemberDomain    = $Member.MemberDomain
                    ObjectClass     = $Member.ObjectClass
                    PrincipalSource = $Member.PrincipalSource
                    Enabled         = $Member.Enabled
                    SID             = $Member.SID
                    Error           = $null
                }
            }
        }
        catch {
            $Msg = $_.Exception.Message
            $Status = if ($Msg -match 'Access is denied' -or $Msg -match '401') { 'AccessDenied' } else { 'Unreachable' }
            [PSCustomObject]@{
                ComputerName    = $Computer.Name
                Status          = $Status
                MemberName      = $null
                MemberDomain    = $null
                ObjectClass     = $null
                PrincipalSource = $null
                Enabled         = $null
                SID             = $null
                Error           = $Msg
            }
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
                    -Properties DNSHostName, Enabled, OperatingSystem `
                    -ErrorAction Stop
                $script:ComputersToAudit.Add($AdComp)
            }
            catch {
                $script:ComputersToAudit.Add([PSCustomObject]@{
                        Name            = $Name
                        DNSHostName     = $Name
                        Enabled         = $true
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
            Properties  = @('DNSHostName', 'Enabled', 'OperatingSystem')
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

    Write-Verbose "Starting local administrators audit of $($script:ComputersToAudit.Count) computer(s)."

    $Results = if ($PSVersionTable.PSVersion.Major -ge 7) {
        $FunctionDef = ${function:Get-LocalAdminResult}.ToString()

        $script:ComputersToAudit | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            ${function:Get-LocalAdminResult} = $using:FunctionDef

            Get-LocalAdminResult `
                -Computer $_ `
                -SkipPing ([bool]$using:SkipPing) `
                -PingTimeoutSeconds $using:PingTimeoutSeconds `
                -Credential $using:Credential
        }
    }
    else {
        foreach ($Computer in $script:ComputersToAudit) {
            Get-LocalAdminResult `
                -Computer $Computer `
                -SkipPing ([bool]$SkipPing) `
                -PingTimeoutSeconds $PingTimeoutSeconds `
                -Credential $Credential
        }
    }

    if ($ExcludeBuiltIn) {
        $Results = $Results | Where-Object {
            # Always keep status rows (offline/unreachable/no members) so coverage is visible.
            if (-not $_.MemberName) { return $true }
            # Drop the built-in local Administrator (RID 500) and the Domain Admins group (RID 512).
            if ($_.SID -match '-500$' -or $_.SID -match '-512$') { return $false }
            if ($_.MemberName -eq 'Administrator' -or $_.MemberName -eq 'Domain Admins') { return $false }
            return $true
        }
    }

    $Results = $Results | Sort-Object ComputerName, ObjectClass, MemberName

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
