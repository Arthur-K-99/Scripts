<#
.SYNOPSIS
    Audits Active Directory computer objects for UEFI Secure Boot status and the presence of 2023 certificates.

.DESCRIPTION
    This script queries Active Directory for Windows computer objects (clients and servers) and uses 
    PowerShell Remoting to audit their UEFI Secure Boot status, checking specifically for the new 2023 
    certificates required to replace the expiring 2011 certificate chain.

    It validates:
    1. If the computer supports UEFI boot mode and has Secure Boot enabled.
    2. If the Allowed Signature Database (db) contains the 'Windows UEFI CA 2023' certificate.
    3. If the Key Exchange Key (KEK) contains the 'Microsoft Corporation KEK 2K CA 2023' certificate.
    4. The registry status of the Secure Boot Servicing update (UEFICA2023Status).

    By default, the script scans the current Active Directory domain. You can specify a target OU 
    using SearchBase, or run it against a list of specific computers using the ComputerName parameter.

    Under PowerShell 7 or later, computers are audited in parallel (see ThrottleLimit). Under 
    Windows PowerShell 5.1, computers are audited sequentially.

.PARAMETER SearchBase
    Optional Distinguished Name (DN) of an OU to scan. If omitted, the entire current 
    Active Directory domain is scanned.

.PARAMETER IncludeDisabled
    Includes disabled Windows computer objects. Disabled objects are excluded by default.

.PARAMETER ComputerName
    Audits specific computer name(s) directly, bypassing the Active Directory domain-wide search. 
    Supports pipeline input.

.PARAMETER SkipPing
    Skips the ICMP ping check. When specified, the script attempts remote query on all target computers.

.PARAMETER PingTimeoutSeconds
    Number of seconds to wait for each ping response. Defaults to 1 second.

.PARAMETER Credential
    Optional alternate credentials to run the remote query on target computers.

.PARAMETER ThrottleLimit
    Maximum number of concurrent computer connections to run at once (PowerShell 7+ only). Defaults to 16.

.EXAMPLE
    .\Get-DomainSecureBootAudit.ps1

    Audits all enabled Windows computers in the current Active Directory domain.

.EXAMPLE
    .\Get-DomainSecureBootAudit.ps1 -SearchBase "OU=Workstations,DC=uhc-nyc,DC=org" | Out-GridView -Title "Secure Boot Audit"

    Audits computers under the specified OU and displays the results in a searchable grid.

.EXAMPLE
    "PC-001", "PC-002" | .\Get-DomainSecureBootAudit.ps1 -SkipPing | Export-Csv .\SecureBootAudit.csv -NoTypeInformation

    Audits specific computers directly, skipping ping checks, and outputs results to a CSV file.
#>

#requires -Modules ActiveDirectory

[CmdletBinding(DefaultParameterSetName = "ADSearch")]
param(
    [Parameter(ParameterSetName = "ADSearch", Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,

    [Parameter(ParameterSetName = "ADSearch")]
    [switch]$IncludeDisabled,

    [Parameter(ParameterSetName = "ManualList", Mandatory = $true, ValueFromPipeline = $true)]
    [string[]]$ComputerName,

    [switch]$SkipPing,

    [ValidateRange(1, 10)]
    [int]$PingTimeoutSeconds = 1,

    [System.Management.Automation.PSCredential]$Credential,

    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 16
)

begin {
    function Get-SecureBootAuditResult {
        param(
            [Parameter(Mandatory)]
            $Computer,

            [bool]$SkipPing,
            [int]$PingTimeoutSeconds,
            [System.Management.Automation.PSCredential]$Credential
        )

        $Target = if ($Computer.DNSHostName) {
            $Computer.DNSHostName
        }
        else {
            $Computer.Name
        }

        Write-Verbose "Auditing $Target"

        # 1. Ping Check
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

        # 2. Remote Audit Check
        $BootMode = "Unknown"
        $SecureBootEnabled = "Unknown"
        $DbHas2023Cert = "Unknown"
        $KekHas2023Cert = "Unknown"
        $UefiCA2023Status = "Unknown"
        $AuditStatus = "Failed"
        $ErrorDetails = ""

        if ($PingStatus -eq 'Online' -or $SkipPing) {
            # Script block to execute on the remote machine
            $ScriptBlock = {
                $Result = [ordered]@{
                    BootMode          = "Unknown"
                    SecureBootEnabled = "Unknown"
                    DbHas2023Cert     = "Unknown"
                    KekHas2023Cert    = "Unknown"
                    UefiCA2023Status  = "Unknown"
                    ErrorDetails      = ""
                }

                try {
                    # Test Secure Boot cmdlet support (only UEFI systems running Windows 8+/Server 2012+ support this)
                    $sbConfirm = Confirm-SecureBootUEFI -ErrorAction Stop
                    $Result.BootMode = "UEFI"
                    $Result.SecureBootEnabled = $sbConfirm.ToString()

                    # Check DB (Allowed Signature Database) for the 2023 certificates
                    try {
                        $dbBytes = Get-SecureBootUEFI -Name db -ErrorAction Stop
                        if ($dbBytes -and $dbBytes.Bytes) {
                            $dbText = [System.Text.Encoding]::ASCII.GetString($dbBytes.Bytes)
                            $Result.DbHas2023Cert = ($dbText -match 'Windows UEFI CA 2023').ToString()
                        }
                    }
                    catch {
                        $Result.DbHas2023Cert = "Error: $($_.Exception.Message)"
                    }

                    # Check KEK (Key Exchange Key) for the 2023 certificate
                    try {
                        $kekBytes = Get-SecureBootUEFI -Name kek -ErrorAction Stop
                        if ($kekBytes -and $kekBytes.Bytes) {
                            $kekText = [System.Text.Encoding]::ASCII.GetString($kekBytes.Bytes)
                            $Result.KekHas2023Cert = ($kekText -match 'Microsoft Corporation KEK 2K CA 2023').ToString()
                        }
                    }
                    catch {
                        $Result.KekHas2023Cert = "Error: $($_.Exception.Message)"
                    }

                    # Check registry for the servicing status of the update (KB5036980 onwards)
                    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing"
                    if (Test-Path $regPath) {
                        $status = Get-ItemPropertyValue -Path $regPath -Name "UEFICA2023Status" -ErrorAction SilentlyContinue
                        if ($null -ne $status) {
                            $Result.UefiCA2023Status = $status.ToString()
                        }
                        else {
                            $Result.UefiCA2023Status = "NotSet"
                        }
                    }
                    else {
                        $Result.UefiCA2023Status = "NotStarted"
                    }
                }
                catch {
                    $msg = $_.Exception.Message
                    if ($msg -match "not supported" -or $msg -match "is not enabled on this machine") {
                        $Result.BootMode = "Legacy/BIOS"
                        $Result.SecureBootEnabled = "NotSupported"
                    }
                    else {
                        $Result.ErrorDetails = $msg
                    }
                }

                return $Result
            }

            try {
                $InvokeParams = @{
                    ComputerName = $Target
                    ScriptBlock  = $ScriptBlock
                    ErrorAction  = 'Stop'
                }
                if ($Credential) {
                    $InvokeParams.Credential = $Credential
                }

                $remoteResult = Invoke-Command @InvokeParams

                if ($null -ne $remoteResult) {
                    $BootMode = $remoteResult.BootMode
                    $SecureBootEnabled = $remoteResult.SecureBootEnabled
                    $DbHas2023Cert = $remoteResult.DbHas2023Cert
                    $KekHas2023Cert = $remoteResult.KekHas2023Cert
                    $UefiCA2023Status = $remoteResult.UefiCA2023Status
                    $ErrorDetails = $remoteResult.ErrorDetails

                    if ($ErrorDetails) {
                        $AuditStatus = "PartialSuccess"
                    }
                    else {
                        $AuditStatus = "Success"
                    }
                }
                else {
                    $ErrorDetails = "Received null response from remote session."
                }
            }
            catch {
                $ErrorDetails = $_.Exception.Message
            }
        }
        else {
            $AuditStatus = "Offline"
        }

        [PSCustomObject]@{
            ComputerName             = $Computer.Name
            DNSHostName              = $Computer.DNSHostName
            Enabled                  = $Computer.Enabled
            IPAddress                = $Computer.IPv4Address
            PingStatus               = $PingStatus
            ADOperatingSystem        = $Computer.OperatingSystem
            ADOperatingSystemVersion = $Computer.OperatingSystemVersion
            BootMode                 = $BootMode
            SecureBootEnabled        = $SecureBootEnabled
            DbHas2023Cert            = $DbHas2023Cert
            KekHas2023Cert           = $KekHas2023Cert
            UefiCA2023Status         = $UefiCA2023Status
            AuditStatus              = $AuditStatus
            ErrorDetails             = $ErrorDetails
        }
    }

    $script:ComputersToAudit = [System.Collections.Generic.List[PSCustomObject]]::new()
}

process {
    if ($PSBoundParameters.ContainsKey('ComputerName')) {
        foreach ($Name in $ComputerName) {
            if ($Name) {
                try {
                    $AdComp = Get-ADComputer -Identity $Name -Properties DNSHostName, Enabled, IPv4Address, OperatingSystem, OperatingSystemVersion -ErrorAction Stop
                    $script:ComputersToAudit.Add($AdComp)
                }
                catch {
                    $script:ComputersToAudit.Add([PSCustomObject]@{
                            Name                   = $Name
                            DNSHostName            = $Name
                            Enabled                = $true
                            IPv4Address            = $null
                            OperatingSystem        = "Unknown"
                            OperatingSystemVersion = "Unknown"
                        })
                }
            }
        }
    }
}

end {
    if ($PSCmdlet.ParameterSetName -eq "ADSearch") {
        $ADQueryParameters = @{
            Filter      = "OperatingSystem -like '*Windows*'"
            Properties  = @(
                'DNSHostName'
                'Enabled'
                'IPv4Address'
                'OperatingSystem'
                'OperatingSystemVersion'
            )
            ErrorAction = 'Stop'
        }

        if ($SearchBase) {
            $ADQueryParameters.SearchBase = $SearchBase
        }

        Write-Verbose "Querying Active Directory for Windows computer objects."
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
        Write-Warning "No computer objects were found to audit."
        return
    }

    Write-Verbose "Starting audit of $($script:ComputersToAudit.Count) computers."

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $GetSecureBootAuditResultDefinition = ${function:Get-SecureBootAuditResult}.ToString()

        $script:ComputersToAudit | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            ${function:Get-SecureBootAuditResult} = $using:GetSecureBootAuditResultDefinition

            Get-SecureBootAuditResult `
                -Computer $_ `
                -SkipPing ([bool]$using:SkipPing) `
                -PingTimeoutSeconds $using:PingTimeoutSeconds `
                -Credential $using:Credential
        }
    }
    else {
        foreach ($Computer in $script:ComputersToAudit) {
            Get-SecureBootAuditResult `
                -Computer $Computer `
                -SkipPing ([bool]$SkipPing) `
                -PingTimeoutSeconds $PingTimeoutSeconds `
                -Credential $Credential
        }
    }
}
