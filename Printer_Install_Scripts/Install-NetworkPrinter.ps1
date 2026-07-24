<#
.SYNOPSIS
    Installs a network printer with an OEM driver over IPP, RAW/TCP, or LPR.

.DESCRIPTION
    A deployment-tool-neutral Windows printer installer suitable for PDQ Deploy,
    Microsoft Configuration Manager, Intune, an elevated console, or another
    system-management tool. It stages a complete OEM driver package locally,
    verifies an optional preapproved catalog signer, registers the driver,
    creates the requested port and queue, logs live progress, and verifies the
    result. External processes have hard timeouts.

    Supported protocols:
      IPP - OEM driver on an http:// or https:// IPP URL through PrintUIEntry.
      TCP - Standard TCP/IP port, normally RAW port 9100.
      LPR - Standard TCP/IP port in LPR mode, normally TCP port 515.

.PARAMETER PrinterName
    Name of the local Windows printer queue.

.PARAMETER DriverName
    Exact printer model name contained in the driver's INF.

.PARAMETER DriverInfPath
    Full local or UNC path to the driver INF. Adjacent driver files are copied too.

.PARAMETER PrinterAddress
    Printer IP address or resolvable hostname.

.PARAMETER Protocol
    IPP, TCP, or LPR. Default: IPP.

.PARAMETER PortName
    Optional explicit Windows port name. For IPP, supply the complete IPP URL.

.PARAMETER IppPath
    IPP resource path used when PortName is omitted. Default: /ipp/print.

.PARAMETER IppPort
    IPP TCP port used when PortName is omitted. Default: 631.

.PARAMETER UseIpps
    Builds an https:// IPP URL instead of http:// when PortName is omitted.

.PARAMETER TcpPortNumber
    RAW/TCP destination port. Default: 9100.

.PARAMETER LprQueueName
    Remote LPR queue name. Required when Protocol is LPR.

.PARAMETER DriverPublisherThumbprint
    Optional preapproved SHA-1 thumbprint of the driver catalog signer. The
    certificate is added to Local Machine Trusted Publishers only after the
    script verifies a valid catalog signature with this exact thumbprint.

.PARAMETER ForceRecreate
    Recreates an already compliant printer queue.

.PARAMETER SkipConnectivityTest
    Skips the pre-install TCP connection test.

.EXAMPLE
    .\Install-NetworkPrinter.ps1 -PrinterName "HR Color" `
        -DriverName "Brother MFC-L8930CDW Printer" `
        -DriverInfPath "\\server\drivers\Brother\BRPRC23A.INF" `
        -PrinterAddress "10.10.13.31" -Protocol IPP `
        -DriverPublisherThumbprint "059ED235B29963FE15AE5B4EA92820AFF3C9C0D2"

.EXAMPLE
    .\Install-NetworkPrinter.ps1 -PrinterName "Front Office Xerox" `
        -DriverName "Xerox Global Print Driver PCL6" `
        -DriverInfPath "\\server\drivers\Xerox\x3UNIVX.inf" `
        -PrinterAddress "10.10.20.50" -Protocol TCP

.EXAMPLE
    .\Install-NetworkPrinter.ps1 -PrinterName "Legacy LPR Queue" `
        -DriverName "Generic PCL6 Driver" `
        -DriverInfPath "C:\Drivers\Generic\driver.inf" `
        -PrinterAddress "printer.example.org" -Protocol LPR `
        -LprQueueName "PASSTHRU"

.NOTES
    Requirements: Windows 10/11 or supported Windows Server, local administrator,
    64-bit Windows PowerShell 5.1, Print Spooler, and PrintManagement module.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$PrinterName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DriverName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DriverInfPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$PrinterAddress,

    [ValidateSet("IPP", "TCP", "LPR")]
    [string]$Protocol = "IPP",

    [string]$PortName,

    [string]$IppPath = "/ipp/print",

    [ValidateRange(1, 65535)]
    [int]$IppPort = 631,

    [switch]$UseIpps,

    [ValidateRange(1, 65535)]
    [int]$TcpPortNumber = 9100,

    [string]$LprQueueName,

    [string]$DriverPublisherThumbprint,

    [string]$LogDirectory = "$env:ProgramData\PrinterDeployment\Logs",

    [string]$DriverCacheRoot = "$env:ProgramData\PrinterDeployment\Drivers",

    [ValidateRange(5, 120)]
    [int]$ConnectTimeoutSeconds = 10,

    [ValidateRange(30, 900)]
    [int]$DriverCopyTimeoutSeconds = 180,

    [ValidateRange(30, 900)]
    [int]$DriverStageTimeoutSeconds = 180,

    [ValidateRange(30, 900)]
    [int]$PrintUITimeoutSeconds = 120,

    [switch]$ForceRecreate,

    [switch]$SkipConnectivityTest
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$SafePrinterName = $PrinterName -replace '[^A-Za-z0-9._-]', '_'
New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
$LogFile = Join-Path $LogDirectory "$SafePrinterName-install.log"

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $Line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -LiteralPath $LogFile -Value $Line -Encoding UTF8
    Write-Host $Line
}

function Invoke-NativeProcess {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string]$Arguments,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 3600)]
        [int]$TimeoutSeconds,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $Process = $null

    try {
        Write-Log "Starting $Description."

        $StartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $StartInfo.FileName = $FilePath
        $StartInfo.Arguments = $Arguments
        $StartInfo.UseShellExecute = $false
        $StartInfo.CreateNoWindow = $true
        $StartInfo.RedirectStandardOutput = $true
        $StartInfo.RedirectStandardError = $true

        $Process = New-Object System.Diagnostics.Process
        $Process.StartInfo = $StartInfo
        if (-not $Process.Start()) {
            throw "Unable to start $Description."
        }

        $StdOutTask = $Process.StandardOutput.ReadToEndAsync()
        $StdErrTask = $Process.StandardError.ReadToEndAsync()

        if (-not $Process.WaitForExit($TimeoutSeconds * 1000)) {
            Write-Log "$Description exceeded its $TimeoutSeconds-second timeout; terminating PID $($Process.Id)." "ERROR"
            & "$env:SystemRoot\System32\taskkill.exe" /PID $Process.Id /T /F 2>&1 | Out-Null
            throw "$Description timed out after $TimeoutSeconds seconds."
        }

        $Process.WaitForExit()
        $ExitCode = $Process.ExitCode
        $StdOut = $StdOutTask.GetAwaiter().GetResult()
        $StdErr = $StdErrTask.GetAwaiter().GetResult()

        $StdOut -split "`r?`n" | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_)) {
                Write-Log "$Description output: $_"
            }
        }

        $StdErr -split "`r?`n" | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_)) {
                Write-Log "$Description error output: $_" "WARN"
            }
        }

        Write-Log "$Description completed with exit code $ExitCode."

        [PSCustomObject]@{
            ExitCode  = $ExitCode
            ProcessId = $Process.Id
        }
    }
    finally {
        if ($null -ne $Process) {
            $Process.Dispose()
        }
    }
}

function Test-TcpEndpoint {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [int]$Port,

        [Parameter(Mandatory = $true)]
        [int]$TimeoutSeconds
    )

    $Client = New-Object System.Net.Sockets.TcpClient
    $AsyncResult = $null

    try {
        $AsyncResult = $Client.BeginConnect($ComputerName, $Port, $null, $null)
        if (-not $AsyncResult.AsyncWaitHandle.WaitOne($TimeoutSeconds * 1000, $false)) {
            throw "TCP connection to $ComputerName`:$Port timed out after $TimeoutSeconds seconds."
        }

        $Client.EndConnect($AsyncResult)
        return $true
    }
    finally {
        if ($null -ne $AsyncResult) {
            $AsyncResult.AsyncWaitHandle.Close()
        }
        $Client.Close()
    }
}

function Confirm-DriverPublisher {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriverDirectory,

        [Parameter(Mandatory = $true)]
        [string]$InfPath,

        [string]$ApprovedThumbprint
    )

    # Prefer catalogs explicitly named by the selected INF. Fall back to every
    # catalog in the driver folder when the INF does not expose a literal name.
    $CatalogNames = @(
        Select-String -LiteralPath $InfPath -Pattern '^\s*CatalogFile(?:\.[^=]+)?\s*=\s*([^;\r\n]+)' -AllMatches |
        ForEach-Object {
            foreach ($Match in $_.Matches) {
                $Match.Groups[1].Value.Trim().Trim('"')
            }
        } |
        Select-Object -Unique
    )

    $CatalogFiles = @()
    foreach ($CatalogName in $CatalogNames) {
        $CatalogFiles += @(Get-ChildItem -LiteralPath $DriverDirectory -Filter $CatalogName -File -Recurse -ErrorAction SilentlyContinue)
    }

    if ($CatalogFiles.Count -eq 0) {
        $CatalogFiles = @(Get-ChildItem -LiteralPath $DriverDirectory -Filter "*.cat" -File -Recurse -ErrorAction Stop)
    }

    if ($CatalogFiles.Count -eq 0) {
        Write-Log "No catalog files were found for '$InfPath'." "WARN"
        return
    }

    $ApprovedNormalized = $null
    if (-not [string]::IsNullOrWhiteSpace($ApprovedThumbprint)) {
        $ApprovedNormalized = ($ApprovedThumbprint -replace '[^A-Fa-f0-9]', '').ToUpperInvariant()
        if ($ApprovedNormalized.Length -ne 40) {
            throw "DriverPublisherThumbprint must contain exactly 40 hexadecimal characters."
        }
    }

    $ApprovedSignature = $null
    foreach ($Catalog in ($CatalogFiles | Sort-Object FullName -Unique)) {
        $Signature = Get-AuthenticodeSignature -LiteralPath $Catalog.FullName
        $Signer = $Signature.SignerCertificate

        if ($null -eq $Signer) {
            Write-Log "Catalog '$($Catalog.Name)' has no readable signer; status='$($Signature.Status)'." "WARN"
            continue
        }

        $Thumbprint = ($Signer.Thumbprint -replace '\s', '').ToUpperInvariant()
        Write-Log "Catalog='$($Catalog.Name)'; signer='$($Signer.Subject)'; thumbprint='$Thumbprint'; status='$($Signature.Status)'."

        if ($null -ne $ApprovedNormalized -and $Thumbprint -eq $ApprovedNormalized) {
            if ($Signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid) {
                throw "Approved catalog '$($Catalog.Name)' has signature status '$($Signature.Status)': $($Signature.StatusMessage)"
            }
            $ApprovedSignature = $Signature
            break
        }
    }

    if ($null -eq $ApprovedNormalized) {
        Write-Log "No approved publisher thumbprint supplied; certificate stores will not be changed."
        return
    }

    if ($null -eq $ApprovedSignature) {
        throw "No valid selected-driver catalog matched approved thumbprint '$ApprovedNormalized'."
    }

    $CertificatePath = "Cert:\LocalMachine\TrustedPublisher\$ApprovedNormalized"
    if (Test-Path -LiteralPath $CertificatePath) {
        Write-Log "Approved driver publisher is already trusted."
        return
    }

    $Store = New-Object `
        -TypeName System.Security.Cryptography.X509Certificates.X509Store `
        -ArgumentList @(
        "TrustedPublisher",
        [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
    )

    try {
        $Store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $Store.Add($ApprovedSignature.SignerCertificate)
    }
    finally {
        $Store.Close()
    }

    if (-not (Test-Path -LiteralPath $CertificatePath)) {
        throw "Failed to add the approved driver publisher certificate to Local Machine Trusted Publishers."
    }

    Write-Log "Added verified publisher '$($ApprovedSignature.SignerCertificate.Subject)' to Local Machine Trusted Publishers." "WARN"
}

try {
    Add-Content -LiteralPath $LogFile -Value "`r`n============================================================" -Encoding UTF8
    Write-Log "Beginning network-printer deployment."

    if (-not [Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem) {
        throw "Run this script in 64-bit Windows PowerShell."
    }

    Import-Module PrintManagement -ErrorAction Stop

    if ($Protocol -eq "LPR" -and [string]::IsNullOrWhiteSpace($LprQueueName)) {
        throw "LprQueueName is required when Protocol is LPR."
    }

    $ConnectivityPort = 0
    if ([string]::IsNullOrWhiteSpace($PortName)) {
        switch ($Protocol) {
            "IPP" {
                if (-not $IppPath.StartsWith("/")) {
                    $IppPath = "/$IppPath"
                }
                $Scheme = if ($UseIpps) { "https" } else { "http" }
                $PortName = "$Scheme`://$PrinterAddress`:$IppPort$IppPath"
                $ConnectivityPort = $IppPort
            }
            "TCP" {
                $PortName = "TCP_$($PrinterAddress)_$TcpPortNumber"
                $ConnectivityPort = $TcpPortNumber
            }
            "LPR" {
                $PortName = "LPR_$($PrinterAddress)_$LprQueueName"
                $ConnectivityPort = 515
            }
        }
    }
    elseif ($Protocol -eq "IPP") {
        $IppUri = $null
        if (-not [Uri]::TryCreate($PortName, [UriKind]::Absolute, [ref]$IppUri)) {
            throw "IPP PortName must be a valid absolute URL."
        }
        if ($IppUri.Scheme -notin @("http", "https")) {
            throw "IPP PortName must use http:// or https://."
        }
        $PrinterAddress = $IppUri.Host
        $ConnectivityPort = $IppUri.Port
    }
    elseif ($Protocol -eq "TCP") {
        $ConnectivityPort = $TcpPortNumber
    }
    else {
        $ConnectivityPort = 515
    }

    Write-Log "Printer='$PrinterName'; protocol='$Protocol'; address='$PrinterAddress'; port='$PortName'; driver='$DriverName'."

    $Spooler = Get-Service -Name Spooler -ErrorAction Stop
    if ($Spooler.Status -ne "Running") {
        Write-Log "Print Spooler is $($Spooler.Status); starting it." "WARN"
        Start-Service -Name Spooler
        $Spooler.WaitForStatus("Running", [TimeSpan]::FromSeconds(30))
    }

    $ExistingPrinter = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
    if ($null -ne $ExistingPrinter) {
        $Compliant = $ExistingPrinter.DriverName -eq $DriverName -and $ExistingPrinter.PortName -eq $PortName
        if ($Compliant -and -not $ForceRecreate) {
            Write-Log "Printer already has the requested driver and port; no changes required." "SUCCESS"
            Write-Log "Live log: $LogFile"
            exit 0
        }

        Write-Log "Removing existing queue with driver='$($ExistingPrinter.DriverName)' and port='$($ExistingPrinter.PortName)'." "WARN"
        Remove-Printer -Name $PrinterName -Confirm:$false

        $Deadline = (Get-Date).AddSeconds(30)
        while ((Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) -and (Get-Date) -lt $Deadline) {
            Start-Sleep -Seconds 1
        }
        if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
            throw "Printer '$PrinterName' still exists after the 30-second removal timeout."
        }
    }

    $ExistingPort = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
    if ($null -ne $ExistingPort) {
        $PortUsers = @(Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $PortName })
        if ($PortUsers.Count -eq 0) {
            Write-Log "Removing orphaned port '$PortName'." "WARN"
            Remove-PrinterPort -Name $PortName -Confirm:$false
            $ExistingPort = $null
        }
        else {
            Write-Log "Port '$PortName' is shared with another queue and will be preserved."
        }
    }

    $SourceDirectory = (Split-Path -Path $DriverInfPath -Parent).TrimEnd("\")
    $InfFileName = Split-Path -Path $DriverInfPath -Leaf
    $SafeDriverName = $DriverName -replace '[^A-Za-z0-9._-]', '_'
    $LocalDirectory = (Join-Path $DriverCacheRoot $SafeDriverName).TrimEnd("\")
    $LocalInf = Join-Path $LocalDirectory $InfFileName

    if ($SourceDirectory -ine $LocalDirectory) {
        Remove-Item -LiteralPath $LocalDirectory -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -Path $LocalDirectory -ItemType Directory -Force | Out-Null

        $CopyArguments = "`"$SourceDirectory`" `"$LocalDirectory`" /E /COPY:DAT /DCOPY:T /R:2 /W:2 /NP /NFL /NDL /NJH /NJS"
        $CopyResult = Invoke-NativeProcess -FilePath "$env:SystemRoot\System32\robocopy.exe" `
            -Arguments $CopyArguments -TimeoutSeconds $DriverCopyTimeoutSeconds `
            -Description "driver folder copy"

        if ($CopyResult.ExitCode -ge 8) {
            throw "Driver folder copy failed with Robocopy exit code $($CopyResult.ExitCode)."
        }
    }
    else {
        New-Item -Path $LocalDirectory -ItemType Directory -Force | Out-Null
        Write-Log "Driver is already in the local cache; copy skipped."
    }

    if (-not (Test-Path -LiteralPath $LocalInf -PathType Leaf)) {
        throw "Driver INF was not found after local staging: $LocalInf"
    }

    Write-Log "Driver package cached at '$LocalDirectory'."
    Confirm-DriverPublisher -DriverDirectory $LocalDirectory -InfPath $LocalInf `
        -ApprovedThumbprint $DriverPublisherThumbprint

    $PnpResult = Invoke-NativeProcess -FilePath "$env:SystemRoot\System32\pnputil.exe" `
        -Arguments "/add-driver `"$LocalInf`"" -TimeoutSeconds $DriverStageTimeoutSeconds `
        -Description "PnPUtil driver staging"

    if ($PnpResult.ExitCode -notin @(0, 259, 1641, 3010)) {
        throw "PnPUtil failed with exit code $($PnpResult.ExitCode). Review the logged PnPUtil output."
    }

    if (-not (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
        Write-Log "Registering printer driver '$DriverName'."
        Add-PrinterDriver -Name $DriverName -ErrorAction Stop
    }
    else {
        Write-Log "Printer driver is already registered."
    }

    if (-not $SkipConnectivityTest) {
        Write-Log "Testing TCP connectivity to $PrinterAddress`:$ConnectivityPort."
        if (-not (Test-TcpEndpoint -ComputerName $PrinterAddress -Port $ConnectivityPort -TimeoutSeconds $ConnectTimeoutSeconds)) {
            throw "TCP connectivity test failed."
        }
        Write-Log "TCP connectivity test succeeded."
    }
    else {
        Write-Log "Connectivity test skipped by request." "WARN"
    }

    switch ($Protocol) {
        "IPP" {
            $Arguments = "printui.dll,PrintUIEntry /if /b `"$PrinterName`" /r `"$PortName`" /m `"$DriverName`" /z /u /q /Y"
            $PrintResult = Invoke-NativeProcess -FilePath "$env:SystemRoot\System32\rundll32.exe" `
                -Arguments $Arguments -TimeoutSeconds $PrintUITimeoutSeconds `
                -Description "PrintUI IPP queue creation"
            if ($PrintResult.ExitCode -ne 0) {
                throw "PrintUI failed with exit code $($PrintResult.ExitCode)."
            }
        }
        "TCP" {
            if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
                Write-Log "Creating RAW/TCP port '$PortName' at $PrinterAddress`:$TcpPortNumber."
                Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterAddress `
                    -PortNumber $TcpPortNumber -ErrorAction Stop
            }
            Write-Log "Creating TCP printer queue '$PrinterName'."
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -ErrorAction Stop
        }
        "LPR" {
            if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
                Write-Log "Creating LPR port '$PortName' at $PrinterAddress with queue '$LprQueueName'."
                Add-PrinterPort -Name $PortName -LprHostAddress $PrinterAddress `
                    -LprQueueName $LprQueueName -ErrorAction Stop
            }
            Write-Log "Creating LPR printer queue '$PrinterName'."
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -ErrorAction Stop
        }
    }

    $Deadline = (Get-Date).AddSeconds(30)
    $InstalledPrinter = $null
    do {
        Start-Sleep -Seconds 2
        $InstalledPrinter = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
    } while ($null -eq $InstalledPrinter -and (Get-Date) -lt $Deadline)

    if ($null -eq $InstalledPrinter) {
        throw "Printer '$PrinterName' was not found within the verification timeout."
    }
    if ($InstalledPrinter.DriverName -ne $DriverName) {
        throw "Expected driver '$DriverName', found '$($InstalledPrinter.DriverName)'."
    }
    if ($InstalledPrinter.PortName -ne $PortName) {
        throw "Expected port '$PortName', found '$($InstalledPrinter.PortName)'."
    }

    Write-Log "Printer '$PrinterName' installed successfully using $Protocol on '$PortName'." "SUCCESS"
    Write-Log "Live log: $LogFile"
    exit 0
}
catch {
    $Message = $_.Exception.Message
    Write-Log "FATAL: $Message" "ERROR"
    Write-Log "Live log: $LogFile" "ERROR"
    Write-Error "Network-printer deployment failed: $Message" -ErrorAction Continue
    exit 1
}