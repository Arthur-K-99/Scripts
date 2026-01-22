<#
.SYNOPSIS
    Installs a printer using the IPP Protocol (HTTP/631) and creates the necessary port and driver associations.

.DESCRIPTION
    Stages a driver from a specified INF file (local or UNC), registers it, and maps the printer 
    to an IPP URL (http://<IP>:631/ipp/print) using printui.dll.

.PARAMETER PrinterName
    The name of the printer to be created on the system.
.PARAMETER DriverName
    The exact name of the driver model as it appears inside the INF file.
.PARAMETER DriverInfPath
    The full path (UNC or Local) to the driver .INF file.
.PARAMETER PrinterIP
    The IP address of the printer.
.PARAMETER PortName
    Optional. Defaults to the standard IPP path: http://<IP>:631/ipp/print.
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]
    [string]$PrinterName,

    [Parameter(Mandatory = $true)]
    [string]$DriverName,

    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$DriverInfPath,

    [Parameter(Mandatory = $true)]
    [string]$PrinterIP,

    [string]$PortName
)

$ErrorActionPreference = "Stop"

# Construct default IPP URL if PortName is not explicitly provided
if ([string]::IsNullOrWhiteSpace($PortName)) {
    $PortName = "http://$($PrinterIP):631/ipp/print"
}

Write-Verbose "Configuration :: Printer: $PrinterName | IP: $PrinterIP | Driver: $DriverName"

try {
    # --- Step 1: Cleanup Existing Instances ---
    if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
        Write-Host "[-] Printer '$PrinterName' exists. Removing to enforce clean configuration..."
        Remove-Printer -Name $PrinterName
    }

    # --- Step 2: Stage Driver via PNPUtil ---
    Write-Host "[*] Staging driver from: $DriverInfPath"
    $pnp = Start-Process pnputil.exe -ArgumentList "/add-driver `"$DriverInfPath`" /install" -Wait -PassThru
    
    if ($pnp.ExitCode -ne 0) {
        Write-Warning "PNPUtil returned exit code $($pnp.ExitCode). Check if driver requires a reboot or is already present."
    }

    # --- Step 3: Register Driver with Spooler ---
    Write-Host "[*] Registering Driver: $DriverName"
    if (-not (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
        Add-PrinterDriver -Name $DriverName
    }
    else {
        Write-Host "    -> Driver already registered."
    }

    # --- Step 4: Install Printer (IPP Port) ---
    # Using rundll32 printui.dll because Add-PrinterPort does not natively support IPP creation easily.
    Write-Host "[*] Creating IPP connection to $PortName..."
    
    $PrintUIArgs = "/if /b `"$PrinterName`" /r `"$PortName`" /m `"$DriverName`" /z /u"
    $proc = Start-Process "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry $PrintUIArgs" -Wait -PassThru

    if ($proc.ExitCode -ne 0) {
        throw "PrintUI execution failed with code $($proc.ExitCode)."
    }

    # --- Step 5: Verification ---
    Start-Sleep -Seconds 3 # Allow spooler to refresh
    if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
        Write-Host "[+] Success: Printer '$PrinterName' installed successfully." -ForegroundColor Green
    }
    else {
        throw "Installation verification failed. Printer object not found."
    }
}
catch {
    Write-Error "FATAL: $($_.Exception.Message)"
    exit 1
}