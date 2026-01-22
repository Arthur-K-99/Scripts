<#
.SYNOPSIS
    Disables Windows default printer management and forces a specific default printer.
    
.DESCRIPTION
    Sets the 'LegacyDefaultPrinterMode' registry key to prevent Windows from managing the default printer,
    then uses CIM/WMI to set the specified printer as default.
    
.PARAMETER PrinterName
    The exact name of the printer to set as default.
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]
    [string]$PrinterName
)

$ErrorActionPreference = "Stop"

try {
    # --- Step 1: Disable "Let Windows manage my default printer" ---
    Write-Host "[*] Configuring registry to disable Windows printer management..."
    
    $RegPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows"
    
    if (!(Test-Path $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }

    # 1 = Disable Windows Management
    Set-ItemProperty -Path $RegPath -Name "LegacyDefaultPrinterMode" -Value 1 -Type DWORD -Force

    # --- Step 2: Validate Printer Existence ---
    $TargetPrinter = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$PrinterName'" -ErrorAction SilentlyContinue

    if (-not $TargetPrinter) {
        throw "Printer '$PrinterName' not found. Ensure it is installed before running this script."
    }

    # --- Step 3: Set Default ---
    Write-Host "[*] Setting default printer to: $PrinterName"
    Invoke-CimMethod -InputObject $TargetPrinter -MethodName SetDefaultPrinter | Out-Null
    
    # --- Verification ---
    $CurrentDefault = Get-CimInstance -ClassName Win32_Printer -Filter "Default=$true"
    if ($CurrentDefault.Name -eq $PrinterName) {
        Write-Host "[+] Success: Default printer is now '$PrinterName'." -ForegroundColor Green
    }
    else {
        throw "Verification failed. Current default is still: '$($CurrentDefault.Name)'"
    }
}
catch {
    Write-Error "FATAL: $($_.Exception.Message)"
    exit 1
}