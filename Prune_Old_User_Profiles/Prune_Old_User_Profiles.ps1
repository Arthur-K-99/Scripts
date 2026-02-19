# ==========================================
# CONFIGURATION
# ==========================================
$DaysThreshold = 0
# SET THIS TO $true TO ENABLE ACTUAL DELETION
$DeleteMode = $true
$LogPrefix = "[PROFILE-CLEANUP]"

# Calculates the date before which profiles are considered old
$CutoffDate = (Get-Date).AddDays(-$DaysThreshold)

# ==========================================
# EXCLUSIONS
# ==========================================
# 1. Always protect these specific paths
$ProtectedPaths = @(
    'C:\Users\Public'
    'C:\Users\Default'
    'C:\Users\Default User'
)

# 2. Dynamic Protections (The account running the script + Admin accounts)
$CurrentRunner = $env:USERNAME
$ProtectedPatterns = @(
    $CurrentRunner    # Protect the PDQ Runner/Service Account
    "Administrator"   # Protect the built-in Admin
    "LAPS_Admin"      # Example: Add other specific admin accounts here
)

# ==========================================
# EXECUTION
# ==========================================
Write-Output "$LogPrefix Starting scan. Threshold: $DaysThreshold days. (Cutoff: $($CutoffDate.ToString('yyyy-MM-dd')))"
Write-Output "$LogPrefix Run Mode: $(If ($DeleteMode) {'DESTRUCTIVE'} Else {'REPORT ONLY (Dry Run)'})"
Write-Output "$LogPrefix Script Runner: $CurrentRunner"

# Stats Counters
$Stats = @{ Removed = 0; Skipped = 0; Errors = 0; Orphans = 0 }

# Get Candidates
# Filter: Not Special, Not Currently Loaded, Must be in C:\Users (Avoids System Profiles)
try {
    $Candidates = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop | Where-Object {
        $_.Special -eq $false -and 
        $_.Loaded -eq $false -and 
        $_.LocalPath -like "C:\Users\*"
    }
}
catch {
    Write-Error "$LogPrefix FATAL: Could not query Win32_UserProfile. WMI may be broken on this host."
    exit 1
}

foreach ($UserProfile in $Candidates) {
    
    $Path = $UserProfile.LocalPath
    $SID = $UserProfile.SID
    $Username = Split-Path $Path -Leaf

    # CHECK 1: Exclusions
    if ($Path -in $ProtectedPaths -or ($ProtectedPatterns | Where-Object { $Username -like $_ })) {
        Write-Output "$LogPrefix SKIPPED: [$Path] matches exclusion list."
        $Stats.Skipped++
        continue
    }

    # CHECK 2: Primary WMI Age Check
    if ($UserProfile.LastUseTime -lt $CutoffDate) {
        
        $IsActuallyStale = $true

        # CHECK 3: Filesystem Validation (The "Triple-Tap")
        # We check both NTUSER.DAT and UsrClass.dat. If EITHER is new, we keep the profile.
        $HivePaths = @(
            (Join-Path $Path "NTUSER.DAT"),
            (Join-Path $Path "AppData\Local\Microsoft\Windows\UsrClass.dat")
        )

        # Check if the folder actually exists
        if (Test-Path $Path) {
            $NewestHiveDate = $null

            foreach ($Hive in $HivePaths) {
                if (Test-Path $Hive) {
                    $HiveDate = (Get-Item $Hive -Force).LastWriteTime
                    if ($null -eq $NewestHiveDate -or $HiveDate -gt $NewestHiveDate) {
                        $NewestHiveDate = $HiveDate
                    }
                }
            }

            if ($NewestHiveDate -and $NewestHiveDate -gt $CutoffDate) {
                Write-Output "$LogPrefix SKIPPED: [$Path]. WMI says old, but Registry Hive modified recently ($NewestHiveDate)."
                $IsActuallyStale = $false
                $Stats.Skipped++
            }
        } 
        elseif (-not (Test-Path $Path)) {
            Write-Output "$LogPrefix ORPHAN FOUND: [$Path] does not exist on disk but exists in WMI."
            $Stats.Orphans++
            # We let $IsActuallyStale remain true so we clean up the dead WMI entry
        }

        # ACTION: Delete
        if ($IsActuallyStale) {
            
            if ($DeleteMode) {
                Write-Output "$LogPrefix REMOVING: [$Path] | SID: $SID | Last Used: $($UserProfile.LastUseTime)"
                try {
                    Remove-CimInstance -InputObject $UserProfile -ErrorAction Stop -Confirm:$false
                    
                    Write-Output "$LogPrefix SUCCESS: Removed $Path"
                    $Stats.Removed++
                }
                catch {
                    Write-Output "$LogPrefix ERROR: Failed to remove $Path. Exception: $($_.Exception.Message)"
                    $Stats.Errors++
                }
            }
            else {
                Write-Output "$LogPrefix [DRY RUN] WOULD REMOVE: [$Path] | SID: $SID | Last Used: $($UserProfile.LastUseTime)"
                $Stats.Removed++ # Count as removed for the report
            }
        }
    }
}

# Summary for PDQ Output
Write-Output "--------------------------------------------------"
Write-Output "$LogPrefix SUMMARY COMPLETE"
Write-Output "Processed: $($Stats.Removed) | Skipped: $($Stats.Skipped) | Orphans Found: $($Stats.Orphans) | Errors: $($Stats.Errors)"
Write-Output "--------------------------------------------------"
