#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    One-stop Entra ID (Azure AD) identity & authentication audit toolkit.

.DESCRIPTION
    A single console-driven tool that replaces a dozen Entra admin-center blades.
    Connects once to Microsoft Graph and runs any combination of reports against
    users you scope by group, by UPN, or tenant-wide.

    REPORTS AVAILABLE
    ------------------------------------------------------------------------------
      RegistrationSummary  MFA / SSPR / Authenticator registration (per user)
      AuthMethods          Detailed registered methods (phones, FIDO2, WHfB, TAP, OATH)
      PerUserMfa           Legacy per-user MFA state (disabled/enabled/enforced)
      SignInActivity       Last interactive + non-interactive sign-in per user
      StaleAccounts        Enabled accounts with no sign-in in N days
      PasswordAudit        Password-never-expires, last-set date, on-prem sync
      RiskyUsers           Identity Protection risky users
      RiskDetections       Identity Protection risk detections (last N days)
      RiskySignIns         Sign-ins flagged at/above a risk level (last N days)
      ConditionalAccess    All Conditional Access policies + state
      AdminRoles           Directory role assignments (active + PIM-eligible)
      Licenses             Per-user license (SKU) assignment
      Devices              Registered / owned devices per user
      OAuthGrants          Delegated OAuth2 consent grants (consent-phishing hunt)
      AppCredentials       App registration secrets/certs + expiry
      AuditLogs            Directory audit events (last N days)
      GroupMembership      Transitive group membership per user
      TenantOverview       Org name, domains, user/license counts, security defaults
      GroupDirectory       Browsable list of ALL security / M365 / distribution groups

    Output: color-coded console, per-report CSV files, and a single consolidated
    self-contained HTML dashboard.

.PARAMETER Report
    One or more reports to run (see list above). Use 'All' for everything.
    Omit to launch the interactive menu.

.PARAMETER Groups
    Display name(s) or Object ID(s) to scope user-level reports to.

.PARAMETER Users
    UPN(s) or Object ID(s) to scope user-level reports to.

.PARAMETER AllUsers
    Run user-level reports against every user in the tenant. Use with care.

.PARAMETER Days
    Look-back window for time-bound reports (sign-ins, risk, audit). Default 30.

.PARAMETER StaleDays
    Threshold (days) for the StaleAccounts report. Default 90.

.PARAMETER RiskLevelFilter
    Minimum risk level for RiskySignIns: low | medium | high. Default medium.

.PARAMETER OutputFolder
    Folder for CSV + HTML output. Default .\EntraAudit_<timestamp>

.PARAMETER NoHtml
    Skip the HTML dashboard (CSV only).

.PARAMETER UseBeta
    Use the Graph /beta endpoint where it exposes richer fields.

.EXAMPLE
    .\Invoke-EntraAuditToolkit.ps1
    # Interactive menu

.EXAMPLE
    .\Invoke-EntraAuditToolkit.ps1 -Report RegistrationSummary,RiskyUsers,AdminRoles -Groups "All Staff"

.EXAMPLE
    .\Invoke-EntraAuditToolkit.ps1 -Report All -AllUsers -Days 14 -StaleDays 60

.NOTES
    Graph scopes requested (read-only):
        User.Read.All  Group.Read.All  GroupMember.Read.All
        UserAuthenticationMethod.Read.All  AuditLog.Read.All  Reports.Read.All
        Policy.Read.All  IdentityRiskyUser.Read.All  IdentityRiskEvent.Read.All
        RoleManagement.Read.Directory  RoleManagement.Read.All
        Directory.Read.All  Application.Read.All  Device.Read.All
        Organization.Read.All

    Requires the Microsoft.Graph PowerShell SDK:
        Install-Module Microsoft.Graph -Scope CurrentUser
#>

[CmdletBinding()]
param (
    [ValidateSet(
        'All', 'TenantOverview', 'RegistrationSummary', 'AuthMethods', 'PerUserMfa',
        'SignInActivity', 'StaleAccounts', 'PasswordAudit', 'RiskyUsers', 'RiskDetections',
        'RiskySignIns', 'ConditionalAccess', 'AdminRoles', 'Licenses', 'Devices',
        'OAuthGrants', 'AppCredentials', 'AuditLogs', 'GroupMembership', 'GroupDirectory'
    )]
    [string[]] $Report,

    [string[]] $Groups,
    [string[]] $Users,
    [switch]   $AllUsers,

    [int]    $Days = 30,
    [int]    $StaleDays = 90,

    [ValidateSet('low', 'medium', 'high')]
    [string] $RiskLevelFilter = 'medium',

    [string] $OutputFolder = ".\EntraAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    [switch] $NoHtml,
    [switch] $UseBeta
)

$script:GraphBase = if ($UseBeta) { 'https://graph.microsoft.com/beta' } else { 'https://graph.microsoft.com/v1.0' }
$script:Sections = [ordered]@{}   # report name -> array of result objects (for HTML)

#region ── Console helpers ──────────────────────────────────────────────────────

function Write-Banner {
    param([string]$Text, [System.ConsoleColor]$Color = 'Cyan')
    Write-Host "`n$('═'*78)" -ForegroundColor DarkGray
    Write-Host "  $Text"     -ForegroundColor $Color
    Write-Host "$('═'*78)"   -ForegroundColor DarkGray
}
function Write-Sub {
    param([string]$Text)
    Write-Host "`n  ── $Text $('─'*([math]::Max(0,68-$Text.Length)))" -ForegroundColor DarkCyan
}
function Write-Info { param($m) Write-Host "  [*] $m" -ForegroundColor Cyan }
function Write-Good { param($m) Write-Host "  [+] $m" -ForegroundColor Green }
function Write-Warn2 { param($m) Write-Host "  [!] $m" -ForegroundColor Yellow }
function Write-Bad { param($m) Write-Host "  [x] $m" -ForegroundColor Red }

#endregion

#region ── Graph plumbing ───────────────────────────────────────────────────────

function Invoke-Graph {
    <#
        Wrapper around Invoke-MgGraphRequest that:
          - follows @odata.nextLink paging and returns a flat array of items
          - retries on 429 / 503 honoring Retry-After
          - returns $null on hard failure (caller decides)
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        [int]$MaxRetries = 5
    )
    if ($Uri -notmatch '^https?://') { $Uri = "$script:GraphBase/$($Uri.TrimStart('/'))" }

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    do {
        $attempt = 0
        while ($true) {
            try {
                $resp = Invoke-MgGraphRequest -Method $Method -Uri $next -OutputType PSObject -ErrorAction Stop
                break
            }
            catch {
                $code = $null
                try { $code = $_.Exception.Response.StatusCode.value__ } catch {}
                if (($code -eq 429 -or $code -eq 503) -and $attempt -lt $MaxRetries) {
                    $wait = 5
                    try { $wait = [int]$_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds } catch {}
                    if (-not $wait -or $wait -lt 1) { $wait = [math]::Pow(2, $attempt) }
                    Write-Warn2 "Throttled ($code). Waiting $wait s..."
                    Start-Sleep -Seconds $wait
                    $attempt++
                    continue
                }
                throw
            }
        }

        if ($resp.PSObject.Properties.Name -contains 'value') {
            foreach ($v in $resp.value) { $items.Add($v) }
            $next = $resp.'@odata.nextLink'
        }
        else {
            $items.Add($resp)   # single object
            $next = $null
        }
    } while ($next)

    return $items
}

function Connect-Tenant {
    $needed = @(
        'User.Read.All', 'Group.Read.All', 'GroupMember.Read.All',
        'UserAuthenticationMethod.Read.All', 'AuditLog.Read.All', 'Reports.Read.All',
        'Policy.Read.All', 'IdentityRiskyUser.Read.All', 'IdentityRiskEvent.Read.All',
        'RoleManagement.Read.Directory', 'RoleManagement.Read.All',
        'Directory.Read.All', 'Application.Read.All', 'Device.Read.All',
        'Organization.Read.All'
    )
    $ctx = $null
    try { $ctx = Get-MgContext } catch {}

    $haveAll = $ctx -and -not ($needed | Where-Object { $_ -notin $ctx.Scopes })
    if (-not $haveAll) {
        Write-Info "Connecting to Microsoft Graph (consent prompt may appear)..."
        Connect-MgGraph -Scopes $needed -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
    }
    Write-Good "Connected as $($ctx.Account)"
    Write-Host "      Tenant : $($ctx.TenantId)" -ForegroundColor DarkGray
    return $ctx
}

#endregion

#region ── Scope resolution (which users do reports run against?) ────────────────

function Resolve-Group {
    param([string]$Identity)
    if ($Identity -match '^[0-9a-fA-F-]{36}$') {
        $g = Invoke-Graph "groups/$Identity" 2>$null
        if ($g) { return $g[0] }
    }
    $esc = $Identity.Replace("'", "''")
    $g = Invoke-Graph "groups?`$filter=displayName eq '$esc'&`$select=id,displayName,groupTypes,mailEnabled,securityEnabled"
    if ($g.Count -ge 1) {
        if ($g.Count -gt 1) { Write-Warn2 "Multiple groups named '$Identity'; using first ($($g[0].id))." }
        return $g[0]
    }
    return $null
}

function Get-ScopedUsers {
    <# Returns a de-duplicated list of user objects based on -Groups / -Users / -AllUsers #>
    $sel = 'id,displayName,userPrincipalName,accountEnabled,userType,createdDateTime,onPremisesSyncEnabled'
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $out = [System.Collections.Generic.List[object]]::new()

    if ($AllUsers) {
        Write-Info "Scope: ALL users in tenant."
        foreach ($u in (Invoke-Graph "users?`$select=$sel&`$top=999")) {
            if ($seen.Add($u.id)) { $out.Add($u) }
        }
    }
    if ($Groups) {
        foreach ($gid in $Groups) {
            $g = Resolve-Group $gid
            if (-not $g) { Write-Warn2 "Group '$gid' not found."; continue }
            Write-Info "Scope: group '$($g.displayName)' (transitive members)."
            $uri = "groups/$($g.id)/transitiveMembers/microsoft.graph.user?`$select=$sel&`$top=999"
            foreach ($u in (Invoke-Graph $uri)) { if ($seen.Add($u.id)) { $out.Add($u) } }
        }
    }
    if ($Users) {
        foreach ($uid in $Users) {
            $u = if ($uid -match '@') {
                (Invoke-Graph "users?`$filter=userPrincipalName eq '$($uid.Replace("'","''"))'&`$select=$sel")[0]
            }
            else {
                (Invoke-Graph "users/$uid`?`$select=$sel")[0]
            }
            if ($u) { if ($seen.Add($u.id)) { $out.Add($u) } } else { Write-Warn2 "User '$uid' not found." }
        }
    }
    return $out
}

#endregion

#region ── Friendly value maps ──────────────────────────────────────────────────

$script:MethodFriendly = @{
    'microsoftAuthenticatorPush'         = 'Authenticator (push)'
    'microsoftAuthenticatorPasswordless' = 'Authenticator (passwordless)'
    'softwareOneTimePasscode'            = 'Software OATH (TOTP)'
    'mobilePhone'                        = 'Phone (SMS/call)'
    'alternateMobilePhone'               = 'Alt mobile phone'
    'officePhone'                        = 'Office phone'
    'email'                              = 'Email (SSPR)'
    'securityQuestion'                   = 'Security questions'
    'fido2SecurityKey'                   = 'FIDO2 key'
    'windowsHelloForBusiness'            = 'Windows Hello for Business'
    'temporaryAccessPass'                = 'Temporary Access Pass'
    'passKeyDeviceBound'                 = 'Passkey (device-bound)'
    'passKeyDeviceBoundAuthenticator'    = 'Passkey (Authenticator)'
    'hardwareOneTimePasscode'            = 'Hardware OATH token'
}
function Convert-Methods { param($arr) (($arr | ForEach-Object { if ($script:MethodFriendly[$_]) { $script:MethodFriendly[$_] } else { $_ } }) -join '; ') }

#endregion

#region ── REPORTS ──────────────────────────────────────────────────────────────

function Report-TenantOverview {
    Write-Sub "Tenant Overview"
    $org = (Invoke-Graph "organization")[0]
    $domains = Invoke-Graph "domains?`$select=id,isDefault,isVerified"
    $userCount = (Invoke-Graph "users/`$count" -) 2>$null
    # users/$count needs ConsistencyLevel header; fall back to length
    $skus = Invoke-Graph "subscribedSkus?`$select=skuPartNumber,prepaidUnits,consumedUnits"

    $secDefaults = $null
    try { $secDefaults = (Invoke-Graph "policies/identitySecurityDefaultsEnforcementPolicy")[0].isEnabled } catch {}

    $o = [PSCustomObject]@{
        Organization     = $org.displayName
        TenantId         = $org.id
        DefaultDomain    = ($domains | Where-Object isDefault).id
        VerifiedDomains  = ($domains | Where-Object isVerified).Count
        SecurityDefaults = if ($null -ne $secDefaults) { if ($secDefaults) { 'ENABLED' } else { 'Disabled' } } else { 'Unknown' }
        TechContact      = ($org.technicalNotificationMails -join '; ')
    }
    Write-Host ""
    $o.PSObject.Properties | ForEach-Object {
        $c = if ($_.Name -eq 'SecurityDefaults' -and $_.Value -eq 'Disabled') { 'Yellow' } else { 'White' }
        Write-Host ("    {0,-18}: {1}" -f $_.Name, $_.Value) -ForegroundColor $c
    }

    Write-Host "`n    Licenses (SKU consumed/total):" -ForegroundColor DarkCyan
    $skuList = foreach ($s in $skus) {
        $total = $s.prepaidUnits.enabled
        Write-Host ("      {0,-32} {1}/{2}" -f $s.skuPartNumber, $s.consumedUnits, $total) -ForegroundColor Gray
        [PSCustomObject]@{ Sku = $s.skuPartNumber; Consumed = $s.consumedUnits; Total = $total }
    }
    $script:Sections['TenantOverview'] = , $o + $skuList
    return , $o
}

function Report-RegistrationSummary {
    param($Users)
    Write-Sub "Authentication Registration Summary"
    $rows = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($u in $Users) {
        $i++
        Write-Host ("    [$i/$($Users.Count)] $($u.userPrincipalName)") -ForegroundColor DarkGray
        $d = (Invoke-Graph "reports/authenticationMethods/userRegistrationDetails/$($u.id)") 2>$null
        if (-not $d) { continue }
        $d = $d[0]
        $hasAuth = ($d.methodsRegistered -contains 'microsoftAuthenticatorPush') -or
        ($d.methodsRegistered -contains 'microsoftAuthenticatorPasswordless')
        $rows.Add([PSCustomObject]@{
                DisplayName            = $d.userDisplayName
                UserPrincipalName      = $d.userPrincipalName
                AccountEnabled         = $u.accountEnabled
                MfaCapable             = [bool]$d.isMfaCapable
                MfaRegistered          = [bool]$d.isMfaRegistered
                SsprCapable            = [bool]$d.isSsprCapable
                SsprRegistered         = [bool]$d.isSsprRegistered
                SsprEnabled            = [bool]$d.isSsprEnabled
                PasswordlessCapable    = [bool]$d.isPasswordlessCapable
                AuthenticatorApp       = $hasAuth
                SystemPreferredEnabled = [bool]$d.isSystemPreferredAuthenticationMethodEnabled
                DefaultMfaMethod       = $d.defaultMfaMethod
                MethodsRegistered      = (Convert-Methods $d.methodsRegistered)
            })
    }
    # quick console summary
    $t = $rows.Count
    if ($t) {
        Write-Host ""
        Write-Host ("    MFA capable        : {0}/{1}" -f ($rows | ? MfaCapable).Count, $t) -ForegroundColor $(if (($rows | ? { -not $_.MfaCapable })) { 'Yellow' }else { 'Green' })
        Write-Host ("    SSPR capable       : {0}/{1}" -f ($rows | ? SsprCapable).Count, $t) -ForegroundColor Cyan
        Write-Host ("    Authenticator app  : {0}/{1}" -f ($rows | ? AuthenticatorApp).Count, $t) -ForegroundColor Cyan
        Write-Host ("    NOT MFA capable    : {0}" -f ($rows | ? { -not $_.MfaCapable }).Count) -ForegroundColor $(if (($rows | ? { -not $_.MfaCapable })) { 'Red' }else { 'Green' })
    }
    $script:Sections['RegistrationSummary'] = $rows
    return $rows
}

function Report-AuthMethods {
    param($Users)
    Write-Sub "Detailed Registered Authentication Methods"
    $rows = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($u in $Users) {
        $i++
        Write-Host ("    [$i/$($Users.Count)] $($u.userPrincipalName)") -ForegroundColor DarkGray
        $methods = (Invoke-Graph "users/$($u.id)/authentication/methods") 2>$null
        if (-not $methods) { continue }
        $phones = @(); $emails = @(); $fido = @(); $whfb = @(); $tap = @(); $oath = @(); $authApp = @()
        foreach ($m in $methods) {
            switch -Wildcard ($m.'@odata.type') {
                '*phoneAuthenticationMethod*' { $phones += "$($m.phoneType):$($m.phoneNumber)" }
                '*emailAuthenticationMethod*' { $emails += $m.emailAddress }
                '*fido2AuthenticationMethod*' { $fido += $m.model }
                '*windowsHelloForBusinessAuthenticationMethod*' { $whfb += $m.displayName }
                '*temporaryAccessPassAuthenticationMethod*' { $tap += "active=$($m.isUsable)" }
                '*softwareOathAuthenticationMethod*' { $oath += 'TOTP' }
                '*microsoftAuthenticatorAuthenticationMethod*' { $authApp += $m.displayName }
            }
        }
        $rows.Add([PSCustomObject]@{
                DisplayName       = $u.displayName
                UserPrincipalName = $u.userPrincipalName
                AuthenticatorApps = ($authApp -join '; ')
                Phones            = ($phones -join '; ')
                Emails            = ($emails -join '; ')
                FIDO2Keys         = ($fido -join '; ')
                WindowsHello      = ($whfb -join '; ')
                SoftwareOATH      = ($oath -join '; ')
                TempAccessPass    = ($tap -join '; ')
                TotalMethods      = $methods.Count
            })
    }
    $script:Sections['AuthMethods'] = $rows
    return $rows
}

function Report-PerUserMfa {
    param($Users)
    Write-Sub "Legacy Per-User MFA State (disabled/enabled/enforced)"
    Write-Warn2 "Reads /authentication/requirements (beta). Prefer Conditional Access over per-user MFA."
    $rows = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($u in $Users) {
        $i++
        Write-Host ("    [$i/$($Users.Count)] $($u.userPrincipalName)") -ForegroundColor DarkGray
        $state = 'unknown'
        try {
            $r = Invoke-MgGraphRequest -Method GET -OutputType PSObject -ErrorAction Stop `
                -Uri "https://graph.microsoft.com/beta/users/$($u.id)/authentication/requirements"
            $state = $r.perUserMfaState
        }
        catch {}
        $rows.Add([PSCustomObject]@{
                DisplayName       = $u.displayName
                UserPrincipalName = $u.userPrincipalName
                PerUserMfaState   = $state
            })
    }
    $script:Sections['PerUserMfa'] = $rows
    return $rows
}

function Report-SignInActivity {
    param($Users)
    Write-Sub "Sign-in Activity (last interactive / non-interactive)"
    $rows = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($u in $Users) {
        $i++
        Write-Host ("    [$i/$($Users.Count)] $($u.userPrincipalName)") -ForegroundColor DarkGray
        $full = (Invoke-Graph "users/$($u.id)?`$select=signInActivity,displayName,userPrincipalName") 2>$null
        $sa = $full[0].signInActivity
        $rows.Add([PSCustomObject]@{
                DisplayName        = $u.displayName
                UserPrincipalName  = $u.userPrincipalName
                LastInteractive    = $sa.lastSignInDateTime
                LastNonInteractive = $sa.lastNonInteractiveSignInDateTime
            })
    }
    $script:Sections['SignInActivity'] = $rows
    return $rows
}

function Report-StaleAccounts {
    param($Users)
    Write-Sub "Stale Accounts (enabled, no interactive sign-in in $StaleDays days)"
    $cut = (Get-Date).AddDays(-$StaleDays)
    $rows = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($u in $Users) {
        $i++
        $full = (Invoke-Graph "users/$($u.id)?`$select=signInActivity,accountEnabled,displayName,userPrincipalName,createdDateTime") 2>$null
        $sa = $full[0].signInActivity
        $last = $sa.lastSignInDateTime
        $isStale = -not $last -or ([datetime]$last -lt $cut)
        if ($u.accountEnabled -and $isStale) {
            $days = if ($last) { [int]((Get-Date) - [datetime]$last).TotalDays } else { 'never' }
            Write-Host ("    [$i] STALE: $($u.userPrincipalName) (last: $(if($last){$last}else{'never'}))") -ForegroundColor Yellow
            $rows.Add([PSCustomObject]@{
                    DisplayName       = $u.displayName
                    UserPrincipalName = $u.userPrincipalName
                    AccountEnabled    = $u.accountEnabled
                    LastSignIn        = $last
                    DaysSinceSignIn   = $days
                    Created           = $full[0].createdDateTime
                })
        }
    }
    Write-Good "$($rows.Count) stale enabled account(s) found."
    $script:Sections['StaleAccounts'] = $rows
    return $rows
}

function Report-PasswordAudit {
    param($Users)
    Write-Sub "Password Audit (never-expires / last-set / sync state)"
    $rows = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($u in $Users) {
        $i++
        $full = (Invoke-Graph "users/$($u.id)?`$select=displayName,userPrincipalName,passwordPolicies,lastPasswordChangeDateTime,onPremisesSyncEnabled,accountEnabled") 2>$null
        $f = $full[0]
        $neverExpires = ($f.passwordPolicies -like '*DisablePasswordExpiration*')
        if ($neverExpires) { Write-Host ("    [$i] never-expires: $($u.userPrincipalName)") -ForegroundColor Yellow }
        $rows.Add([PSCustomObject]@{
                DisplayName          = $f.displayName
                UserPrincipalName    = $f.userPrincipalName
                AccountEnabled       = $f.accountEnabled
                PasswordNeverExpires = [bool]$neverExpires
                LastPasswordChange   = $f.lastPasswordChangeDateTime
                OnPremSynced         = [bool]$f.onPremisesSyncEnabled
                PasswordPolicies     = ($f.passwordPolicies)
            })
    }
    $script:Sections['PasswordAudit'] = $rows
    return $rows
}

function Report-RiskyUsers {
    Write-Sub "Identity Protection — Risky Users"
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        $ru = Invoke-Graph "identityProtection/riskyUsers?`$filter=riskState ne 'remediated' and riskState ne 'dismissed'&`$top=200"
        foreach ($r in $ru) {
            $col = switch ($r.riskLevel) { 'high' { 'Red' } 'medium' { 'Yellow' } default { 'Gray' } }
            Write-Host ("    [$($r.riskLevel)] $($r.userPrincipalName) — state:$($r.riskState)") -ForegroundColor $col
            $rows.Add([PSCustomObject]@{
                    DisplayName       = $r.userDisplayName
                    UserPrincipalName = $r.userPrincipalName
                    RiskLevel         = $r.riskLevel
                    RiskState         = $r.riskState
                    RiskDetail        = $r.riskDetail
                    LastUpdated       = $r.riskLastUpdatedDateTime
                })
        }
    }
    catch { Write-Bad "RiskyUsers requires Microsoft Entra ID P2: $_" }
    Write-Good "$($rows.Count) at-risk user(s)."
    $script:Sections['RiskyUsers'] = $rows
    return $rows
}

function Report-RiskDetections {
    Write-Sub "Identity Protection — Risk Detections (last $Days days)"
    $since = (Get-Date).AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        $rd = Invoke-Graph "identityProtection/riskDetections?`$filter=detectedDateTime ge $since&`$top=500"
        foreach ($r in $rd) {
            $rows.Add([PSCustomObject]@{
                    UserPrincipalName = $r.userPrincipalName
                    RiskEventType     = $r.riskEventType
                    RiskLevel         = $r.riskLevel
                    RiskState         = $r.riskState
                    Activity          = $r.activity
                    IpAddress         = $r.ipAddress
                    Location          = "$($r.location.city), $($r.location.countryOrRegion)"
                    DetectedDateTime  = $r.detectedDateTime
                })
        }
    }
    catch { Write-Bad "RiskDetections requires Microsoft Entra ID P2: $_" }
    Write-Good "$($rows.Count) detection(s)."
    $script:Sections['RiskDetections'] = $rows
    return $rows
}

function Report-RiskySignIns {
    Write-Sub "Risky Sign-ins (>= $RiskLevelFilter, last $Days days)"
    $since = (Get-Date).AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $levels = switch ($RiskLevelFilter) {
        'low' { @('low', 'medium', 'high') }
        'medium' { @('medium', 'high') }
        'high' { @('high') }
    }
    $orFilter = ($levels | ForEach-Object { "riskLevelDuringSignIn eq '$_'" }) -join ' or '
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        $si = Invoke-Graph "auditLogs/signIns?`$filter=createdDateTime ge $since and ($orFilter)&`$top=500"
        foreach ($s in $si) {
            $rows.Add([PSCustomObject]@{
                    UserPrincipalName = $s.userPrincipalName
                    App               = $s.appDisplayName
                    RiskLevel         = $s.riskLevelDuringSignIn
                    RiskState         = $s.riskState
                    Status            = $s.status.errorCode
                    IpAddress         = $s.ipAddress
                    Location          = "$($s.location.city), $($s.location.countryOrRegion)"
                    ClientApp         = $s.clientAppUsed
                    DateTime          = $s.createdDateTime
                })
        }
    }
    catch { Write-Bad "RiskySignIns query failed (P2 fields may be unavailable): $_" }
    Write-Good "$($rows.Count) risky sign-in(s)."
    $script:Sections['RiskySignIns'] = $rows
    return $rows
}

function Report-ConditionalAccess {
    Write-Sub "Conditional Access Policies"
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        $pol = Invoke-Graph "identity/conditionalAccess/policies"
        foreach ($p in $pol) {
            $col = switch ($p.state) { 'enabled' { 'Green' } 'enabledForReportingButNotEnforced' { 'Yellow' } default { 'DarkGray' } }
            Write-Host ("    [$($p.state)] $($p.displayName)") -ForegroundColor $col
            $rows.Add([PSCustomObject]@{
                    DisplayName   = $p.displayName
                    State         = $p.state
                    Users         = ($p.conditions.users.includeUsers -join ',')
                    Apps          = ($p.conditions.applications.includeApplications -join ',')
                    GrantControls = ($p.grantControls.builtInControls -join ',')
                    Created       = $p.createdDateTime
                    Modified      = $p.modifiedDateTime
                })
        }
    }
    catch { Write-Bad "ConditionalAccess read failed: $_" }
    Write-Good "$($rows.Count) policy(ies)."
    $script:Sections['ConditionalAccess'] = $rows
    return $rows
}

function Report-AdminRoles {
    Write-Sub "Directory Role Assignments (active + PIM-eligible)"
    $rows = [System.Collections.Generic.List[object]]::new()
    # Active assignments
    try {
        $roleDefs = @{}
        foreach ($rd in (Invoke-Graph "roleManagement/directory/roleDefinitions?`$select=id,displayName")) { $roleDefs[$rd.id] = $rd.displayName }
        $assign = Invoke-Graph "roleManagement/directory/roleAssignments?`$expand=principal"
        foreach ($a in $assign) {
            $rows.Add([PSCustomObject]@{
                    Role       = $roleDefs[$a.roleDefinitionId]
                    Principal  = $a.principal.userPrincipalName ?? $a.principal.displayName
                    Type       = ($a.principal.'@odata.type' -replace '#microsoft.graph.', '')
                    Assignment = 'Active'
                })
        }
    }
    catch { Write-Bad "Active role assignments failed: $_" }
    # PIM-eligible (P2)
    try {
        $elig = Invoke-Graph "roleManagement/directory/roleEligibilitySchedules?`$expand=principal,roleDefinition"
        foreach ($e in $elig) {
            $rows.Add([PSCustomObject]@{
                    Role       = $e.roleDefinition.displayName
                    Principal  = $e.principal.userPrincipalName ?? $e.principal.displayName
                    Type       = ($e.principal.'@odata.type' -replace '#microsoft.graph.', '')
                    Assignment = 'Eligible (PIM)'
                })
        }
    }
    catch { Write-Warn2 "PIM-eligible roles unavailable (needs Entra ID P2)." }

    $rows | Sort-Object Role | ForEach-Object {
        $c = if ($_.Role -match 'Global Admin') { 'Red' } else { 'Gray' }
        Write-Host ("    {0,-32} {1,-14} {2}" -f $_.Role, $_.Assignment, $_.Principal) -ForegroundColor $c
    }
    Write-Good "$($rows.Count) role assignment(s)."
    $script:Sections['AdminRoles'] = $rows
    return $rows
}

function Report-Licenses {
    param($Users)
    Write-Sub "Per-User License Assignment"
    $skuMap = @{}
    foreach ($s in (Invoke-Graph "subscribedSkus?`$select=skuId,skuPartNumber")) { $skuMap[$s.skuId] = $s.skuPartNumber }
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($u in $Users) {
        $full = (Invoke-Graph "users/$($u.id)?`$select=displayName,userPrincipalName,assignedLicenses") 2>$null
        $names = $full[0].assignedLicenses | ForEach-Object { $skuMap[$_.skuId] ?? $_.skuId }
        $rows.Add([PSCustomObject]@{
                DisplayName       = $u.displayName
                UserPrincipalName = $u.userPrincipalName
                Licenses          = ($names -join '; ')
                LicenseCount      = @($names).Count
            })
    }
    $script:Sections['Licenses'] = $rows
    return $rows
}

function Report-Devices {
    param($Users)
    Write-Sub "Registered / Owned Devices"
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($u in $Users) {
        $devs = (Invoke-Graph "users/$($u.id)/registeredDevices?`$select=displayName,operatingSystem,operatingSystemVersion,isCompliant,isManaged,trustType,approximateLastSignInDateTime") 2>$null
        foreach ($d in $devs) {
            $rows.Add([PSCustomObject]@{
                    Owner       = $u.userPrincipalName
                    DeviceName  = $d.displayName
                    OS          = "$($d.operatingSystem) $($d.operatingSystemVersion)"
                    TrustType   = $d.trustType
                    IsManaged   = [bool]$d.isManaged
                    IsCompliant = [bool]$d.isCompliant
                    LastSignIn  = $d.approximateLastSignInDateTime
                })
        }
    }
    Write-Good "$($rows.Count) device record(s)."
    $script:Sections['Devices'] = $rows
    return $rows
}

function Report-OAuthGrants {
    Write-Sub "Delegated OAuth2 Consent Grants (consent-phishing hunt)"
    $rows = [System.Collections.Generic.List[object]]::new()
    $spMap = @{}
    try {
        $grants = Invoke-Graph "oauth2PermissionGrants?`$top=500"
        foreach ($g in $grants) {
            if (-not $spMap.ContainsKey($g.clientId)) {
                $sp = (Invoke-Graph "servicePrincipals/$($g.clientId)?`$select=displayName,appId") 2>$null
                $spMap[$g.clientId] = if ($sp) { $sp[0].displayName } else { $g.clientId }
            }
            $rows.Add([PSCustomObject]@{
                    ClientApp   = $spMap[$g.clientId]
                    ConsentType = $g.consentType        # AllPrincipals = admin/tenant-wide
                    Scopes      = ($g.scope).Trim()
                    PrincipalId = $g.principalId
                })
        }
    }
    catch { Write-Bad "OAuthGrants read failed: $_" }
    # Highlight tenant-wide grants with sensitive scopes
    $rows | Where-Object { $_.ConsentType -eq 'AllPrincipals' -and $_.Scopes -match 'Mail|Files|offline_access|Directory' } |
    ForEach-Object { Write-Host ("    TENANT-WIDE: $($_.ClientApp) -> $($_.Scopes)") -ForegroundColor Yellow }
    Write-Good "$($rows.Count) grant(s)."
    $script:Sections['OAuthGrants'] = $rows
    return $rows
}

function Report-AppCredentials {
    Write-Sub "App Registration Secrets / Certificates + Expiry"
    $rows = [System.Collections.Generic.List[object]]::new()
    $now = Get-Date
    try {
        $apps = Invoke-Graph "applications?`$select=displayName,appId,passwordCredentials,keyCredentials&`$top=500"
        foreach ($a in $apps) {
            foreach ($pc in $a.passwordCredentials) {
                $exp = [datetime]$pc.endDateTime; $d = [int]($exp - $now).TotalDays
                if ($d -lt 30) { Write-Host ("    EXPIRING($d d): $($a.displayName) secret") -ForegroundColor Yellow }
                $rows.Add([PSCustomObject]@{ App = $a.displayName; CredType = 'Secret'; Name = $pc.displayName; Expires = $pc.endDateTime; DaysToExpiry = $d })
            }
            foreach ($kc in $a.keyCredentials) {
                $exp = [datetime]$kc.endDateTime; $d = [int]($exp - $now).TotalDays
                $rows.Add([PSCustomObject]@{ App = $a.displayName; CredType = 'Certificate'; Name = $kc.displayName; Expires = $kc.endDateTime; DaysToExpiry = $d })
            }
        }
    }
    catch { Write-Bad "AppCredentials read failed: $_" }
    Write-Good "$($rows.Count) credential(s)."
    $script:Sections['AppCredentials'] = $rows
    return $rows
}

function Report-AuditLogs {
    Write-Sub "Directory Audit Events (last $Days days)"
    $since = (Get-Date).AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        $logs = Invoke-Graph "auditLogs/directoryAudits?`$filter=activityDateTime ge $since&`$top=500"
        foreach ($l in $logs) {
            $rows.Add([PSCustomObject]@{
                    DateTime  = $l.activityDateTime
                    Category  = $l.category
                    Activity  = $l.activityDisplayName
                    Result    = $l.result
                    Initiator = $l.initiatedBy.user.userPrincipalName ?? $l.initiatedBy.app.displayName
                    Target    = ($l.targetResources.userPrincipalName -join ',')
                })
        }
    }
    catch { Write-Bad "AuditLogs read failed: $_" }
    Write-Good "$($rows.Count) audit event(s)."
    $script:Sections['AuditLogs'] = $rows
    return $rows
}

function Report-GroupMembership {
    param($Users)
    Write-Sub "Transitive Group Membership"
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($u in $Users) {
        $grps = (Invoke-Graph "users/$($u.id)/transitiveMemberOf/microsoft.graph.group?`$select=displayName") 2>$null
        $rows.Add([PSCustomObject]@{
                DisplayName       = $u.displayName
                UserPrincipalName = $u.userPrincipalName
                GroupCount        = @($grps).Count
                Groups            = (($grps.displayName) -join '; ')
            })
    }
    $script:Sections['GroupMembership'] = $rows
    return $rows
}

function Report-GroupDirectory {
    <#
    Lists every group in the tenant — Security, Microsoft 365 (Unified),
    Distribution, and Mail-enabled Security — with enough detail to copy-paste
    a name straight into -Groups on the next run.

    Columns
    -------
    GroupType        : Security | M365 | Distribution | MailEnabledSecurity
    DisplayName      : the name you pass to -Groups
    ObjectId         : alternative identifier for -Groups
    Email            : primary SMTP address (if mail-enabled)
    Description      : group description
    MembershipType   : Assigned | Dynamic (user) | Dynamic (device)
    MemberCount      : direct member count (omitted for very large groups)
    Owners           : display names of group owners (up to 5)
    OnPremSynced     : True = synced from on-prem AD
    HideFromGAL      : 'n/a' on v1.0 (requires -UseBeta flag)
    Created          : group creation date
    #>
    param(
        # Optional keyword to filter groups by display name (substring, case-insensitive)
        [string] $Filter
    )
    Write-Sub "Group Directory (Security / M365 / Distribution)"

    # renewedDateTime and hideFromAddressLists are not available on the v1.0
    # groups endpoint without Exchange extensions; omitting avoids a 405 error.
    $sel = 'id,displayName,description,groupTypes,mailEnabled,securityEnabled,' +
    'mail,membershipRule,membershipRuleProcessingState,onPremisesSyncEnabled,' +
    'createdDateTime'

    Write-Info "Fetching all groups (paged)..."
    $allGroups = Invoke-Graph "groups?`$select=$sel&`$top=999"
    Write-Good "$($allGroups.Count) group(s) returned from tenant."

    # Classify each group
    $rows = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($g in $allGroups) {
        $i++

        # ── Type classification ───────────────────────────────────────────────
        $type = if ($g.groupTypes -contains 'Unified') {
            'M365'
        }
        elseif ($g.mailEnabled -and $g.securityEnabled) {
            'MailEnabledSecurity'
        }
        elseif ($g.mailEnabled -and -not $g.securityEnabled) {
            'Distribution'
        }
        else {
            'Security'
        }

        # ── Membership type ───────────────────────────────────────────────────
        $membershipType = if ($g.membershipRule) {
            if ($g.membershipRule -match 'device\.') { 'Dynamic (device)' } else { 'Dynamic (user)' }
        }
        else { 'Assigned' }

        # Apply optional display-name filter
        if ($Filter -and $g.displayName -notmatch [regex]::Escape($Filter)) { continue }

        # ── Member count (best-effort; Graph requires $count header) ──────────
        $memberCount = '?'
        try {
            $countUri = "https://graph.microsoft.com/v1.0/groups/$($g.id)/members/`$count"
            $cr = Invoke-MgGraphRequest -Method GET -Uri $countUri -OutputType PSObject `
                -Headers @{ ConsistencyLevel = 'eventual' } -ErrorAction Stop
            $memberCount = $cr
        }
        catch { }

        # ── Owners (up to 5) ─────────────────────────────────────────────────
        $ownerNames = ''
        try {
            $owners = Invoke-Graph "groups/$($g.id)/owners?`$select=displayName&`$top=5"
            $ownerNames = ($owners.displayName -join '; ')
        }
        catch { }

        # ── Console line (color-coded by type) ───────────────────────────────
        $typeColor = switch ($type) {
            'M365' { 'Cyan' }
            'Security' { 'Green' }
            'Distribution' { 'Yellow' }
            'MailEnabledSecurity' { 'Magenta' }
        }
        Write-Host ("    [{0,-20}] {1,-45} members:{2,-5} {3}" -f `
                $type, $g.displayName, $memberCount,
            $(if ($membershipType -ne 'Assigned') { "[$membershipType]" } else { '' })
        ) -ForegroundColor $typeColor

        $rows.Add([PSCustomObject]@{
                GroupType      = $type
                DisplayName    = $g.displayName
                ObjectId       = $g.id
                Email          = $g.mail
                Description    = $g.description
                MembershipType = $membershipType
                MemberCount    = $memberCount
                Owners         = $ownerNames
                OnPremSynced   = [bool]$g.onPremisesSyncEnabled
                HideFromGAL    = 'n/a (requires beta endpoint)'
                Created        = $g.createdDateTime
            })
    }

    # ── Summary by type ───────────────────────────────────────────────────────
    Write-Host ""
    $types = $rows | Group-Object GroupType | Sort-Object Name
    foreach ($t in $types) {
        $c = switch ($t.Name) {
            'M365' { 'Cyan' }
            'Security' { 'Green' }
            'Distribution' { 'Yellow' }
            'MailEnabledSecurity' { 'Magenta' }
        }
        Write-Host ("    {0,-22}: {1}" -f $t.Name, $t.Count) -ForegroundColor $c
    }

    # ── Interactive scope launcher ────────────────────────────────────────────
    Write-Host ""
    $pick = Read-Host "  Copy a group name above, or press ENTER to continue"
    if ($pick.Trim()) {
        Write-Good "Tip: re-run with:  -Groups `"$($pick.Trim())`""
    }

    $script:Sections['GroupDirectory'] = $rows
    return $rows
}

#endregion

#region ── HTML dashboard ───────────────────────────────────────────────────────

function Export-HtmlDashboard {
    param([string]$Path, $Context)
    $css = @'
body{font-family:Segoe UI,Arial,sans-serif;background:#0f1419;color:#e6e6e6;margin:0;padding:24px}
h1{color:#4fc3f7;border-bottom:2px solid #263238;padding-bottom:8px}
h2{color:#81d4fa;margin-top:36px;cursor:pointer}
.meta{color:#90a4ae;font-size:13px;margin-bottom:24px}
table{border-collapse:collapse;width:100%;margin:8px 0 24px;font-size:13px}
th{background:#1c2733;color:#4fc3f7;text-align:left;padding:8px 10px;position:sticky;top:0}
td{border-bottom:1px solid #1c2733;padding:6px 10px;vertical-align:top}
tr:hover{background:#16202b}
.t{color:#66bb6a;font-weight:600}.f{color:#ef5350;font-weight:600}
.pill{padding:2px 8px;border-radius:10px;font-size:11px}
.warn{background:#5d4037;color:#ffcc80}.ok{background:#1b3a26;color:#a5d6a7}
.count{color:#90a4ae;font-weight:normal;font-size:14px}
'@
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8'><title>Entra Audit</title><style>$css</style></head><body>")
    [void]$sb.AppendLine("<h1>Entra ID Audit Dashboard</h1>")
    [void]$sb.AppendLine("<div class='meta'>Tenant $($Context.TenantId) &nbsp;·&nbsp; Account $($Context.Account) &nbsp;·&nbsp; Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>")

    foreach ($name in $script:Sections.Keys) {
        $data = $script:Sections[$name]
        if (-not $data -or @($data).Count -eq 0) { continue }
        [void]$sb.AppendLine("<h2>$name <span class='count'>($(@($data).Count))</span></h2>")
        $cols = $data[0].PSObject.Properties.Name
        [void]$sb.AppendLine("<table><thead><tr>")
        foreach ($c in $cols) { [void]$sb.AppendLine("<th>$c</th>") }
        [void]$sb.AppendLine("</tr></thead><tbody>")
        foreach ($row in $data) {
            [void]$sb.AppendLine("<tr>")
            foreach ($c in $cols) {
                $v = $row.$c
                $cell = if ($v -is [bool]) {
                    if ($v) { "<span class='t'>True</span>" } else { "<span class='f'>False</span>" }
                }
                else { [System.Web.HttpUtility]::HtmlEncode([string]$v) }
                [void]$sb.AppendLine("<td>$cell</td>")
            }
            [void]$sb.AppendLine("</tr>")
        }
        [void]$sb.AppendLine("</tbody></table>")
    }
    [void]$sb.AppendLine("<script>document.querySelectorAll('h2').forEach(h=>h.onclick=()=>{let t=h.nextElementSibling;t.style.display=t.style.display=='none'?'':'none'})</script>")
    [void]$sb.AppendLine("</body></html>")
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    $sb.ToString() | Out-File -FilePath $Path -Encoding UTF8
}

#endregion

#region ── Menu + dispatcher ────────────────────────────────────────────────────

$AllReports = @(
    'TenantOverview', 'GroupDirectory',
    'RegistrationSummary', 'AuthMethods', 'PerUserMfa', 'SignInActivity',
    'StaleAccounts', 'PasswordAudit', 'RiskyUsers', 'RiskDetections', 'RiskySignIns',
    'ConditionalAccess', 'AdminRoles', 'Licenses', 'Devices', 'OAuthGrants',
    'AppCredentials', 'AuditLogs', 'GroupMembership'
)
# reports that need a scoped user list
$UserScopedReports = @('RegistrationSummary', 'AuthMethods', 'PerUserMfa', 'SignInActivity',
    'StaleAccounts', 'PasswordAudit', 'Licenses', 'Devices', 'GroupMembership')

function Show-Menu {
    Write-Banner "Entra ID Audit Toolkit — Interactive Menu" Magenta
    for ($i = 0; $i -lt $AllReports.Count; $i++) {
        Write-Host ("   {0,2}. {1}" -f ($i + 1), $AllReports[$i]) -ForegroundColor White
    }
    Write-Host "    A. Run ALL reports" -ForegroundColor Green
    Write-Host "    Q. Quit" -ForegroundColor DarkGray
    Write-Host ""
    $sel = Read-Host "  Select report number(s) (comma-separated), A, or Q"
    if ($sel -match '^[Qq]') { return @() }
    if ($sel -match '^[Aa]') { return $AllReports }
    $chosen = @()
    foreach ($n in ($sel -split '[, ]+')) {
        if ($n -match '^\d+$' -and [int]$n -ge 1 -and [int]$n -le $AllReports.Count) {
            $chosen += $AllReports[[int]$n - 1]
        }
    }
    return $chosen
}

# ── Entry point ─────────────────────────────────────────────────────────────────
Write-Banner "ENTRA ID AUDIT TOOLKIT"
$ctx = Connect-Tenant

# Determine which reports to run
$toRun = if ($Report) {
    if ($Report -contains 'All') { $AllReports } else { $Report }
}
else {
    Show-Menu
}
if (-not $toRun -or $toRun.Count -eq 0) { Write-Info "Nothing to run. Bye."; Disconnect-MgGraph | Out-Null; return }

# If interactive and no scope given, but a user-scoped report is selected, ask for scope
$needsUsers = $toRun | Where-Object { $_ -in $UserScopedReports }
if ($needsUsers -and -not ($Groups -or $Users -or $AllUsers)) {
    Write-Warn2 "Selected report(s) need a user scope: $($needsUsers -join ', ')"
    Write-Info  "Tip: run GroupDirectory first to browse available group names."
    $scopeSel = Read-Host "  Enter group name(s)/UPN(s) comma-separated, or '*' for ALL users"
    if ($scopeSel -eq '*') { $AllUsers = $true }
    elseif ($scopeSel) {
        $parts = $scopeSel -split '\s*,\s*'
        $Users = $parts | Where-Object { $_ -match '@' }
        $Groups = $parts | Where-Object { $_ -notmatch '@' }
    }
}

# Resolve scoped users once (shared by all user-level reports)
$scopedUsers = @()
if ($needsUsers) {
    Write-Banner "Resolving user scope"
    $scopedUsers = Get-ScopedUsers
    Write-Good "$($scopedUsers.Count) user(s) in scope."
}

# Prepare output
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

# Run
foreach ($r in $toRun) {
    Write-Banner "REPORT: $r"
    $result = switch ($r) {
        'TenantOverview' { Report-TenantOverview }
        'GroupDirectory' { Report-GroupDirectory }
        'RegistrationSummary' { Report-RegistrationSummary -Users $scopedUsers }
        'AuthMethods' { Report-AuthMethods       -Users $scopedUsers }
        'PerUserMfa' { Report-PerUserMfa        -Users $scopedUsers }
        'SignInActivity' { Report-SignInActivity    -Users $scopedUsers }
        'StaleAccounts' { Report-StaleAccounts     -Users $scopedUsers }
        'PasswordAudit' { Report-PasswordAudit     -Users $scopedUsers }
        'RiskyUsers' { Report-RiskyUsers }
        'RiskDetections' { Report-RiskDetections }
        'RiskySignIns' { Report-RiskySignIns }
        'ConditionalAccess' { Report-ConditionalAccess }
        'AdminRoles' { Report-AdminRoles }
        'Licenses' { Report-Licenses          -Users $scopedUsers }
        'Devices' { Report-Devices           -Users $scopedUsers }
        'OAuthGrants' { Report-OAuthGrants }
        'AppCredentials' { Report-AppCredentials }
        'AuditLogs' { Report-AuditLogs }
        'GroupMembership' { Report-GroupMembership   -Users $scopedUsers }
    }
    if ($result -and @($result).Count -gt 0) {
        $csv = Join-Path $OutputFolder "$r.csv"
        $result | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
        Write-Good "CSV: $csv"
    }
    else {
        Write-Info "No rows for $r."
    }
}

# HTML dashboard
if (-not $NoHtml -and $script:Sections.Count -gt 0) {
    $html = Join-Path $OutputFolder "Dashboard.html"
    Export-HtmlDashboard -Path $html -Context $ctx
    Write-Banner "Output" Green
    Write-Good "HTML dashboard: $html"
    try { Invoke-Item $html } catch {}
}

Write-Good "All requested reports complete. Output folder: $OutputFolder"
Disconnect-MgGraph | Out-Null
Write-Host ""

#endregion