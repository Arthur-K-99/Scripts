#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Adds users from an Active Directory OU to a security group.

.DESCRIPTION
    Resolves the specified OU and security group, finds users in the OU, and adds each user
    who is not already a direct member of the group. The operation is additive only: users
    who are already in the group are skipped, and no existing group members are removed.

    Child OUs are included by default. Use -SearchScope OneLevel to process only users
    directly in the specified OU. Use -WhatIf to preview changes.

.PARAMETER UserOU
    The distinguished name (DN) of the OU containing the users.

.PARAMETER SecurityGroup
    The identity of the target AD security group. Accepted values include its distinguished
    name, object GUID, object SID, or sAMAccountName.

.PARAMETER SearchScope
    Specifies whether to include child OUs. Subtree is the default; OneLevel processes only
    users directly in the specified OU.

.PARAMETER EnabledOnly
    Processes only enabled user accounts. By default, both enabled and disabled users are
    processed.

.PARAMETER Server
    Optional domain controller or AD LDS instance to use for every directory operation.

.PARAMETER Credential
    Optional credential to use for every directory operation.

.EXAMPLE
    .\Add-OUUsersToSecurityGroup.ps1 `
        -UserOU "OU=Employees,DC=contoso,DC=com" `
        -SecurityGroup "Employees-App-Access" `
        -WhatIf

    Previews the users that would be added, including users in child OUs.

.EXAMPLE
    .\Add-OUUsersToSecurityGroup.ps1 `
        -UserOU "OU=Employees,DC=contoso,DC=com" `
        -SecurityGroup "Employees-App-Access"

    Adds all users in the OU and its child OUs to the group.

.EXAMPLE
    .\Add-OUUsersToSecurityGroup.ps1 `
        -UserOU "OU=Employees,DC=contoso,DC=com" `
        -SecurityGroup "Employees-App-Access" `
        -SearchScope OneLevel `
        -EnabledOnly

    Adds only enabled users located directly in the specified OU.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$UserOU,

    [Parameter(Mandatory, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$SecurityGroup,

    [ValidateSet('OneLevel', 'Subtree')]
    [string]$SearchScope = 'Subtree',

    [switch]$EnabledOnly,

    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [PSCredential]$Credential
)

$directoryParameters = @{
    ErrorAction = 'Stop'
}

if ($Server) {
    $directoryParameters['Server'] = $Server
}

if ($Credential) {
    $directoryParameters['Credential'] = $Credential
}

try {
    $organizationalUnit = Get-ADOrganizationalUnit `
        -Identity $UserOU `
        @directoryParameters
}
catch {
    Write-Error "Could not resolve OU '$UserOU': $($_.Exception.Message)"
    return
}

try {
    $group = Get-ADGroup `
        -Identity $SecurityGroup `
        -Properties GroupCategory `
        @directoryParameters
}
catch {
    Write-Error "Could not resolve group '$SecurityGroup': $($_.Exception.Message)"
    return
}

if ($group.GroupCategory -ne 'Security') {
    Write-Error "Group '$($group.Name)' is a $($group.GroupCategory) group, not a security group."
    return
}

$userFilter = if ($EnabledOnly) {
    'Enabled -eq $true'
}
else {
    '*'
}

try {
    $users = @(
        Get-ADUser `
            -Filter $userFilter `
            -SearchBase $organizationalUnit.DistinguishedName `
            -SearchScope $SearchScope `
            -Properties DisplayName, Enabled `
            @directoryParameters |
            Sort-Object SamAccountName
    )
}
catch {
    Write-Error "Could not query users in OU '$($organizationalUnit.DistinguishedName)': $($_.Exception.Message)"
    return
}

if ($users.Count -eq 0) {
    Write-Warning "No matching users were found in OU '$($organizationalUnit.DistinguishedName)'."
    return
}

try {
    $directMembers = @(
        Get-ADGroupMember `
            -Identity $group `
            @directoryParameters
    )
}
catch {
    Write-Error "Could not read the current members of group '$($group.Name)': $($_.Exception.Message)"
    return
}

# PowerShell hashtables use case-insensitive string keys by default, which is suitable for DNs.
$directMemberDistinguishedNames = @{}
foreach ($member in $directMembers) {
    $directMemberDistinguishedNames[$member.DistinguishedName] = $true
}

Write-Verbose "Found $($users.Count) matching user(s) in '$($organizationalUnit.DistinguishedName)'."
Write-Verbose "Target security group: $($group.Name) [$($group.DistinguishedName)]"

$results = foreach ($user in $users) {
    $status = $null
    $message = $null

    if ($directMemberDistinguishedNames.ContainsKey($user.DistinguishedName)) {
        $status = 'AlreadyMember'
    }
    else {
        $target = "$($user.SamAccountName) -> $($group.Name)"

        if ($PSCmdlet.ShouldProcess($target, 'Add user to security group')) {
            try {
                Add-ADGroupMember `
                    -Identity $group `
                    -Members $user `
                    -Confirm:$false `
                    @directoryParameters

                $status = 'Added'
                $directMemberDistinguishedNames[$user.DistinguishedName] = $true
            }
            catch {
                $status = 'Failed'
                $message = $_.Exception.Message
                Write-Warning "Could not add '$($user.SamAccountName)' to '$($group.Name)': $message"
            }
        }
        elseif ($WhatIfPreference) {
            $status = 'WhatIf'
        }
        else {
            $status = 'Skipped'
        }
    }

    [PSCustomObject]@{
        SamAccountName    = $user.SamAccountName
        DisplayName       = $user.DisplayName
        Enabled           = $user.Enabled
        UserOU            = $organizationalUnit.DistinguishedName
        SecurityGroup     = $group.Name
        GroupScope        = $group.GroupScope
        Status            = $status
        Message           = $message
        DistinguishedName = $user.DistinguishedName
    }
}

Write-Information "`nSummary for security group '$($group.Name)':" -InformationAction Continue
$results |
    Group-Object Status |
    Sort-Object Name |
    ForEach-Object {
        Write-Information "  $($_.Name): $($_.Count)" -InformationAction Continue
    }

$results
