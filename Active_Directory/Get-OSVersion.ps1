# Import AD Module
Import-Module ActiveDirectory

# Define your target OU DN (Replace with your actual OU path)
$TargetOU = "OU=Windows 11,OU=UHC Computers,DC=uhc-nyc,DC=org"

# Fetch all computers and display their versions
Get-ADComputer -Filter * -SearchBase $TargetOU -Properties OperatingSystem, OperatingSystemVersion |
Select-Object Name, OperatingSystem, OperatingSystemVersion |
Sort-Object OperatingSystemVersion |
Out-GridView -Title "OS Version Audit"