param (
    [string]$customIdentifier # Custom Identifier for the user
)

# Check if the customIdentifier is provided
if (-not $customIdentifier) {
    Write-Host "Error: Please provide a Custom Identifier."
    exit 1
}

# Search for the user in Active Directory using the Custom Identifier
$user = Get-AdUser -Filter "msDS-cloudExtensionAttribute1 -eq '$customIdentifier'" -Properties DistinguishedName

# Check if the user with the given Custom Identifier exists
if ($user -eq $null) {
    Write-Host "Error: User with Custom Identifier '$customIdentifier' not found."
    exit 1
}

# Extract OU from the Distinguished Name
$ou = $user.DistinguishedName -replace '^.*?,(OU=.*)$', '$1'
Write-Host $ou
