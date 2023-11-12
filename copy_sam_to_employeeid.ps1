param (
    [string]$customIdentifier # Custom Identifier for the user
)

# Check if the customIdentifier is provided
if (-not $customIdentifier) {
    Write-Host "Error: Please provide a Custom Identifier."
    exit 1
}

# Search for the user in Active Directory using the Custom Identifier
$user = Get-AdUser -Filter "msDS-cloudExtensionAttribute1 -eq '$customIdentifier'" -Properties SAMAccountName, employeeID

# Check if the user with the given Custom Identifier exists
if ($user -eq $null) {
    Write-Host "Error: User with Custom Identifier '$customIdentifier' not found."
    exit 1
}

# Copy SAMAccountName to employeeID
try {
    Set-AdUser -Identity $user -Replace @{employeeID = $user.SAMAccountName}
    Write-Host "SAMAccountName has been successfully copied to employeeID for user '$customIdentifier'."
} catch {
    Write-Host "Error: Failed to copy SAMAccountName to employeeID. $_"
    exit 1
}
