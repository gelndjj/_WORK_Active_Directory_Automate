param (
    [string]$customIdentifier # Custom Identifier for the user
)

# Check if the customIdentifier is provided
if (-not $customIdentifier) {
    Write-Host "Error: Please provide a Custom Identifier."
    exit 1
}

# Search for the user in Active Directory using the Custom Identifier
$user = Get-AdUser -Filter "msDS-cloudExtensionAttribute1 -eq '$customIdentifier'"

# Check if the user with the given Custom Identifier exists
if ($user -eq $null) {
    Write-Host "Error: User with Custom Identifier '$customIdentifier' not found."
    exit 1
}

# Erase employeeID
try {
    Set-AdUser -Identity $user -Clear employeeID
    Write-Host "employeeID has been successfully erased for user '$customIdentifier'."
} catch {
    Write-Host "Error: Failed to erase employeeID. $_"
    exit 1
}
