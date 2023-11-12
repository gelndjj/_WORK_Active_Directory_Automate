param (
    [string]$customIdentifier,  # Custom Identifier for the user whose password will be changed
    [string]$newPassword       # The new password to set
)

# Check if the customIdentifier and new password are provided
if (-not $customIdentifier -or -not $newPassword) {
    Write-Host "Error: Please provide a Custom Identifier and a new password."
    exit 1
}

# Search for the user in Active Directory using the Custom Identifier
$user = Get-AdUser -Filter "msDS-cloudExtensionAttribute1 -eq '$customIdentifier'"

# Check if the user with the given Custom Identifier exists
if ($user -eq $null) {
    Write-Host "Error: User with Custom Identifier '$customIdentifier' not found."
    exit 1
}

# Convert the plain text password to a secure string
$securePassword = ConvertTo-SecureString $newPassword -AsPlainText -Force

# Attempt to set the new password for the user
try {
    Set-AdAccountPassword -Identity $user -NewPassword $securePassword -Reset
    Write-Host "Password for $customIdentifier has been successfully changed."
} catch {
    Write-Host "Error: Password change failed. $_"
    exit 1
}
