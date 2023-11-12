param (
    [string]$customIdentifier # Custom Identifier for the user whose account status will be toggled
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

# Check if the account is enabled or disabled
if ($user.Enabled -eq $true) {
    # Disable the account if it is currently enabled
    try {
        Disable-ADAccount -Identity $user
        Write-Host "Account for user with Custom Identifier '$customIdentifier' has been successfully disabled."
    } catch {
        Write-Host "Error: Failed to disable the account. $_"
        exit 1
    }
} else {
    # Enable the account if it is currently disabled
    try {
        Enable-ADAccount -Identity $user
        Write-Host "Account for user with Custom Identifier '$customIdentifier' has been successfully enabled."
    } catch {
        Write-Host "Error: Failed to enable the account. $_"
        exit 1
    }
}
