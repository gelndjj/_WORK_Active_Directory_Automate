param(
    [string]$customIdentifier
)

# Import Active Directory Module
Import-Module ActiveDirectory

try {
    # Get the user with the specified custom identifier in msDS-cloudExtensionAttribute1
    $user = Get-ADUser -Filter "msDS-cloudExtensionAttribute1 -eq '$customIdentifier'"

    # Check if user exists
    if ($user -eq $null) {
        Write-Error "User with msDS-cloudExtensionAttribute1 '$customIdentifier' not found."
        exit 1
    }

    # Remove the msDS-cloudExtensionAttribute1 attribute
    Set-ADUser -Identity $user -Clear msDS-cloudExtensionAttribute1

    Write-Output "msDS-cloudExtensionAttribute1 removed for user: $($user.SamAccountName)"
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}
