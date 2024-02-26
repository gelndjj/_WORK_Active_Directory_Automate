param (
    [string[]]$customIdentifiers, # Array of Custom Identifiers for the users
    [string]$csvFilePath # Path to the CSV file with group names
)

# Check if the customIdentifiers and csvFilePath are provided
if (-not $customIdentifiers -or $customIdentifiers.Count -eq 0 -or -not $csvFilePath) {
    Write-Host "Error: Please provide at least one Custom Identifier and a CSV File Path."
    exit 1
}

# Read group names from the CSV file (without headers)
$groupNames = Get-Content -Path $csvFilePath

foreach ($customIdentifier in $customIdentifiers) {
    # Search for the user in Active Directory using the Custom Identifier
    $user = Get-AdUser -Filter "msDS-cloudExtensionAttribute1 -eq '$customIdentifier'" -Properties SAMAccountName

    # Check if the user with the given Custom Identifier exists
    if ($user -eq $null) {
        Write-Host "Error: User with Custom Identifier '$customIdentifier' not found."
        continue
    }

    foreach ($groupName in $groupNames) {
        # Attempt to find the group by name to ensure it exists and is unique
        $group = Get-AdGroup -Filter "Name -eq '$groupName'"

        if ($group -eq $null) {
            Write-Host "Error: Group '$groupName' not found."
            continue
        }

        try {
            Add-AdGroupMember -Identity $group.DistinguishedName -Members $user.SAMAccountName
            Write-Host "User '$customIdentifier' has been successfully added to group '$groupName'."
        } catch {
            Write-Host "Error: Failed to add user '$customIdentifier' to group '$groupName'. $_"
        }
    }
}
