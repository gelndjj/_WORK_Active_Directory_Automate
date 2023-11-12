# Import the Active Directory module
Import-Module ActiveDirectory

# Retrieve all users
$users = Get-ADUser -Filter *

# Loop through each user to find their OU
$ousWithUsers = @{}
foreach ($user in $users) {
    $ou = ($user.DistinguishedName -split ",",2)[1]
    if (-not $ousWithUsers.ContainsKey($ou)) {
        $ousWithUsers[$ou] = $true
    }
}

# Create a list to store the OUs for CSV export
$ouList = @()
foreach ($ou in $ousWithUsers.Keys) {
    $ouList += [PSCustomObject]@{
        OrganizationalUnit = $ou
    }
}

# Define the CSV file path
$csvPath = "C:\temp\OU_Users_list.csv"

# Check if the directory exists, create it if it doesn't
$dirPath = Split-Path -Path $csvPath
if (-not (Test-Path -Path $dirPath)) {
    New-Item -ItemType Directory -Path $dirPath
    Write-Host "Created directory: $dirPath"
}

# Export the list of OUs to a CSV file
$ouList | Export-Csv -Path $csvPath -NoTypeInformation

Write-Host "Export complete. File saved to $csvPath"
