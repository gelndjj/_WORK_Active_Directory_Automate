# Import the Active Directory module
Import-Module ActiveDirectory

# Retrieve all users
$users = Get-ADUser -Filter *

# Loop through each user to find their OU
$ousWithUsers = @{}
foreach ($user in $users) {
    $ou = ($user.DistinguishedName -split ",",2)[1]
    if (-not $ousWithUsers.ContainsKey($ou)) {
        $ousWithUsers[$ou] = @()
    }
    $ousWithUsers[$ou] += $user
}

# Define the base directory to store the CSV files
$baseDir = "C:\temp"

# Ensure the base directory exists
if (-not (Test-Path -Path $baseDir)) {
    New-Item -ItemType Directory -Path $baseDir
    Write-Host "Created base directory: $baseDir"
}

# Process each OU and create a CSV file
foreach ($ou in $ousWithUsers.Keys) {
    $splitOU = $ou -split ','
    if ($splitOU.Length -lt 2) { continue }
    $csvNameParts = $splitOU[0..1] -replace 'OU=', '' -replace 'DC=', ''
    $csvName = $csvNameParts[1] + '_' + $csvNameParts[0] # Second-to-last and last parts
    $csvPath = Join-Path -Path $baseDir -ChildPath "$csvName.csv"

    # Export users in the OU to a CSV file
    $ousWithUsers[$ou] | Select-Object Name, SamAccountName, DistinguishedName | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Exported OU $ou to $csvPath"
}
