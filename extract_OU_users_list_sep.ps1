# Import the Active Directory module
Import-Module ActiveDirectory

# Retrieve all users
$users = Get-ADUser -Filter *
$ousWithUsers = @{}

foreach ($user in $users) {
    $ou = ($user.DistinguishedName -split ",",2)[1]
    $ousWithUsers[$ou] = $true
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

    # Write just the OU path to the file
    $ou | Out-File -FilePath $csvPath
    Write-Host "Exported OU path $ou to $csvPath"
}
