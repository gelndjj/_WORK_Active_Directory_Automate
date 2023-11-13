# Declare a parameter for the script to accept the OUPath
param(
    [string]$OUPath
)
# Ensure you have the ActiveDirectory module installed.
Import-Module ActiveDirectory

# Check if the parameter is provided, else use a default value
if (-not $OUPath) {
    $OUPath = "OU=Users,OU=Administration,DC=mybusiness,DC=local"
}

# Query users from the given OU
$users = Get-ADUser -Filter * -SearchBase $OUPath -Property SAMAccountName, UserPrincipalName, StreetAddress, City, State, PostalCode, Company, Title, Department, Description, EmailAddress

# Convert the user data to the desired format
$exportData = $users | Select-Object @{Name="ID"; Expression={""}},
                                        @{Name="Last Name"; Expression={if ($_.Surname) {$_.Surname} else {""}}},
                                        @{Name="First Name"; Expression={if ($_.GivenName) {$_.GivenName} else {""}}},
                                        @{Name="Display Name"; Expression={if ($_.Name) {$_.Name} else {""}}},
                                        @{Name="Description"; Expression={if ($_.Description) {$_.Description} else {""}}},
                                        @{Name="Email"; Expression={if ($_.EmailAddress) {$_.EmailAddress} else {""}}},
                                        @{Name="SAM"; Expression={if ($_.SAMAccountName) {$_.SAMAccountName} else {""}}},
                                        @{Name="UPN"; Expression={if ($_.UserPrincipalName) {$_.UserPrincipalName} else {""}}},
                                        @{Name="Address"; Expression={if ($_.StreetAddress) {$_.StreetAddress} else {""}}},
                                        @{Name="City"; Expression={if ($_.City) {$_.City} else {""}}},
                                        @{Name="State"; Expression={if ($_.State) {$_.State} else {""}}},
                                        @{Name="Zipcode"; Expression={if ($_.PostalCode) {$_.PostalCode} else {""}}},
                                        @{Name="Company"; Expression={if ($_.Company) {$_.Company} else {""}}},
                                        @{Name="Job Title"; Expression={if ($_.Title) {$_.Title} else {""}}},
                                        @{Name="Department"; Expression={if ($_.Department) {$_.Department} else {""}}},
                                        @{Name="Custom Identifier"; Expression={"Placeholder"}}

# Define the filename with a timestamp
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$filename = "$PSScriptRoot\Extract_OU_Users_$timestamp.csv"

# Export data to CSV
$exportData | Export-Csv -NoTypeInformation -Path $filename
