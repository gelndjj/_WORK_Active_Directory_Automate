param(
    [Parameter(Mandatory=$true)]
    [string]$OuPath,

    [Parameter(Mandatory=$true)]
    [string]$DatabaseName
)

Write-Host "OU Path from argument: $OuPath"
Write-Host "Database Name from argument: $DatabaseName"

# Define the path to the SQLite module folder
$SQLiteModulePath = Join-Path (Split-Path -Path $MyInvocation.MyCommand.Path) "SQLite"

# Import the Active Directory module
Import-Module ActiveDirectory

# Load the SQLite module using Import-Module with Force parameter
Import-Module -Name $SQLiteModulePath -Force

# Define the path to the SQLite database file (use provided database name)
$DatabasePath = Join-Path (Split-Path -Path $MyInvocation.MyCommand.Path) $DatabaseName

# Ensure the assembly is loaded from the module folder
[Reflection.Assembly]::LoadFrom("$SQLiteModulePath\2.0\bin\x64\System.Data.SQLite.dll")

# Create a SQLite connection
$Connection = New-Object System.Data.SQLite.SQLiteConnection
$Connection.ConnectionString = "Data Source=$DatabasePath;Version=3;"

# Open the connection
$Connection.Open()

# Define the SQL query to retrieve user data from the database
$Query = "SELECT last_name, first_name, display_name, description, email, sam, upn, address, city, state, zipcode, company, job_title, department, cid as CustomIdentifier FROM customers"

# Create a SQLite command
$Command = $Connection.CreateCommand()
$Command.CommandText = $Query

# Execute the query and retrieve the data
$DataReader = $Command.ExecuteReader()

# Loop through the data and add/update users in Active Directory
while ($DataReader.Read()) {
    $CustomIdentifier = $DataReader["CustomIdentifier"]

    # Attempt to retrieve the user with the same CustomIdentifier from AD
    try {
        $ADUser = Get-ADUser -Filter {msDS-cloudExtensionAttribute1 -eq $CustomIdentifier} -ErrorAction Stop
    } catch {
        $ADUser = $null
    }

    # If the user exists, update them
    if ($ADUser) {
        $ModifiedFields = @{} # To store modified fields

        # Mapping database fields to AD fields
        $attributes = @{
            "last_name"     = "sn"
            "first_name"    = "GivenName"
            "display_name"  = "DisplayName"
            "description"   = "Description"
            "email"         = "mail"
            "sam"           = "sAMAccountName"
            "upn"           = "userPrincipalName"
            "address"       = "StreetAddress"
            "city"          = "l"
            "state"         = "st"
            "zipcode"       = "PostalCode"
            "company"       = "Company"
            "job_title"     = "Title"
            "department"    = "Department"
        }

        # Iterate over each attribute to check and update
        foreach ($pair in $attributes.GetEnumerator()) {
            $dbField = $pair.Key
            $adField = $pair.Value

            $dbValue = $DataReader[$dbField]
            $adValue = $ADUser.$adField

            if ($dbField -eq "sam" -and $dbValue -ne $null -and $dbValue -ne $adValue) {
                # Special handling for sAMAccountName change
                $newSAM = $dbValue
                $currentSAM = $ADUser.SamAccountName
                Set-ADUser -Identity $currentSAM -SamAccountName $newSAM

                try {
                    $ADUser = Get-ADUser -Identity $newSAM
                    Write-Host "User with CustomIdentifier '$CustomIdentifier' updated sAMAccountName to '$newSAM'"
                } catch {
                    Write-Host "Error updating sAMAccountName: $_"
                }
            } elseif ($dbValue -ne $null -and $dbValue -ne $adValue) {
                # Handle updates for other fields
                $ModifiedFields[$adField] = $dbValue
                $ADUser.$adField = $dbValue
            }
        }

        # If there are changes, apply them
        if ($ModifiedFields.Count -gt 0) {
            try {
                Set-ADUser -Instance $ADUser -ErrorAction Stop
                Write-Host "User with CustomIdentifier '$CustomIdentifier' updated fields: $($ModifiedFields.Keys -join ', ')"
            } catch {
                Write-Host "Error updating user attributes: $_"
            }
        } else {
            Write-Host "User with CustomIdentifier '$CustomIdentifier' already exists in Active Directory. No fields updated."
        }
    } else {
        # User doesn't exist in AD, so create them

        # Define the parameters for the New-ADUser cmdlet
        $UserParams = @{
            "Path"             = $OuPath
            "Name"             = $DataReader["display_name"]
            "GivenName"        = $DataReader["first_name"]
            "Surname"          = $DataReader["last_name"]
            "DisplayName"      = $DataReader["display_name"]
            "Description"      = $DataReader["description"]
            "EmailAddress"     = $DataReader["email"]
            "SamAccountName"   = $DataReader["sam"]
            "UserPrincipalName"= $DataReader["upn"]
            "StreetAddress"    = $DataReader["address"]
            "City"             = $DataReader["city"]
            "State"            = $DataReader["state"]
            "PostalCode"       = $DataReader["zipcode"]
            "Company"          = $DataReader["company"]
            "Title"            = $DataReader["job_title"]
            "Department"       = $DataReader["department"]
            "AccountPassword"  = (ConvertTo-SecureString "Password123@" -AsPlainText -Force)
            "Enabled"          = $true
            "OtherAttributes"  = @{"msDS-cloudExtensionAttribute1" = $CustomIdentifier}
        }

        try {
            # Attempt to create the user
            New-ADUser @UserParams -ErrorAction Stop
            Write-Host "User with CustomIdentifier '$CustomIdentifier' added to Active Directory in OU $OuPath."Write-Host "User with CustomIdentifier '$CustomIdentifier' added to Active Directory in OU $OuPath."
        } catch {
            Write-Host "Error creating the user: $_"
        }
    }
}

# Close the SQLite connection
$Connection.Close()
