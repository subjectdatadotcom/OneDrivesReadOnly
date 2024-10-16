<#
.SYNOPSIS
This script processes user information from a CSV file, retrieves OneDrive sites for each user, sets them to read-only, and checks their lock state.

.DESCRIPTION
The script ensures that the Microsoft.Online.Sharepoint.PowerShell module is installed and imported. It reads a list of OneDrive users from a CSV file, connects to a SharePoint Online tenant, retrieves the corresponding OneDrive site for each user, sets the site to read-only, and checks the lock state for each OneDrive site, providing feedback on the operation.

.NOTES
The script requires an account with SharePoint Online Administrator permissions for authentication and access to OneDrive sites.

.AUTHOR
SubjectData

.EXAMPLE
.\OneDriveLockScript.ps1
This will run the script in the current directory, processing the 'OneDriveUsers.csv' file and setting the corresponding OneDrive sites to read-only while checking their lock states.
#>

# Define the module name for SharePoint Online
$moduleName = "Microsoft.Online.Sharepoint.PowerShell"

# Check if the SharePoint Online module is already installed
if (-not(Get-Module -Name $moduleName)) {
    # Install the SharePoint Online module if it's not installed
    Install-Module -Name $moduleName -Force
}

# Import the SharePoint Online module
Import-Module Microsoft.Online.Sharepoint.PowerShell -Force

# Connect to the SharePoint Online service using the provided admin URL
Connect-SPOService -url "https://m365x84490777-admin.sharepoint.com"

# Get the directory of the current script
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define the location of the CSV file containing OneDrive user emails
$XLloc = "$myDir\"

try {
    # Import the list of OneDrive users from the CSV file
    $OneDriveUsers = import-csv ($XLloc + "OneDriveUsers.csv").ToString()
} catch {
    # Handle the error if the CSV file is not found
    Write-Host "No CSV file to read" -BackgroundColor Black -ForegroundColor Red
    exit
}

# Loop through each user in the imported CSV
foreach ($User in $OneDriveUsers) {
    try {
        # Check if the email field is not empty
        if ($User.Email.ToString() -ne "") {
            # Retrieve the OneDrive site URL for the user
            $OneDriveLink = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | Where-Object { $_.Owner -ieq $User.Email } | Select-Object -ExpandProperty Url

            # Set the OneDrive site to read-only
            Set-SPOSite -Identity $OneDriveLink.ToString() -LockState ReadOnly

            # Check the lock state of the OneDrive site
            $Result = Get-SPOSite $OneDriveLink.ToString() | Select-Object lockstate
            if ($Result.LockState.ToString() -eq "Unlock") {
                Write-Host $OneDriveLink.ToString() " OneDrive is" $Result.LockState
            } else {
                Write-Host $OneDriveLink.ToString() " OneDrive is" $Result.LockState -BackgroundColor Black -ForegroundColor Red
            }
        }
    } catch {
        # Handle any exceptions that occur during processing
        Write-Host "Exception occurred for" $OneDriveLink.ToString() -BackgroundColor Black -ForegroundColor Red
        Continue
    }
}
