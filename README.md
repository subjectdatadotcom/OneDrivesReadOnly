# OneDrive Lock Script

This PowerShell script processes users from a CSV file, retrieves their OneDrive sites, sets them to read-only, and checks their lock state.

## Prerequisites

1. **PowerShell**: Ensure you have PowerShell installed.
2. **SharePoint Online Management Shell**: This script uses the `Microsoft.Online.Sharepoint.PowerShell` module. The script will automatically install it if it's not already installed.

## Instructions

1. **Edit the Script**:
   - Open the script file.
   - Locate the following line:
     ```powershell
     Connect-SPOService -url "https://m365x84490777-admin.sharepoint.com"
     ```
   - Replace `"https://m365x84490777-admin.sharepoint.com"` with your own SharePoint tenant admin URL. It should look something like this:
     ```powershell
     Connect-SPOService -url "https://your-tenant-admin.sharepoint.com"
     ```

2. **Prepare the CSV File**:
   - Ensure you have a `OneDriveUsers.csv` file in the same directory as the script.
   - The CSV file should have the following structure:
     ```
     Email
     user1@example.com
     user2@example.com
     ```

3. **Run the Script**:
   - Open PowerShell as an administrator.
   - Navigate to the directory containing the script.
   - Run the script:
     ```powershell
     .\OneDriveLockScript.ps1
     ```
   - Authenticate using an admin account.
   - The script will connect to your SharePoint tenant, process the users, and set their OneDrive sites to read-only.

4. **Check the Output**:
   - The script provides feedback in the console about the lock state of each OneDrive site.

## Troubleshooting

- **No CSV file to read**: Ensure the `OneDriveUsers.csv` file is present and correctly formatted.
- **Permission Issues**: Ensure you have the necessary permissions to connect to the SharePoint tenant and access the OneDrive sites.
- **Module Installation**: If the script fails to install the module, try manually installing it:
  ```powershell
  Install-Module -Name Microsoft.Online.Sharepoint.PowerShell -Force
Additional Notes
This script is designed to be run in an environment with access to the SharePoint Online Management Shell and requires SharePoint Online Administrator permissions.
