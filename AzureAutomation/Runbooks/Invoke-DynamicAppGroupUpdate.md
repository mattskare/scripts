# Dynamic Application Group Update Script
## Overview
This PowerShell script automates the creation and management of Microsoft Entra (formerly Azure AD) device groups based on installed applications detected by Microsoft Intune. The script finds devices with specified applications installed and synchronizes Entra ID group membership accordingly.

## Features
- Automatically creates device groups named "Intune - DynamicApp - [Application Name]"
- Updates existing group memberships based on current application installation status
- Adds devices with the application installed to the corresponding group
- Removes devices that no longer have the application installed
- Provides detailed logging and summary reports of all actions taken

## Prerequisites
- PowerShell 5.1 or higher
- Microsoft Graph PowerShell modules:
  - Microsoft.Graph.Authentication
  - Microsoft.Graph.Beta.DeviceManagement
  - Microsoft.Graph.Beta.Groups
- Appropriate Microsoft Entra permissions to:
  - Read device information from Intune
  - Create and manage device groups
  - Read and modify group memberships

## Installation
1. Clone the repository or download the script files.
2. Import into Azure Automation or run locally.
3. Install required Microsoft Graph PowerShell modules if not already installed.
4. Run the script with the necessary parameters.

## Usage
Run the script with the `-Applications` parameter to specify which applications to process:

`.\Invoke-DynamicAppGroupUpdate.ps1 -Applications @("Microsoft Office", "Adobe Acrobat")`

### This will:

- Check for devices with Microsoft Office and Adobe Acrobat installed
- Create or update Entra ID groups for each application
- Synchronize group memberships based on current application installation status

## Parameters
- **Applications** (Mandatory): Array of application names to check for installed devices

## Output
The script provides:

- Console output showing progress and results
- A summary of changes made to each group
- Detailed logs for all operations and errors

## Examples
### Basic Example
`.\Invoke-DynamicAppGroupUpdate.ps1 -Applications @("Microsoft Visio", "7-Zip")`

## Notes
- Application names must match those detected by Intune (case-sensitive)
- This script is designed to run both locally and as an Azure Automation runbook
- Requires appropriate Graph API permissions to interact with Intune and Entra ID
## Troubleshooting
- If the script fails to connect, verify your Microsoft Graph permissions
- Check the logs for detailed error messages (located in the .\Logs folder)
- Ensure application names match exactly as they appear in Intune