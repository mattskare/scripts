<#
.SYNOPSIS
    Updates Microsoft Entra dynamic groups based on installed applications on devices.
    This script checks for devices with specified applications installed and updates 
    the corresponding dynamic groups accordingly.
.DESCRIPTION
    This script retrieves devices with specified applications installed from Microsoft Intune,
    checks if corresponding dynamic groups exist in Microsoft Entra, and updates group memberships
    based on the installed applications. If a group does not exist, it creates a new dynamic group.
.PARAMETER Applications
    An array of application names to check for installed devices. This parameter is mandatory.
    The application names should match those in discovered applications.
.EXAMPLE
    Invoke-DynamicAppGroupUpdate -Applications @("Docker Desktop", "7-Zip")
    This example updates dynamic groups for devices with Docker Desktop and 7-Zip installed.
.INPUTS
    Accepts an array of application names from the pipeline.
.OUTPUTS
    Outputs the status of group updates, including devices added or removed from groups.
    Logs are written to a specified log folder.
.NOTES
    Requires the following Microsoft Graph PowerShell modules:
    Microsoft.Graph.Authentication, Microsoft.Graph.Beta.DeviceManagement, and Microsoft.Graph.Beta.Groups.
.NOTES
    Created by: Matt Skare
    Date: 05-27-2025
    Version: 1.0

    Version History:
    - 1.0: Initial version - Created script to manage dynamic groups based on installed applications.
#>

#region Script variables
########################
# Modify this section to set the tenant ID and client ID
########################
$yourTenantId = "your-tenant-id"  # Replace with your tenant ID
$yourApplicationId = "your-application-id"  # Replace with your application ID
$yourCertificateName = "your-certificate-name"  # Replace with your Azure Automation certificate name
#endregion

# Parameters for the script
param(
    [Parameter(Mandatory=$false, HelpMessage="Array of applications")]
    [string[]]$Applications
)

function Write-Log {
    <#
    .SYNOPSIS
        Writes log message to file and console.
    .PARAMETER Message
        Message to log
    .PARAMETER Level
        Severity level (Info, Warning, Error)
    .PARAMETER Action
        Action being performed
    .NOTES
        TODO: Update documentation with more details
        More information about the function.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Debug', 'Verbose')]
        [string]$Level = 'Info',
        [string]$Action = 'General'
    )

    begin {
        # Get the name of the script that called this function
        $callingScript = (Get-PSCallStack)[1].ScriptName
        if (-not $callingScript) {
            $callingScript = "PowerShell_Console"
        }

        # Create log filename based on calling script
        $logFileName = [System.IO.Path]::GetFileNameWithoutExtension($callingScript)

        # Test if $LogFolder is provided, if not use the default log folder
        if (-not $LogFolder) {
            # Default log folder is wherever the script is running from
            Write-Warning "Log folder not specified. Please specify a log folder. Using default log folder."
            $LogFolder = Join-Path -Path (Get-Location) -ChildPath "Logs"
        }

        # Create log folder if it doesn't exist then append the log file name with the date
        if (!(Test-Path $LogFolder)) {
            $null = New-Item -ItemType Directory -Path $LogFolder
        }
        $LogFileName = "{0}_{1}.log" -f $logFileName, (Get-Date -Format 'yyyyMMdd')
        $logFile = Join-Path -Path $LogFolder -ChildPath $LogFileName
    }

    process {
        $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] [$Action] $Message"
        Add-Content -Path $logFile -Value $logEntry

        switch ($Level) {
            'Info' { Write-Host $logEntry }
            'Warning' { Write-Warning $logEntry }
            'Error' { Write-Error $logEntry }
            'Debug' { Write-Debug $logEntry }
            'Verbose' { Write-Verbose $logEntry }
        }
    }
}

function Invoke-DynamicAppGroupUpdate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Application array must be application names from discovered applications", ValueFromPipeline = $true)]
        [array]$Applications,

        [Parameter(Mandatory = $false)]
        [string]$TenantId = $yourTenantId,

        [Parameter(Mandatory = $false)]
        [string]$ClientId = $yourApplicationId,

        [Parameter(Mandatory = $false)]
        [string]$CertificateName = $yourCertificateName,

        [Parameter(Mandatory = $false)]
        [string]$CertThumbprint
    )

    begin {
        # Install and import required modules
        $modules = @(
            "Microsoft.Graph.Authentication",
            "Microsoft.Graph.Beta.DeviceManagement",
            "Microsoft.Graph.Beta.Groups"
        )
        # Try to install the required modules if they are not already installed
        foreach ($module in $modules) {
            if (!(Get-Module -ListAvailable -Name $module)) {
                Write-Verbose "Installing $module module..."
                try {
                    Install-Module -Name $module -Force -AllowClobber -ErrorAction Stop
                }
                catch {
                    Write-Error "Failed to install module $module : $_"
                    return $false
                }
            }
        }
        # Import the modules with strict error handling
        foreach ($module in $modules) {
            try {
                Import-Module $module -ErrorAction Stop
                Write-Verbose "Successfully imported $module"
            }
            catch {
                Write-Error "Failed to import $module : $_"
                return $false
            }
        }

        # Ensure we're connected to the Microsoft Graph
        if (-not (Get-MgContext)) {
            try {
                if ($env:AUTOMATION_ASSET_ACCOUNTID) {
                    try {
                        $cert = Get-AutomationCertificate -Name $CertificateName
                        if (-not $cert) {
                            throw "Certificate not found"
                        }
                        Write-Verbose "Retrieved certificate: $($cert.Name) with thumbprint: $($cert.Thumbprint)"
                    }
                    catch {
                        Write-Error "Failed to retrieve Azure Automation certificate: $_"
                        throw
                    }

                    try {
                        Connect-MgGraph -ClientId $ApplicationId -TenantId $TenantId -Certificate $cert
                        Write-Verbose "Connected to Microsoft Graph using Azure Automation certificate."
                    }
                    catch {
                        Write-Error "Failed to connect to Microsoft Graph using AzAutomation certificate: $_"
                        throw
                    }
                } else {
                    Connect-MgGraph -ClientId $ClientId `
                                      -TenantId $TenantId `
                                      -CertificateThumbprint $CertThumbprint
                }
                Write-Verbose "Connected to Microsoft Graph successfully."
            }
            catch {
                Write-Error "Failed to connect to Microsoft Graph: $_"
                return $false
            }
        }

        # Set up the log folder
        $LogFolder = ".\Logs"
        Write-Verbose "Log folder set to $LogFolder"
    }

    process {
        foreach ($app in $Applications) {
            try {
                $intuneDetectedApp = Get-MgBetaDeviceManagementDetectedApp -Sort "displayName " -Filter "(contains(displayName, '$app'))"
            }
            catch {
                Write-Log -Message "Failed to retrieve Intune detected applications for '$app': $_" -Level Error -Action "GetIntuneDetectedAppsFailed"
                exit 1
            }

            Write-Log -Message "Processing application: '$app'" -Level Info -Action "ProcessingApplication"

            # Check if the dynamic group already exists
            $groupName = "Intune - DynamicApp - $app"
            $entraGroup = Get-MgBetaGroup -Filter "displayName eq '$groupName'" -ErrorAction SilentlyContinue
            if ($entraGroup) {
                Write-Log -Message "Dynamic group '$groupName' already exists." -Level Info -Action "GroupExists"
                $dynamicGroup = $entraGroup
            } else {
                Write-Log -Message "Dynamic group '$groupName' does not exist. Creating..." -Level Info -Action "CreateGroup"

                # Generate a description for the group
                $Description = "Dynamic group containing devices with $app installed"
                # Create a static device group
                $MailNickname = ($groupName -replace '[^a-zA-Z0-9]', '') + (Get-Random -Maximum 99999)
                $params = @{
                    DisplayName     = $groupName
                    Description     = $Description
                    MailEnabled     = $false
                    MailNickname    = $MailNickname
                    SecurityEnabled = $true
                }
                Write-Host "Creating static device group '$groupName'..."

                try {
                    $dynamicGroup = New-MgBetaGroup @params -ErrorAction Stop
                    Write-Log -Message "Dynamic group '$groupName' created successfully with ID: $($dynamicGroup.Id)" -Level Info -Action "GroupCreated"
                }
                catch {
                    Write-Log -Message "Failed to create dynamic group '$groupName': $_" -Level Error -Action "GroupCreationFailed"
                }
            }

            $devicesAdded = @()
            $devicesRemoved = @()
            $currentApp = 0
            $devicesWithApplicationInstalled = @()
            foreach ($appVersion in $intuneDetectedApp) {
                $currentApp++
                Write-Output "Detected application: $($appVersion.DisplayName) with version $($appVersion.Version)"

                Write-Progress -Activity "Processing application: '$app'" `
                    -Status "Processing application $currentApp of $($intuneDetectedApp.Count): $($appVersion.DisplayName) version $($appVersion.Version)" `
                    -PercentComplete (($currentApp / $intuneDetectedApp.Count) * 100) `
                    -Id 1

                # Set parameters for graph request
                $appId = $appVersion.Id
                $uri = "https://graph.microsoft.com/beta/deviceManagement/detectedApps('$appId')/managedDevices?`$top=20&`$orderby=deviceName%20asc"

                do {
                    $response = Invoke-MgGraphRequest -Method GET -Uri $uri

                    # Add Entra Object ID to each device
                    $currentDevice = 0
                    foreach ($device in $response.value) {
                        $currentDevice++
                        Write-Progress -Activity "Processing app version: '$($appVersion.Version)'" `
                            -Status "Processing device $currentDevice of $($response.value.Count): $($device.DeviceName)" `
                            -PercentComplete (($currentDevice / $response.value.Count) * 100) `
                            -Id 2 `
                            -ParentId 1

                        $intuneDevice = Get-MgBetaDeviceManagementManagedDevice -ManagedDeviceId $device.Id -ErrorAction SilentlyContinue
                        $entraObjectId = ((Get-MgBetaDevice -Filter "deviceId eq '$($intuneDevice.AzureAdDeviceId)'") | Select-Object -ExpandProperty Id)
                        $device | Add-Member -MemberType NoteProperty -Name 'EntraObjectId' -Value $entraObjectId
                    }
                    Write-Progress -Activity "Processing app version: '$($appVersion.DisplayName)'" -Completed

                    # Add devices with the application installed to the collection
                    $devicesWithApplicationInstalled += $response.value

                    # Get link to next page if it exists
                    $uri = $response.'@odata.nextLink'
                } while ($uri)
            }
            Write-Progress -Activity "Processing application: '$app'" -Completed
            Write-Log -Message "Total devices with application '$($intuneDetectedApp[0].DisplayName)' installed: $($devicesWithApplicationInstalled.Count)" -Level Info -Action "ApplicationDevicesCount"

            # Add devices to dynamic group/update group membership
            $currentDevice = 0
            foreach ($entraDevice in $devicesWithApplicationInstalled) {
                $currentDevice++
                Write-Progress -Activity "Processing devices for app: '$app'" `
                    -Status "Processing device $currentDevice of $($devicesWithApplicationInstalled.Count): $($entraDevice.DeviceName)" `
                    -PercentComplete (($currentDevice / $devicesWithApplicationInstalled.Count) * 100)

                # Check if the device is already a member of the group
                try {
                    $isMember = Get-MgBetaGroupMember -GroupId $dynamicGroup.Id -Filter "id eq '$($entraDevice.EntraObjectId)'" -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Log -Message "Failed to check if device '$($entraDevice.DeviceName)' is a member of group '$($dynamicGroup.DisplayName)': $_" -Level Error -Action "CheckDeviceMembershipFailed"
                    continue
                }
                if ($isMember) {
                    Write-Progress -Activity "Processing devices for app: '$app'" `
                        -Status "Device $currentDevice of $($devicesWithApplicationInstalled.Count) already exists: $($entraDevice.DeviceName)" `
                        -PercentComplete (($currentDevice / $devicesWithApplicationInstalled.Count) * 100)
                    Write-Log -Message "Device '$($entraDevice.DeviceName)' is already a member of the group '$($dynamicGroup.DisplayName)'. Skipping..." -Level Verbose -Action "DeviceAlreadyInGroup"
                    continue
                }

                Write-Progress -Activity "Processing devices for app: '$app'" `
                    -Status "Adding device $currentDevice of $($devicesWithApplicationInstalled.Count): $($entraDevice.DeviceName)" `
                    -PercentComplete (($currentDevice / $devicesWithApplicationInstalled.Count) * 100)
                # Write-Host "Adding device '$($entraDevice.DeviceName)' to group..." -ForegroundColor Yellow
                try {
                    New-MgBetaGroupMember -GroupId $dynamicGroup.Id -DirectoryObjectId $entraDevice.EntraObjectId -ErrorAction Stop
                    $devicesAdded += $entraDevice.DeviceName
                    Write-Log -Message "Added device '$($entraDevice.DeviceName)' to group '$($dynamicGroup.DisplayName)'" -Level Verbose -Action "DeviceAddedToGroup"
                }
                catch {
                    Write-Warning "Could not add device '$($entraDevice.DeviceName)': $_"
                }
            }
            Write-Progress -Activity "Processing devices for app: '$app'" -Completed

            # Get all member devices of the target group and remove those without the application installed
            $existingGroupMemberDevices = Get-MgBetaGroupMember -GroupId $dynamicGroup.Id
            foreach ($existingDevice in $existingGroupMemberDevices) {
                # If the device is not in the source group, remove it from the target group
                if ($devicesWithApplicationInstalled.EntraObjectId -notcontains $existingDevice.Id) {
                    Write-Host "Removing device '$($existingDevice.AdditionalProperties.displayName)' from group..." -ForegroundColor Yellow
                    try {
                        Remove-MgBetaGroupMemberByRef -GroupId $dynamicGroup.Id -DirectoryObjectId $existingDevice.Id
                        $devicesRemoved += $existingDevice.AdditionalProperties.displayName
                        Write-Log -Message "Removed device '$($existingDevice.AdditionalProperties.displayName)' from group '$($dynamicGroup.DisplayName)'" -Level Verbose -Action "DeviceRemovedFromGroup"
                    }
                    catch {
                        Write-Warning "Could not remove device '$($existingDevice.AdditionalProperties.displayName)': $_"
                    }
                }
            }

            Write-Log -Message "Group synchronization for application '$app' completed." -Level Info -Action "GroupSyncCompleted"

            # Display summary of changes
            Write-Host "`nGroup Synchronization Summary:" -ForegroundColor Cyan
            if ($devicesAdded.Count -gt 0) {
                Write-Host "Added $($devicesAdded.Count) device(s):" -ForegroundColor Green
                $devicesAdded | ForEach-Object { Write-Host "  - $_" }
            }
            if ($devicesRemoved.Count -gt 0) {
                Write-Host "Removed $($devicesRemoved.Count) device(s):" -ForegroundColor Yellow
                $devicesRemoved | ForEach-Object { Write-Host "  - $_" }
            }
            if ($devicesAdded.Count -eq 0 -and $devicesRemoved.Count -eq 0) {
                Write-Host "No changes were necessary. The group is already in sync." -ForegroundColor Green
            }
        }
    }

    end {
        # Cleanup or final actions if needed
        Write-Log -Message "Dynamic group update completed." -Level Info -Action "UpdateCompleted"
    }
}

#region Main Execution
try {
    Invoke-DynamicAppGroupUpdate -Applications $Applications
}
catch {
    Write-Error "An error occurred: $_"
}
