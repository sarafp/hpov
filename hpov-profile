# Import required modules
Import-Module ImportExcel
Import-Module HPEOneView.800

# Define the HPE OneView connection details
$oneviewIP = "ONEVIEW_IP_ADDRESS"
$oneviewUsername = "ONEVIEW_USERNAME"
$oneviewPassword = "ONEVIEW_PASSWORD"

# Path to the Excel file
$excelFilePath = "C:\path\to\servers.xlsx"

# Connect to HPE OneView
$ovConnection = Connect-OVMgmt -Hostname $oneviewIP -UserName $oneviewUsername -Password $oneviewPassword

if ($ovConnection) {
    Write-Output "Connected to HPE OneView successfully."
} else {
    Write-Output "Failed to connect to HPE OneView."
    exit 1
}

# Import the Excel file
$servers = Import-Excel -Path $excelFilePath

foreach ($server in $servers) {
    $iLOName = $server.iLOName
    $rackName = $server.RackName
    $profileName = $server.ProfileName
    $templateName = $server.TemplateName

    # Process each server
    Write-Output "Processing host: $iLOName"

    # Retrieve server hardware based on iLO Name
    $serverHardware = Get-OVServer | Where-Object { $_.name -eq $iLOName }

    if ($null -eq $serverHardware) {
        Write-Output "Server hardware with iLO Name $iLOName not found."
        continue
    }

    Write-Output "Found server hardware: $($serverHardware.name), URI: $($serverHardware.uri)"

    # Check if the rack exists
    $existingRack = Get-OVRack | Where-Object { $_.name -eq $rackName }

    if ($existingRack) {
        Write-Output "Rack '$rackName' already exists."
    } else {
        Write-Output "Creating rack '$rackName'."

        # Define the rack configuration
        $rackConfig = @{
            name = $rackName
            # Add other necessary configurations as per your requirements
        }

        # Create the rack
        try {
            $newRack = New-OVRack -InputObject $rackConfig
            Write-Output "Rack created successfully: $($newRack.uri)"
        } catch {
            Write-Output "Failed to create rack: $_"
            continue
        }
    }

    # Assign the server to the rack
    try {
        $serverHardware.rackName = $rackName
        Set-OVServer -InputObject $serverHardware
        Write-Output "Server '$iLOName' assigned to rack '$rackName' successfully."
    } catch {
        Write-Output "Failed to assign server '$iLOName' to rack '$rackName': $_"
        continue
    }

    # Check if the profile already exists
    $existingProfile = Get-OVServerProfile | Where-Object { $_.name -eq $profileName }

    if ($existingProfile) {
        Write-Output "Server profile '$profileName' already exists."
    } else {
        Write-Output "Creating server profile '$profileName' using template '$templateName'."

        # Retrieve the server profile template
        $template = Get-OVServerProfileTemplate | Where-Object { $_.name -eq $templateName }

        if ($null -eq $template) {
            Write-Output "Template '$templateName' not found."
            continue
        }

        # Define the server profile configuration based on the template
        $serverProfileConfig = @{
            type = "ServerProfileV11"
            name = $profileName
            serverHardwareUri = $serverHardware.uri
            serverProfileTemplateUri = $template.uri
            hideUnusedFlexNics = $true
            # Add other necessary configurations as per your requirements
        }

        # Create the server profile
        try {
            $newServerProfile = New-OVServerProfile -InputObject $serverProfileConfig
            Write-Output "Server profile created successfully: $($newServerProfile.uri)"
        } catch {
            Write-Output "Failed to create server profile: $_"
            continue
        }

        # Apply the server profile
        try {
            $assignedProfile = Set-OVServerProfile -InputObject $newServerProfile
            Write-Output "Server profile applied successfully: $($assignedProfile.uri)"
        } catch {
            Write-Output "Failed to apply server profile: $_"
            continue
        }
    }
}

# Disconnect from HPE OneView
Disconnect-OVMgmt
Write-Output "Disconnected from HPE OneView."
