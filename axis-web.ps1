# Set the FPS, Bandwidth, and Stream Profile values
$fps = 7
$bandwidth = 500
$stream_profile = "Main"

# Set the path to the Excel file
$excel_file = "IP1.xlsx"

# Ask for credentials
Write-Host "What is the username?"
$username = Read-Host
Write-Host "What is the password?"
$password = Read-Host -AsSecureString

# Set the sheet and range of rows to read from the Excel file
$sheet = "Sheet1"
$startRow = 1
$startColumn = 1
$endRow = 32

# Iterate through the rows in the Excel file
for ($i = $startRow; $i -le $endRow; $i++) {
    # Read the values for the camera IP, username, and password from the Excel file
    $cameraIp = Get-Content -Path $excel_file -Line $i

    # Log in to the camera
    $loginResponse = Invoke-WebRequest -Method POST -Uri "http://$cameraIp/axis-cgi/admin/login.cgi" -Body @{user=$username;pwd=$password}
    # Check if the login was successful
    if ($loginResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully logged in to camera $i"
    } else {
        Write-Host "Failed to log in to camera $i"
        continue
    }

    # Set the FPS
    $setFpsResponse = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.FPS.FPS=$fps"
    # Check if the FPS was set successfully
    if ($setFpsResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully set the FPS for camera $i"
    } else {
        Write-Host "Failed to set the FPS for camera $i"
    }

    # Set the Bandwidth
    $setBandwidthResponse = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.FPS.Bandwidth=$bandwidth"
    # Check if the Bandwidth was set successfully
    if ($setBandwidthResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully set the Bandwidth for camera $i"
    } else {
        Write-Host "Failed to set the Bandwidth for camera $i"
    }

    # Set the Stream Profile
    $setStreamProfileResponse = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.FPS.StreamProfile=$stream_profile"
    # Check if the Stream Profile was set successfully
    if ($setStreamProfileResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully set the Stream Profile for camera $i"
    } else {
        Write-Host "Failed to set the Stream Profile for camera $i"
    }
}
