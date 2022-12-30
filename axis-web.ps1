# Set the values for the cameras
$fps = 7
$zipstream = "Medium"
$dynamicFPS = "on"
$lowerLimit = 1
$pFrames = 60
$bitrateControl = "Maximum"
$bitrate = 500
$priority = "No Priority"
$h264Profile = "Main"
$dns1 = "169.88.8.8"
$dns2 = "169.88.9.9"
$factUser = "root"
$factPass = "pass"

# Set the path to the Excel file
$excel_file = "IP1.xlsx"

# Ask for credentials
Write-Host "What is the username?"
$username = Read-Host
Write-Host "What is the new password?"
$password = Read-Host
Write-Host "What is the NTP Server?"
$ntpServer = Read-Host
Write-Host "What is the time zone? ex. PST, CST etc.."
$timeZone = Read-Host

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
    $loginResponse = Invoke-WebRequest -Method POST -Uri "http://$cameraIp/axis-cgi/admin/login.cgi" -Body @{user=$factUser;pwd=$factPass}
    # Check if the login was successful
    if ($loginResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully logged in to camera $i"
    } else {
        Write-Host "Failed to log in to camera $i"
        continue
    }

    # Set the parameters for the 2nd stream H.264 stream profile
    $secresolution = "640x360"
    $secframeRate = 2
    $seccompression = 30
    $secgopLength = 60
    $sech264Profile = "Main"
    $secmaxBitrate = 500
    $secpriority = "No Priority"

    # Create the H.264 stream profile
    $createProfileResponse = Invoke-WebRequest -Method POST -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update" -Body @{
        "Video.StreamProfile.H264.Name" = "Bandwidth";
        "Video.StreamProfile.H264.Resolution" = $secresolution;
        "Video.StreamProfile.H264.FrameRate" = $secframeRate;
        "Video.StreamProfile.H264.Compression" = $seccompression;
        "Video.StreamProfile.H264.GopLength" = $secgopLength;
        "Video.StreamProfile.H264.Profile" = $sech264Profile;
        "Video.StreamProfile.H264.MaxBitrate" = $secmaxBitrate;
        "Video.StreamProfile.H264.Priority" = $secpriority


    # Check if the H.264 stream profile was created successfully
    if ($createProfileResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully created H.264 stream profile"
    } else {
        Write-Host "Failed to create H.264 stream profile"
    }


    # Set the FPS
    $setFpsResponse = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.FPS.FPS=$fps"
    # Check if the FPS was set successfully
    if ($setFpsResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully set the FPS for camera $i"
    } else {
        Write-Host "Failed to set the FPS for camera $i"
    }

    # Set the Bitrate
    $setBandwidthResponse = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.FPS.Bandwidth=$bitrate"
    # Check if the Bandwidth was set successfully
    if ($setBandwidthResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully set the Bandwidth for camera $i"
    } else {
        Write-Host "Failed to set the Bandwidth for camera $i"
    }

    # Set the Stream Profile
    $setStreamProfileResponse = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.FPS.StreamProfile=$h264Profile"
    # Check if the Stream Profile was set successfully
    if ($setStreamProfileResponse.StatusDescription -match "200 OK") {
        Write-Host "Successfully set the Stream Profile for camera $i"
    } else {
        Write-Host "Failed to set the Stream Profile for camera $i"
    }
	
	# Set the zipstream
	$setZipStreamResponse = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.Zipstream=$zipstream"
	# Check if zipstream was set successfully
	if ($setZipStreamResponse.StatusDescription -match "200 OK") {
		Write-Host "Successfully set the Zipstream for camera $i"
	} else {
		Write-Host "Failed to set the Zipstream for camera $i"
	}
	
	# Enable Dynamic FPS
	$setDynamicFps = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.DynamicFPS.Enabled=$dynamicFPS"
	# Check if Dynamic FPS was set to enabled
	if ($setDynamicFps.StatusDescription -match "200 OK") {
		Write-Host "Successfully enabled Dynamic FPS for camera $i"
	} else {
		Write-Host "Failed to enable Dynamic FPS for camera $i"
	}
	
	# Set the Dynamic Lower Limit FPS
	$setDynamicLower = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.DynamicFPS.LowerLimit=$lowerLimit"
	# Check if LowerLimit was set
	if ($setDynamicLower.StatusDescription -match "200 OK") {
		Write-Host "Successfully set the Dynamic Lower Limit for camera $1"
	} else {
		Write-Host "Failed to set Dynamic Lower Limit for camera $i"
	}
	
	# Set the P-Frames
	$setPframes = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.PFrames=$pFrames"
	# Check if the P-Frames were set
	if ($setPframes.StatusDescription -match "200 OK") {
		Write-Host "Successfully set the PFrames for camera $i"
	} else {
		Write-Host "Failed to set the PFrames for camera $i"
	}
	
	# Set the Bitrate Control
	$setBitrateControl = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.BitrateConrol=$bitrateControl"
	# Check if the bitrate control was set
	if ($setBitrateControl.StatusDescription -match "200 OK") {
		Write-Host "Successfully set Bitrate Control for camera $i"
	} else {
		Write-Host "Failed to set Bitrate Control for camera $i"
	}
	
	# Set the priority
	$setPriority = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.Priority=$priority"
	# Check if the priority was set
	if ($setPriority.StatusDescription -match "200 OK" {
		Write-Host "Successfully set Priority on camera $i"
	} else {
		Write-Host "Failed to set Priority on camera $i"
	}
	
	# Set the H.264 Profile
	$setH264Profile = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&Video.H264Profile=$h264Profile"
	# Check if the profile was set
	if ($setH264Profile.StatusDescription -match "200 OK" {
		Write-Host "Successfully set the H.264 Profile on camera $i"
	} else {
		Write-Host "Failed to set the H.264 Profile on camera $i"
	}
	
	# Enable Motion Detection
	$setMotionDetection = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&MothingDetectionApp=enabled"
	# Check if Motion Detection was enabled
	if ($setMotionDetection.StatusDescription -match "200 OK" {
		Write-Host "Successfully Enabled Motion Detection for camera $i"
	} else {
		Write-Host "Failed to enable Motion Detection for camera $i"
	}
	
	# Set DNS1 and 2
	$setDnsNetwork = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/param.cgi?action=update&PrimaryDNS=$dns1&SecondaryDNS=$dns2"
	# Check if the DNS was set correctly
	if ($setDnsNetwork.StatusDescription -match "200 OK" {
		Write-Host "Successfully set DNS1 and DNS2 for camera $i"
	} else {
		Write-Host "Failed to set DNS1 and DNS2 for camera $i"
	}
	
	$setUserPass = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/pwdgrp.cgi?action=update&user=$username&pwd1=$password&pwd2=$password"
	# Check if User/Pass was set correctly
	if ($setUserPass.StatusDescription -match "200 OK" {
		Write-Host "Successfully set username and password for camera $i"
	} else {
		Write-Host "Failed to set username and password for camera $i"
	}
	
	# Set NTP Server and TimeZone
	$setNtpTimeZone = Invoke-WebRequest -Uri "http://$cameraIp/axis-cgi/admin/pwdgrp.cgi?action=update&ntpServer=$ntpServer&TZ=$timeZone"
	# Check if NTP Server and TimeZone were set
	if ($setNtpTimeZone.StatusDescription -match "200 OK" {
		Write-Host "Successfully set NTP Server and TimeZone for camera $i"
	} else {
		Write-Host "Failed to set NTP Server and TimeZone for camera $i"
	}
}
