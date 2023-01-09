# Import the necessary module
Import-Module Microsoft.PowerShell.Utility

# Load the Excel file and get the sheet with the IP addresses
$workbook = Import-Excel -Path 'cameras.xlsx'
$sheet = $workbook | Select-Object -ExpandProperty 'Sheet1'

# Initialize the row counter
$row = 2

# Loop through the rows of the sheet
for ($i = 2; $i -le $sheet.Count; $i++) {
    # Get the IP address
    $ip_address = $sheet[$i].IP_Address

    # Connect to the camera and get the response
    $res = Invoke-WebRequest -Uri "http://$ip_address/axis-cgi/about.cgi"

    # Skip to the next camera if we get a non-200 response
    if ($res.StatusCode -ne 200) {
        continue
    }

    # Parse the response
    $soup = Invoke-WebRequest -Uri "http://$ip_address/axis-cgi/about.cgi" | ConvertFrom-Html

    # Get the IP address, MAC address, make, and model
    $ip_address = $soup.table.tr | Where-Object {$_.th.innerText -eq 'IP address'} | Select-Object -ExpandProperty 'td' | Select-Object -ExpandProperty 'innerText' -First 1
    $mac_address = $soup.table.tr | Where-Object {$_.th.innerText -eq 'MAC address'} | Select-Object -ExpandProperty 'td' | Select-Object -ExpandProperty 'innerText' -First 1
    $make = $soup.table.tr | Where-Object {$_.th.innerText -eq 'Make'} | Select-Object -ExpandProperty 'td' | Select-Object -ExpandProperty 'innerText' -First 1
    $model = $soup.table.tr | Where-Object {$_.th.innerText -eq 'Model'} | Select-Object -ExpandProperty 'td' | Select-Object -ExpandProperty 'innerText' -First 1

    # Write the camera information to the sheet
    $sheet[$row] | Add-Member -MemberType NoteProperty -Name 'IP_Address_(Scraped)' -Value $ip_address
    $sheet[$row] | Add-Member -MemberType NoteProperty -Name 'MAC_Address' -Value $mac_address
    $sheet[$row] | Add-Member -MemberType NoteProperty -Name 'Make' -Value $make
    $sheet[$row] | Add-Member -MemberType NoteProperty -Name 'Model' -Value $model

    # Increment the row counter
    $row++
}

# Save the workbook
Export-Excel -Workbook $workbook -Path 'cameras.xlsx'

Write-Host 'Done!'
