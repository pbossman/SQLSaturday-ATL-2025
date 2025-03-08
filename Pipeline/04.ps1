# Demo 4: Export-CSV in the Pipeline
# Shows how to export pipeline results to a CSV file

# Define output path
$outputPath = "C:\GitHub\ImportExcel\__SQL-Saturday__\DiskReport.csv"

# Get disk information
Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |
    # Select and calculate properties
    Select-Object DeviceID, 
    @{Name = 'SizeGB'; Expression = { [math]::Round($_.Size / 1GB, 2) } },
    @{Name = 'FreeSpaceGB'; Expression = { [math]::Round($_.FreeSpace / 1GB, 2) } },
    @{Name = 'PercentFree'; Expression = { [math]::Round(($_.FreeSpace / $_.Size) * 100, 2) } } |
    # Export to CSV
    Export-Csv -Path $outputPath -NoTypeInformation

# Display results and file location
Gent $outputPath | Select-Object -First 50
Write-Host "Full report exported to: $outputPath"
Write-Host "Export-CSV is a pipeline endpoint that saves results to a CSV file."