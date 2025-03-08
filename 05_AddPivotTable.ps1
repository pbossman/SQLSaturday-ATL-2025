## Data from 1
Copy-Item -Path "$demoFolder\DemoExcel_1.xlsx" -Destination "$demoFolder\DemoExcel_5.xlsx"
$excelPath = "$demoFolder\DemoExcel_5.xlsx"

# 5. Use PivotTable
$processes = Import-Excel -Path $excelPath -WorksheetName "Processes"
$processes | Export-Excel -Path $excelPath -WorksheetName "PivotDemo" -PivotRows "Company" -PivotData @{"Name" = "Count" } -PivotTableName "ProcessesByCompany" -Show -Activate

Write-Host "Added a PivotTable showing process count by company" -ForegroundColor Green

# Clean up (optional)
# Remove-Item $excelPath -Force
Write-Host "Demo Excel file can be found at: $excelPath" -ForegroundColor Cyan

