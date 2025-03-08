#region setup
Copy-Item -Path "$demoFolder\DemoExcel_3.xlsx" -Destination "$demoFolder\DemoExcel_4.xlsx"
$excelPath = "$demoFolder\DemoExcel_4.xlsx"
#endregion

# 4. Import data from Excel
Write-Host "Importing data from the Excel file:" -ForegroundColor Yellow
$importedData = Import-Excel -Path $excelPath -WorksheetName "SalesData"
$importedData | Format-Table
