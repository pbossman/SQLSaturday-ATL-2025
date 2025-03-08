# data from 3
Copy-Item -Path "$demoFolder\DemoExcel_3.xlsx" -Destination "$demoFolder\DemoExcel_9.xlsx"
$excelPath = "$demoFolder\DemoExcel_9.xlsx"

# 9. Password protect the workbook
$excel = Open-ExcelPackage -Path $excelPath 

# Lock the structure and windows
$excel.Workbook.Protection.LockStructure = $true
$excel.Workbook.Protection.LockWindows = $true

# Enable protection on all worksheets to prevent modification
foreach ($worksheet in $excel.Workbook.Worksheets) {
    $worksheet.Protection.IsProtected = $true
    $worksheet.Protection.SetPassword("Demo123")
    $worksheet.Protection.AllowInsertRows = $true
    $worksheet.Protection.AllowSort = $true
    $worksheet.Protection.AllowSelectLockedCells = $true  # Allow navigation but not editing
}

# Save the changes
Close-ExcelPackage $excel -Show 

Write-Host "Password protected the workbook with password 'Demo123'" -ForegroundColor Green

& $excelPath
