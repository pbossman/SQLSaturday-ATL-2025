$excelPath = "$demoFolder\DemoExcel_1.xlsx"

#region 1. Create a simple Excel file with process data
Get-Process | 
    Where-Object { $_.Company } | 
    Select-Object -Property Name, ID, CPU, WorkingSet, Description, Company |
    Export-Excel -Path $excelPath 

Write-Host "Created Excel file with process data at $excelPath" -ForegroundColor Green
Invoke-Item $excelPath
#endregion







#region Better looking Excel file with process data
Get-Process | 
    Where-Object { $_.Company } | 
    Select-Object -Property Name, ID, CPU, WorkingSet, Description, Company |
    Tee-Object -Variable DemoExcel_1 |
    Export-Excel -Path $excelPath -WorksheetName "Processes" -AutoSize -TableName "ProcessTable"

Write-Host "Better looking Excel file with process data at $excelPath" -ForegroundColor Green
Invoke-Item $excelPath

#endregion