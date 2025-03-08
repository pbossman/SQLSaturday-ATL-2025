$excelPath = "$demoFolder\DemoExcel_6.xlsx"

# 6. Add multiple worksheets with different data
Write-Host "Adding multiple worksheets with service and event log data..." -ForegroundColor Yellow
Get-Service | Select-Object -First 10 Name, Status, DisplayName | 
    Export-Excel -Path $excelPath -WorksheetName "Services" -AutoSize -TableName "ServicesTable" -TableStyle Medium2

Get-EventLog -LogName System -Newest 10 | 
    Select-Object TimeGenerated, EntryType, Source, Message |
    Export-Excel -Path $excelPath -WorksheetName "SystemLogs" -AutoSize -TableName "SystemLogsTable" -TableStyle Medium3

& $excelPath