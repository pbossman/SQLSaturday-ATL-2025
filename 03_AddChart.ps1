$excelPath = "$demoFolder\DemoExcel_3.xlsx"

# 3. Add a chart
$chartData = @"
"Month","Sales","Expenses"
"January","25000","15000"
"February","28500","16500"
"March","30200","18000"
"April","32800","19500"
"May","35000","21000"
"June","38500","22500"
"July","40000","24000"
"August","42500","25500"
"@ | ConvertFrom-Csv

$chartData | Export-Excel -Path $excelPath -WorksheetName "SalesData" 
$excel = Open-ExcelPackage -Path $excelPath
Add-ExcelChart -Worksheet $excel.Workbook.Worksheets["SalesData"] `
    -ChartType LineMarkersStacked `
    -XRange "Month" `
    -YRange "Sales" `
    -Title "Monthly Sales vs Expenses" `
    -LegendPosition Bottom
Close-ExcelPackage $excel -Show
Write-Host "Added sales data with a line chart comparing sales and expenses" -ForegroundColor Green
