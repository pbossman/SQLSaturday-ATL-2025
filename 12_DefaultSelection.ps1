# data from 3
Copy-Item -Path "$demoFolder\DemoExcel_3.xlsx" -Destination "$demoFolder\DemoExcel_12.xlsx"
$excelPath = "$demoFolder\DemoExcel_12.xlsx"

$excel = Open-ExcelPackage -Path $excelPath
$ws = $excel.Workbook.Worksheets

$sheet = $ws['SalesData']

$sheet.WorksheetXml.worksheet.sheetViews.sheetView.selection.activeCell = 'B2'
$sheet.WorksheetXml.worksheet.sheetViews.sheetView.selection.sqref = 'B2:C7'
$sheet.View.ZoomScale = 145
Close-ExcelPackage $excel

ii $excelPath



$PSDefaultParameterValues.Add("Export-Excel:TableStyle", 'Medium2')
$PSDefaultParameterValues.Add("Export-Excel:AutoSize", $true)
$PSDefaultParameterValues.Add("Export-Excel:FreezeTopRow", $true)
$PSDefaultParameterValues.Add("Get-Service:ErrorAction", 'SilentlyContinue')
