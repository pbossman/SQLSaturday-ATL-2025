# data from 3
Copy-Item -Path "$demoFolder\DemoExcel_3.xlsx" -Destination "$demoFolder\DemoExcel_11.xlsx"
$excelPath = "$demoFolder\DemoExcel_11.xlsx"

$excel = Open-ExcelPackage -Path $excelPath
$ws = $excel.Workbook.Worksheets

$sheet = $ws['SalesData']

$sheet.View.ZoomScale
$sheet.View.ZoomScale = 10
Close-ExcelPackage $excel

ii $excelPath
