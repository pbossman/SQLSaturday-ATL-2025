#region setup
Copy-Item -Path "$demoFolder\DemoExcel_1.xlsx" -Destination "$demoFolder\DemoExcel_2.xlsx"
$excelPath = "$demoFolder\DemoExcel_2.xlsx"
#endregion

# 2. Add conditional formatting
$excelPackage = Open-ExcelPackage -Path $excelPath
$ws = $excelPackage.Workbook.Worksheets["Processes"]
Add-ConditionalFormatting -Worksheet $ws -Range "C:C" -RuleType GreaterThan -ConditionValue "50" -BackgroundColor LightPink
Add-ConditionalFormatting -Worksheet $ws -Range "D:D" -RuleType GreaterThan -ConditionValue "24743936" -BackgroundColor LightGreen
Close-ExcelPackage $excelPackage

& $excelPath

Write-Host "Added conditional formatting to highlight high CPU and memory usage" -ForegroundColor Green
