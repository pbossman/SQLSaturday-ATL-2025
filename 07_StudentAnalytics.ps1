# Create multiple advanced pivot tables using New-PivotTableDefinition
Write-Host "Creating advanced analytical pivot tables..." -ForegroundColor Yellow

# Create a new Excel package for advanced pivot tables
$advancedPivotPath = "$demoFolder\StudentAnalytics.xlsx"

# First, export the raw data
$students | Export-Excel -Path $advancedPivotPath -WorksheetName "RawData" -AutoSize -TableName "StudentData"

# Define multiple pivot tables
$pivotTableDefinitions = New-PivotTableDefinition -SourceWorksheet "RawData"`
    -PivotTableName "TeacherGradeAnalysis" `
    -PivotRows "Teacher" `
    -PivotColumns "Grade" `
    -PivotData @{"StudentID" = "Count" } `
    -IncludePivotChart `
    -ChartType ColumnClustered `
    -ShowPercent `
    -ChartTitle "Grade Distribution by Teacher" `
    -PivotTotals Both `
    -PivotFilter "Gender" `
    -ChartRow 2 -ChartColumn 5  `
-Activate
    

# Gender balance analysis
$pivotTableDefinitions += New-PivotTableDefinition -SourceWorksheet "RawData" `
    -PivotTableName "GenderBalanceAnalysis" `
    -PivotRows "Teacher" `
    -PivotColumns "Gender" `
    -PivotData @{"StudentID" = "Count" } `
    -IncludePivotChart `
    -ChartType PieExploded3D `
    -ChartTitle "Gender Distribution" `
    -ShowPercent `
    -ChartRow 2 -ChartColumn 5

# Grade distribution analysis
$pivotTableDefinitions += New-PivotTableDefinition -SourceWorksheet "RawData" `
    -PivotTableName "GradeOverview" `
    -PivotRows "Grade" `
    -PivotData @{"StudentID" = "Count" } `
    -IncludePivotChart -ChartType BarStacked `
    -ChartTitle "Overall Grade Distribution" `
    -NoLegend `
    -ChartRow 2 -ChartColumn 3
    
# Export with all pivot table definitions
Export-Excel -Path $advancedPivotPath -PivotTableDefinition $pivotTableDefinitions -AutoSize -Show

Write-Host "Created advanced pivot table analytics in $advancedPivotPath" -ForegroundColor Green