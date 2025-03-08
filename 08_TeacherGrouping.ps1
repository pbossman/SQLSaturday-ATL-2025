# Create multiple advanced pivot tables using New-PivotTableDefinition


# Create a new Excel package for advanced pivot tables
$TeacherList= "$demoFolder\TeacherList.xlsx"

# First, export the raw data
$students = Import-Excel -Path C:\GitHub\ImportExcel\__SQL-Saturday__\_Output\StudentAnalytics.xlsx -WorksheetName 'RawData'

$students | group teacher | ForEach-Object {
    $currTeacher = $_.Name
    $group = $_.Group
    Write-Verbose "Adding Sheet $($currTeacher)" -Verbose
    $group | Export-Excel -Path $TeacherList -WorksheetName $currTeacher -ClearSheet -AutoSize -FreezeTopRow -TableName ($currTeacher -replace '. ')  -TableStyle Medium3
}

& $TeacherList




$students | Group-Object teacher | ForEach-Object {
    $currTeacher = $_.Name
    $group = $_.Group
    Write-Verbose "Adding Sheet $($currTeacher)" -Verbose
    $group | Sort lastname, FirstName | Export-Excel -Path $TeacherList -WorksheetName $currTeacher -ClearSheet -AutoSize -FreezeTopRow -TableName ($currTeacher -replace '. ') -TableStyle Medium2
}

& $TeacherList
