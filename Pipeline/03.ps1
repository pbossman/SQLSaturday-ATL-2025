# Demo 3: Advanced Select-Object Usage
# Demonstrates different ways to use Select-Object

# Get system event logs
Get-EventLog -LogName System -Newest 10 |
    Select-Object -Property @{
        # Create a calculated property
        Name       = 'TimeGenerated'
        Expression = { $_.TimeGenerated.ToString('yyyy-MM-dd HH:mm:ss') }
    },
    @{
        Name       = 'EventType'
        Expression = { 
            switch ($_.EntryType) {
                'Error' { '❌ Error' }
                'Warning' { '⚠️ Warning' }
                default { $_.EntryType }
            }
        }
    },
    Source,
    Message |
    Format-Table -AutoSize

Write-Host "Select-Object can create calculated properties using hashtables."
Write-Host "This allows for data transformation directly in the pipeline."