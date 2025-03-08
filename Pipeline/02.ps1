# Demo 2: Filtering with Where-Object
# Shows how to filter data in the pipeline

#region Get services and filter for only the running ones
Get-Service | 
    Where-Object { $_.Status -eq 'Running' } |
    # Further filter for services that start with "W"
    Where-Object { $_.Name -like 'W*' } |
    # Select specific properties to display
    Select-Object Name, DisplayName, Status
    
Write-Host "Where-Object acts as a filter in the pipeline,"
Write-host "    allowing only matching items to continue."
Write-Host "You can chain multiple Where-Object cmdlets"
Write-host "    for complex filtering scenarios."
#endregion