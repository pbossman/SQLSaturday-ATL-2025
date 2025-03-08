# Demo 1: Basic Pipeline Introduction
# Shows how data flows through the pipeline from one cmdlet to another


#region - Get all running processes
Get-Process 
#endregion


#region Get all running processes and pass them through the pipeline
Get-Process | 
    # Select only specific properties
    Select-Object Name, Id, CPU, WorkingSet |
    Select-Object -First 10
#endregion


#region Get all running processes and pass them through the pipeline
#          post process the list of items
Get-Process | 
    # Select only specific properties
    Select-Object Name, Id, CPU, WorkingSet |
    # Sort by CPU usage in descending order
    Sort-Object CPU -Descending | 
    Select-Object -First 10
#endregion


#region The Pipeline is VERY capable
Get-Process | 
    # Select only specific properties
    Select-Object Name, Id, CPU, WorkingSet |
    # Sort by CPU usage in descending order
    Sort-Object CPU -Descending |
    # Take only the top 5 processes
    Select-Object -First 5
#endregion

Write-Host "The pipeline allows data to flow between cmdlets, transforming it at each step."