# Demo 6: Combining Pipeline Operations with ForEach-Object
# Shows how to perform complex operations on each pipeline object

#region Find large files in the Windows directory
$largeFiles = Get-ChildItem -Path $env:windir -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $_.Length -gt 50MB } |
    Select-Object -First 10 |
    # Process each file with ForEach-Object
    ForEach-Object {
        # Create custom object with calculated properties
        [PSCustomObject]@{
            Name         = $_.Name
            Directory    = $_.DirectoryName
            SizeMB       = [math]::Round($_.Length / 1MB, 2)
            LastModified = $_.LastWriteTime
            Age          = [math]::Round((New-TimeSpan -Start $_.LastWriteTime -End (Get-Date)).TotalDays, 0)
            Extension    = $_.Extension
        }
    } |
    Sort-Object SizeMB -Descending
#endregion

# Display results
$largeFiles | Out-GridView -PassThru

Write-Host "ForEach-Object allows you to run complex script blocks against each pipeline object."
Write-Host "It's perfect for transformations that are too complex for Select-Object."