# Install module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}

# Create a demo output folder
$demoFolder = ".\_Output"
if (-not (Test-Path $demoFolder)){New-Item -ItemType Directory -Path $demoFolder -Force | Out-Null}
