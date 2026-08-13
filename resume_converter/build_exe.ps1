<#
Build the Windows onedir GUI package with the active Python environment.
The current internal test package includes the project .env as a resource.
#>

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot

python -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller is missing. Run: python -m pip install PyInstaller"
}

python -m PyInstaller --noconfirm --clean resume_converter.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed."
}

$distRoot = Join-Path $projectRoot "dist"
$distDirectory = Get-ChildItem -LiteralPath $distRoot -Directory |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $distDirectory) {
    throw "The build completed without a distribution directory."
}

Write-Host ""
Write-Host "Build completed: $($distDirectory.FullName)"
Write-Host "Internal test build: the project .env is bundled into the application resources."
