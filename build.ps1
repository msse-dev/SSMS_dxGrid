[CmdletBinding()]
param([string]$Configuration = "Release")

$ErrorActionPreference = "Stop"
$workspace = $PSScriptRoot
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) { throw "vswhere.exe was not found." }

$ssmsRoot = & $vswhere -products Microsoft.VisualStudio.Product.SSMS -latest -property installationPath
if (-not $ssmsRoot) { throw "SSMS 22 was not found." }

$project = Join-Path $workspace "src\ExcelGrid\ExcelGrid.csproj"
dotnet build $project -c $Configuration --no-restore -p:SsmsRoot="$ssmsRoot" -p:_EnableDefaultWindowsPlatform=false
if ($LASTEXITCODE -ne 0) { throw "Build failed." }

$tests = Join-Path $workspace "tests\ExcelGrid.Tests\ExcelGrid.Tests.csproj"
$env:EXCELGRID_SSMS_ROOT = $ssmsRoot
dotnet run --project $tests -c $Configuration --no-restore -p:SsmsRoot="$ssmsRoot" -p:_EnableDefaultWindowsPlatform=false
if ($LASTEXITCODE -ne 0) { throw "Tests failed." }

$stage = Join-Path $workspace "artifacts\stage"
$artifact = Join-Path $workspace "artifacts\ExcelGrid.Ssms22.vsix"
if (Test-Path $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $workspace "extension\extension.vsixmanifest") -Destination $stage
Copy-Item -LiteralPath (Join-Path $workspace "extension\[Content_Types].xml") -Destination $stage
Copy-Item -LiteralPath (Join-Path $workspace "src\ExcelGrid\bin\$Configuration\net48\ExcelGrid.Ssms.dll") -Destination $stage
Get-ChildItem (Join-Path $workspace "src\ExcelGrid\bin\$Configuration\net48") -Filter "DevExpress*.dll" | Copy-Item -Destination $stage

$zip = Join-Path $workspace "artifacts\ExcelGrid.Ssms22.zip"
if (Test-Path $zip) { Remove-Item -LiteralPath $zip -Force }
if (Test-Path $artifact) { Remove-Item -LiteralPath $artifact -Force }
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zip
Move-Item -LiteralPath $zip -Destination $artifact
Write-Host "Built $artifact"
