# Requires ImportExcel module
# Install once with: Install-Module -Name ImportExcel -Scope CurrentUser

# Determine script and repository locations in a robust way so scripts work after folder moves
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$repoRoot = Resolve-Path -Path $scriptDir

# Ensure output directory exists at repo root
$outputDir = Join-Path -Path $repoRoot -ChildPath 'output'
if (-not (Test-Path -Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# Look for likely CSV inputs relative to the repository root
$candidates = @(
    'test\data\QADataCopy.csv',
    'test\data\QADataReport.csv',
    'script\data\QAData.csv',
    'script\data\QADataReport.csv',
    'data\QAData.csv',
    'data\QADataCopy.csv'
)

$csvPath = $null
foreach ($rel in $candidates) {
    $candidate = Join-Path -Path $repoRoot -ChildPath $rel
    if (Test-Path -Path $candidate) { $csvPath = $candidate; break }
}

if (-not $csvPath) {
    Write-Error "No input CSV found. Searched: $($candidates -join '; ')"
    exit 1
}

# Build output Excel path using same base name, placed in the output folder
$excelName = [System.IO.Path]::ChangeExtension((Split-Path -Leaf $csvPath), '.xlsx')
$excelPath = Join-Path -Path $outputDir -ChildPath $excelName

# Remove existing output if present
if (Test-Path -Path $excelPath) { Remove-Item -Path $excelPath -Force }

# Convert CSV -> Excel (requires ImportExcel)
Import-Csv -Path $csvPath |
    Export-Excel -Path $excelPath -WorksheetName 'QA_Report' -AutoSize -BoldTopRow

Write-Output "Converted: $csvPath -> $excelPath"
