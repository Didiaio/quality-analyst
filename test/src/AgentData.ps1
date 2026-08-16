$data = @(
    [PSCustomObject]@{Agent="Alice"; Audits=8; Percentage=(8/10*100)}
    [PSCustomObject]@{Agent="Bob"; Audits=6; Percentage=(6/10*100)}
    [PSCustomObject]@{Agent="Carla"; Audits=9; Percentage=(9/10*100)}
    [PSCustomObject]@{Agent="David"; Audits=7; Percentage=(7/10*100)}
    [PSCustomObject]@{Agent="Emma"; Audits=10; Percentage=(10/10*100)}
)

$scriptFolder = $PSScriptRoot
# repo root is two levels up from test/src
$repoRoot = Resolve-Path -Path (Join-Path -Path $scriptFolder -ChildPath '..\..')
$outputDir = Join-Path -Path $repoRoot -ChildPath 'output'
if (-not (Test-Path -Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
$excelPath = Join-Path -Path $outputDir -ChildPath 'AgentScores.xlsx'

$data | Export-Excel -Path $excelPath -WorksheetName 'Scores' -AutoSize

# Also create a CSV copy in the repo output folder
$csvPath = [System.IO.Path]::ChangeExtension($excelPath, '.csv')
$data | Export-Csv -Path $csvPath -NoTypeInformation