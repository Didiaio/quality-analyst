<# ===============================
   QA Report: Assessment Outcome Count Report
   =============================== #>

$scriptFolder = $PSScriptRoot
$projectRoot = Split-Path -Path $scriptFolder -Parent
$csvPath = Join-Path -Path $scriptFolder -ChildPath 'data\QADataCopy.csv'

$headers = @(
    'ConsultantName',
    'MonitorTypeChannel',
    'TeamLeader',
    'MonitorType',
    'CallDigitalID',
    'AssessmentOutcome'
)

$data = Import-Csv -Path $csvPath -Header $headers | Select-Object -Skip 1
foreach ($row in $data) {
    foreach ($header in $headers) {
        if ($row.PSObject.Properties[$header]) {
            $row.$header = ($row.$header -as [string]).Trim()
        }
    }
}

$totalAssessments = $data.Count

$outputDir = Join-Path -Path $projectRoot -ChildPath 'output'
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
}
$outcomeCsvPath = Join-Path -Path $outputDir -ChildPath 'QAReport_CategoryCount_AssessmentOccurrences.csv'
$consultantCsvPath = Join-Path -Path $outputDir -ChildPath 'QAReport_CategoryCount_ConsultantAssignment.csv'
$consultantOccurrencesCsvPath = Join-Path -Path $outputDir -ChildPath 'QAReport_CategoryCount_ConsultantOccurrences.csv'

Write-Output "====================================="
Write-Output "ASSESSMENT OUTCOME COUNT REPORT"
Write-Output "====================================="
Write-Output ""
Write-Output "TOTAL ASSESSMENTS: $totalAssessments"
Write-Output ""

$outcomeGroups = $data | Group-Object -Property AssessmentOutcome | Sort-Object Count -Descending
$outcomeRows = $outcomeGroups | Select-Object @{Name='AssessmentOutcome';Expression={$_.Name}}, @{Name='Count';Expression={$_.Count}}

Write-Output "ASSESSMENT OUTCOME OCCURRENCES:"
Write-Output "----------------------------"
Write-Output ""
$outcomeRows | ForEach-Object {
    Write-Output "$($_.AssessmentOutcome) appears $($_.Count) times"
}

$outcomeRows | Export-Csv -Path $outcomeCsvPath -NoTypeInformation
Write-Output ""
Write-Output "====================================="
Write-Output "EXPORT SUMMARY"
Write-Output "====================================="
Write-Output "Assessment outcome occurrences exported to: $outcomeCsvPath"

Write-Output ""
Write-Output "====================================="
Write-Output "CONSULTANT ASSESSMENT ASSIGNMENT"
Write-Output "====================================="
Write-Output ""

$consultantGroups = $data | Group-Object -Property ConsultantName | Sort-Object Count -Descending
$consultantRows = @()
$consultantGroups | ForEach-Object {
    $consultantName = $_.Name
    $teamLeader = ($_.Group | Select-Object -First 1).TeamLeader
    $outcomes = ($_.Group | Select-Object -ExpandProperty AssessmentOutcome | Sort-Object -Unique) -join ", "
    Write-Output "Consultant: $consultantName | TeamLeader: $teamLeader | Outcomes: $outcomes"
    $consultantRows += [PSCustomObject]@{
        ConsultantName = $consultantName
        TeamLeader = $teamLeader
        Outcomes = $outcomes
    }
}

$consultantRows | Export-Csv -Path $consultantCsvPath -NoTypeInformation
Write-Output ""
Write-Output "====================================="
Write-Output "EXPORT SUMMARY"
Write-Output "====================================="
Write-Output "Consultant assessment assignment exported to: $consultantCsvPath"

Write-Output ""
Write-Output "====================================="
Write-Output "CONSULTANT OCCURRENCE REPORT"
Write-Output "====================================="
Write-Output ""

$consultantOccurrenceRows = $consultantGroups | ForEach-Object {
    [PSCustomObject]@{
        ConsultantName = $_.Name
        Assessments = $_.Count
        TeamLeader = ($_.Group | Select-Object -First 1).TeamLeader
    }
}

$consultantOccurrenceRows | ForEach-Object {
    Write-Output "Consultant: $($_.ConsultantName) | TeamLeader: $($_.TeamLeader) | Assessments: $($_.Assessments)"
}

$consultantOccurrenceRows | Export-Csv -Path $consultantOccurrencesCsvPath -NoTypeInformation
Write-Output ""
Write-Output "====================================="
Write-Output "EXPORT SUMMARY"
Write-Output "====================================="
Write-Output "Consultant occurrence report exported to: $consultantOccurrencesCsvPath"

Write-Output ""
Write-Output "Report generation completed successfully!"
