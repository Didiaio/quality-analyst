<# ===============================
   QA Report: Category Count Report
   =============================== #>

# Import and parse the CSV data relative to the script location
$scriptFolder = $PSScriptRoot
$projectRoot = Split-Path -Path $scriptFolder -Parent
$dataFile = Join-Path -Path $scriptFolder -ChildPath "data\QADataReport.csv"
$data = Import-Csv -Path $dataFile
$totalAudits = $data.Count

$outputDir = Join-Path -Path $projectRoot -ChildPath 'output'
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$categoryCsvPath = Join-Path -Path $outputDir -ChildPath "QAReport_CategoryCount_CategoryOccurrences.csv"
$agentCategoryCsvPath = Join-Path -Path $outputDir -ChildPath "QAReport_CategoryCount_AgentCategoryAssignment.csv"
$agentOccurrencesCsvPath = Join-Path -Path $outputDir -ChildPath "QAReport_CategoryCount_AgentOccurrences.csv"

Write-Output "====================================="
Write-Output "CATEGORY COUNT REPORT"
Write-Output "====================================="
Write-Output ""
Write-Output "TOTAL AUDITS: $totalAudits"
Write-Output ""

# Count occurrences of each category
$propertyName = ($data | Get-Member -MemberType NoteProperty | Where-Object {$_.Name -like '*Category*'}).Name

if ($propertyName) {
    $categoryGroups = $data | Group-Object -Property $propertyName | Sort-Object Name
    $categoryRows = $categoryGroups | Select-Object @{Name='Category';Expression={$_.Name}}, @{Name='Count';Expression={$_.Count}}

    Write-Output "CATEGORY OCCURRENCES:"
    Write-Output "-------- -----------"
    Write-Output ""
    $categoryRows | ForEach-Object {
        Write-Output "$($_.Category) appears $($_.Count) times"
    }

    $categoryRows | Export-Csv -Path $categoryCsvPath -NoTypeInformation
    Write-Output "Category occurrences exported to: $categoryCsvPath"

} else {
    Write-Output "Category data not found!"
}

Write-Output ""
Write-Output "====================================="
Write-Output "AGENT CATEGORY ASSIGNMENT"
Write-Output "====================================="
Write-Output ""

$agentCategoryGroups = $data |
    Group-Object -Property AgentID |
    Sort-Object @{Expression={
        $numbers = $_.Group | Select-Object -ExpandProperty $propertyName |
            ForEach-Object { [int]($_ -replace '[^\d]','') }
        $numbers | Sort-Object | Select-Object -First 1
    }}, Name
Write-Output "AGENT CATEGORY LIST:"
Write-Output "----- -------- ----"
Write-Output ""
$agentCategoryRows = @()
$agentCategoryGroups | ForEach-Object {
    $agentID = $_.Name
    $agentName = $_.Group[0].AgentName
    $categories = ($_.Group | Select-Object -ExpandProperty $propertyName |
        Sort-Object {[int]($_ -replace '[^\d]','')} -Unique) -join ", "
    Write-Output "AgentID: $agentID | AgentName: $agentName | Category: $categories"
    $agentCategoryRows += [PSCustomObject]@{
        AgentID = $agentID
        AgentName = $agentName
        Category = $categories
    }
}

$agentCategoryRows | Export-Csv -Path $agentCategoryCsvPath -NoTypeInformation
Write-Output "Agent category assignment exported to: $agentCategoryCsvPath"

Write-Output ""
Write-Output "====================================="
Write-Output "AGENT APPEARANCE REPORT"
Write-Output "====================================="
Write-Output ""

# Count occurrences of each agent
$agentGroups = $data | Group-Object -Property AgentID | Sort-Object Count -Descending
Write-Output "AGENT OCCURRENCES:"
Write-Output "----- -----------"
Write-Output ""
$agentRows = @()
$agentGroups | ForEach-Object {
    $agentName = $_.Group[0].AgentName
    Write-Output "AgentID: $($_.Name) | AgentName: $agentName | Appears: $($_.Count) times"
    $agentRows += [PSCustomObject]@{
        AgentID = $_.Name
        AgentName = $agentName
        Appears = $_.Count
    }
}


Write-Output ""
Write-Output "====================================="
$agentRows | Export-Csv -Path $agentOccurrencesCsvPath -NoTypeInformation
Write-Output "Report exported to: $agentOccurrencesCsvPath"

Write-Output ""
Write-Output "Report generation completed successfully!"
