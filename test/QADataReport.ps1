<# ===============================
   Module 1: Data Import & Validation
   =============================== #>

function Import-AndValidateQAData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputFile
    )

    $headers = @(
        'ConsultantName',
        'MonitorTypeChannel',
        'TeamLeader',
        'MonitorType',
        'CallDigitalID',
        'AssessmentOutcome'
    )

    $data = Import-Csv -Path $InputFile -Header $headers | Select-Object -Skip 1

    $cleanedData = @()
    foreach ($row in $data) {
        foreach ($header in $headers) {
            if ($row.PSObject.Properties[$header]) {
                $row.$header = ($row.$header -as [string]).Trim()
            }
        }

        if ([string]::IsNullOrWhiteSpace($row.ConsultantName) -or [string]::IsNullOrWhiteSpace($row.CallDigitalID) -or [string]::IsNullOrWhiteSpace($row.AssessmentOutcome)) {
            Write-Warning "Skipping incomplete row for consultant '$($row.ConsultantName)' call/digital ID '$($row.CallDigitalID)'"
            continue
        }

        $cleanedData += $row
    }

    return $cleanedData
}

<# ===============================
   Module 2: Outcome & Consultant Analysis
   =============================== #>

function Get-OutcomeAnalysis {
    param(
        [array]$Data
    )

    return $Data |
        Group-Object -Property AssessmentOutcome |
        Sort-Object Count -Descending |
        Select-Object @{Name='AssessmentOutcome';Expression={$_.Name}}, @{Name='Count';Expression={$_.Count}}
}

function Get-ConsultantAnalysis {
    param(
        [array]$Data
    )

    return $Data |
        Group-Object -Property ConsultantName |
        Select-Object @{Name='ConsultantName';Expression={$_.Name}},
                      @{Name='TeamLeader';Expression={ ($_.Group | Select-Object -First 1).TeamLeader }},
                      @{Name='AssessmentCount';Expression={$_.Count}},
                      @{Name='Outcomes';Expression={ ($_.Group | Select-Object -ExpandProperty AssessmentOutcome | Sort-Object -Unique) -join ', ' }} |
        Sort-Object AssessmentCount -Descending
}

function Get-TeamLeaderAnalysis {
    param(
        [array]$Data
    )

    return $Data |
        Group-Object -Property TeamLeader |
        Select-Object @{Name='TeamLeader';Expression={$_.Name}}, @{Name='AssessmentCount';Expression={$_.Count}} |
        Sort-Object AssessmentCount -Descending
}

<# ===============================
   Module 3: Reporting
   =============================== #>

function Generate-QAReport {
    param(
        [array]$Data,
        [string]$OutputFile = "QADataReport_Summary.csv"
    )

    $outcomeAnalysis = Get-OutcomeAnalysis -Data $Data
    $consultantAnalysis = Get-ConsultantAnalysis -Data $Data
    $teamLeaderAnalysis = Get-TeamLeaderAnalysis -Data $Data

    $totalAssessments = $Data.Count
    $totalConsultants = ($Data | Select-Object -Unique ConsultantName).Count

    Write-Output "====================================="
    Write-Output "QUALITY ASSURANCE DATA REPORT"
    Write-Output "====================================="
    Write-Output ""
    Write-Output "SUMMARY STATISTICS"
    Write-Output "------------------"
    Write-Output "Total Assessments: $totalAssessments"
    Write-Output "Total Consultants: $totalConsultants"
    Write-Output ""
    Write-Output "ASSESSMENT OUTCOME BREAKDOWN"
    Write-Output "----------------------------"
    $outcomeAnalysis | Format-Table -AutoSize
    Write-Output ""
    Write-Output "CONSULTANT ASSESSMENT FREQUENCY"
    Write-Output "------------------------------"
    $consultantAnalysis | Format-Table -Property ConsultantName, TeamLeader, AssessmentCount, Outcomes -AutoSize
    Write-Output ""
    Write-Output "TEAM LEADER SUMMARY"
    Write-Output "------------------"
    $teamLeaderAnalysis | Format-Table -AutoSize
    Write-Output ""

    try {
        $categoryFile = $OutputFile -replace '\.[^.]*$', '_Outcome.csv'
        $agentFile = $OutputFile -replace '\.[^.]*$', '_Consultant.csv'
        $teamLeaderFile = $OutputFile -replace '\.[^.]*$', '_TeamLeader.csv'

        $outcomeAnalysis | Export-Csv -Path $categoryFile -NoTypeInformation
        $consultantAnalysis | Export-Csv -Path $agentFile -NoTypeInformation
        $teamLeaderAnalysis | Export-Csv -Path $teamLeaderFile -NoTypeInformation

        Write-Output "====================================="
        Write-Output "EXPORT SUMMARY"
        Write-Output "====================================="
        Write-Output "Categories exported to: $categoryFile"
        Write-Output "Consultants exported to: $agentFile"
        Write-Output "Team leaders exported to: $teamLeaderFile"
        Write-Output ""
    }
    catch {
        Write-Warning "Could not export to CSV: $_"
    }
}

<# ===============================
   Module 4: Alerts
   =============================== #>

function Send-QAAlerts {
    param(
        [array]$Data,
        [string]$SupervisorEmail = "supervisor@company.com",
        [string]$SmtpServer = $null
    )

    if (-not $SmtpServer -or [string]::IsNullOrWhiteSpace($SmtpServer)) {
        if (Get-Variable -Name PSEmailServer -Scope Global -ErrorAction SilentlyContinue) {
            $SmtpServer = (Get-Variable -Name PSEmailServer -Scope Global).Value
        }
        elseif ($env:PSEMAILSERVER) {
            $SmtpServer = $env:PSEMAILSERVER
        }
    }

    $nonAcceptableAssessments = $Data | Where-Object {
        $outcome = ($_.AssessmentOutcome -as [string]).Trim()
        $outcome -and $outcome -ne 'Acceptable'
    }

    if ($nonAcceptableAssessments.Count -gt 0) {
        $assessmentList = $nonAcceptableAssessments |
            ForEach-Object { "{0} ({1}) - {2}" -f $_.ConsultantName, $_.CallDigitalID, $_.AssessmentOutcome } |
            Out-String

        $body = @"
Hi Team,

Please review the following assessments with non-acceptable outcomes:

$assessmentList

Please assess the situation and provide necessary support to help improve performance.
See attached report for more details.

Thank you,
QA Team
"@

        $subject = 'QA Alert: Non-acceptable assessment outcomes'

        if ($SmtpServer -and -not [string]::IsNullOrWhiteSpace($SmtpServer)) {
            Send-MailMessage -To $SupervisorEmail -Subject $subject -Body $body -SmtpServer $SmtpServer -From 'qa-team@company.com'
            Write-Output "Alert email sent via SMTP server: $SmtpServer"
        }
        else {
            Write-Warning "No SMTP server specified. Falling back to opening mail compose in the browser. To send automatically, provide -SmtpServer or set the global variable `$PSEmailServer` or environment variable PSEMAILSERVER."
            $to = [uri]::EscapeDataString($SupervisorEmail)
            $subjectEnc = [uri]::EscapeDataString($subject)
            $bodyEncoded = [uri]::EscapeDataString($body)
            $composeUrl = "https://outlook.live.com/owa/?path=/mail/action/compose&to=$to&subject=$subjectEnc&body=$bodyEncoded"
            Start-Process $composeUrl
        }
    }
}

<# ===============================
   Main Script Flow
   =============================== #>

$scriptFolder = $PSScriptRoot
$projectRoot = Split-Path -Path $scriptFolder -Parent
$dataFile = Join-Path $scriptFolder 'data\QADataCopy.csv'
$outputDir = Join-Path -Path $projectRoot -ChildPath 'output'
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
}
$outputFile = Join-Path $outputDir 'QADataReport_Summary.csv'

$data = Import-AndValidateQAData -InputFile $dataFile
Generate-QAReport -Data $data -OutputFile $outputFile
Send-QAAlerts -Data $data -SupervisorEmail 'supervisor@web.com'