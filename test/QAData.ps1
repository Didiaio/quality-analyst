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

    foreach ($row in $data) {
        foreach ($header in $headers) {
            if ($row.PSObject.Properties[$header]) {
                $row.$header = ($row.$header -as [string]).Trim()
            }
        }

        $missingFields = @()
        if ([string]::IsNullOrWhiteSpace($row.ConsultantName)) { $missingFields += 'ConsultantName' }
        if ([string]::IsNullOrWhiteSpace($row.CallDigitalID)) { $missingFields += 'CallDigitalID' }
        if ([string]::IsNullOrWhiteSpace($row.AssessmentOutcome)) { $missingFields += 'AssessmentOutcome' }

        if ($missingFields.Count -gt 0) {
            Write-Warning "Missing data for call/digital ID '$($row.CallDigitalID)' - missing: $($missingFields -join ', ')"
        }
    }

    return $data
}

<# ===============================
   Module 2: Reporting
   =============================== #>

function Generate-QAReport {
    param(
        [array]$Data,
        [string]$OutputFile = "QAReport.xlsx"
    )

    $summary = $Data | Group-Object ConsultantName | ForEach-Object {
        $group = $_.Group
        [PSCustomObject]@{
            ConsultantName = $_.Name
            TeamLeader = ($group | Select-Object -First 1).TeamLeader
            AssessmentCount = $group.Count
            AcceptableCount = ($group | Where-Object { $_.AssessmentOutcome -ieq 'Acceptable' }).Count
            SignificantCustomerHarmCount = ($group | Where-Object { $_.AssessmentOutcome -like '1*' }).Count
            SignificantCustomerExperienceIssueCount = ($group | Where-Object { $_.AssessmentOutcome -like '2*' }).Count
            SignificantProcessIssueCount = ($group | Where-Object { $_.AssessmentOutcome -like '3*' }).Count
            InternalIssueCount = ($group | Where-Object { $_.AssessmentOutcome -like '4*' }).Count
        }
    }

    if (Get-Command -Name Export-Excel -ErrorAction SilentlyContinue) {
        $summary | Export-Excel $OutputFile -AutoSize -BoldTopRow

        # Also export a CSV copy into the same output folder
        try {
            $csvOutput = [System.IO.Path]::ChangeExtension($OutputFile, '.csv')
            $summary | Export-Csv -Path $csvOutput -NoTypeInformation
            Write-Output "Report generated: $OutputFile (CSV: $csvOutput)"
        }
        catch {
            Write-Warning "Could not export CSV copy: $_"
        }
    }
    else {
        $csvOutput = [System.IO.Path]::ChangeExtension($OutputFile, '.csv')
        $summary | Export-Csv -Path $csvOutput -NoTypeInformation
        Write-Warning "ImportExcel module not found. Exported report as CSV instead: $csvOutput"
        Write-Output ""
        Write-Output "====================================="
        Write-Output "EXPORT SUMMARY"
        Write-Output "====================================="
        Write-Output "Report generated: $csvOutput"
        Write-Output "====================================="
    }
}

<# ===============================
   Module 3: Alerts
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

    $lowPerformers = $Data | Where-Object {
        $outcome = ($_.AssessmentOutcome -as [string]).Trim()
        $outcome -and ($outcome -ne 'Acceptable')
    }

    if ($lowPerformers -and $lowPerformers.Count -gt 0) {
        $agentList = $lowPerformers |
            ForEach-Object { "{0} ({1}) - {2}" -f $_.ConsultantName, $_.CallDigitalID, $_.AssessmentOutcome } |
            Out-String

        $body = @"
Hi Team,

Please review the following assessment(s) with non-acceptable outcomes:

$agentList

Please assess the situation and provide necessary support to help improve performance.
See attached report for more details.

For any questions or further details, feel free to reach out.

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
$outputFile = Join-Path $outputDir 'QAReport.xlsx'

$data = Import-AndValidateQAData -InputFile $dataFile
Generate-QAReport -Data $data -OutputFile $outputFile
Send-QAAlerts -Data $data -SupervisorEmail "supervisor@web.com"