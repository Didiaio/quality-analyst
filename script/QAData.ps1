<# ===============================
   Module 1: Data Import & Validation
   =============================== #>

function Import-AndValidateQAData {
    param(
        [string]$InputFile
    )

    if (-not $InputFile) {
        $InputFile = Join-Path -Path $PSScriptRoot -ChildPath "data\QAData.csv"
    }

    $data = Import-Csv -Path $InputFile

    foreach ($row in $data) {
        if (-not $row.AgentID -or -not $row.Score) {
            Write-Warning "Missing data for $($row.AgentName)"
        }
        # Robust validation: treat empty/whitespace as missing, allow numeric 0, ensure numeric and within 0-100
        $agentMissing = [string]::IsNullOrWhiteSpace($row.AgentID)
        $scoreMissing = [string]::IsNullOrWhiteSpace([string]$row.Score)

        if ($agentMissing -or $scoreMissing) {
            Write-Warning "Missing data for $($row.AgentName)"
            continue
        }

        $parsedScore = 0
        if (-not [int]::TryParse($row.Score, [ref]$parsedScore)) {
            Write-Warning "Non-numeric score for $($row.AgentName): $($row.Score)"
            continue
        }

        if ($parsedScore -lt 0 -or $parsedScore -gt 100) {
            Write-Warning "Invalid score for $($row.AgentName): $parsedScore"
            continue
        }

        # If we reach here, the row is valid and no message is written
    }
    return $data
}

<# ===============================
   Module 2: Reporting
   =============================== #>

function Generate-QAReport {
    param(
        [array]$Data,
        [string]$OutputFile
    )

    if (-not $OutputFile) {
        $projectRoot = Split-Path -Path $PSScriptRoot -Parent
        $outputDir = Join-Path -Path $projectRoot -ChildPath 'output'
        if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
        $OutputFile = Join-Path -Path $outputDir -ChildPath 'QAReport.xlsx'
    }

    $summary = $Data | Group-Object AgentID | ForEach-Object {
        [PSCustomObject]@{
            AgentID  = $_.Name
            AgentName = ($_.Group | Select-Object -First 1).AgentName
            AvgScore = ($_.Group.Score | Measure-Object -Average).Average
        }
    }

    # Requires ImportExcel module
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

<# ===============================
   Module 3: Alerts
   =============================== #>

function Send-QAAlerts {
    param(
        [array]$Data,
        [int]$Threshold = 80,
        [string]$SupervisorEmail = "supervisor@company.com",
        [string]$SmtpServer = $null
    )

    # Determine SMTP server to use: explicit param > $PSEmailServer variable > environment variable
    if (-not $SmtpServer -or [string]::IsNullOrWhiteSpace($SmtpServer)) {
        if (Get-Variable -Name PSEmailServer -Scope Global -ErrorAction SilentlyContinue) {
            $SmtpServer = (Get-Variable -Name PSEmailServer -Scope Global).Value
        }
        elseif ($env:PSEMAILSERVER) {
            $SmtpServer = $env:PSEMAILSERVER
        }
    }

    $lowPerformers = $Data | Where-Object {
        $parsed = 0
        if (-not [int]::TryParse([string]$_.Score, [ref]$parsed)) { return $false }
        $parsed -lt $Threshold
    }

    if ($lowPerformers -and $lowPerformers.Count -gt 0) {
        $agentList = $lowPerformers |
            ForEach-Object { "({0}) {1} - {2}" -f $_.AgentID, $_.AgentName, $_.Score } |
            Out-String

        $body = @"
Hi Team,

Please review the following agent(s) who are below the QA score threshold of $Threshold

$agentList

Please assess the situation and provide necessary support to help improve their performance.
See attached report for more details.

For any questions or further details, feel free to reach out.

Thank you,
QA Team
"@

        $subject = "QA Alert: Agents below $Threshold"

        if ($SmtpServer -and -not [string]::IsNullOrWhiteSpace($SmtpServer)) {
            # Send using SMTP server
            Send-MailMessage -To $SupervisorEmail -Subject $subject -Body $body -SmtpServer $SmtpServer -From "qa-team@company.com"
            Write-Output "Alert email sent via SMTP server: $SmtpServer"
        }
        else {
            # No SMTP server configured: open Outlook.com compose in the default browser with prefilled To, Subject and Body
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
$outputDir = Join-Path -Path $projectRoot -ChildPath 'output'
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$dataFile = Join-Path -Path $scriptFolder -ChildPath 'data\QAData.csv'
$outputFile = Join-Path -Path $outputDir -ChildPath 'QAReport.xlsx'

$data = Import-AndValidateQAData -InputFile $dataFile
Generate-QAReport -Data $data -OutputFile $outputFile
Send-QAAlerts -Data $data -Threshold 80 -SupervisorEmail "supervisor@web.com"