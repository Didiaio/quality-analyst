<# ===============================
   Module 1: Data Import & Validation
   =============================== #>

function Import-AndValidateQAData {
    param(
        [string]$InputFile
    )

    if (-not $InputFile) {
        $InputFile = Join-Path -Path $PSScriptRoot -ChildPath "..\data\QADataReport.csv"
    }

    $data = Import-Csv -Path $InputFile

    # Clean up column names by trimming spaces and filtering empty rows
    $cleanedData = @()
    foreach ($row in $data) {
        # Skip empty rows
        if ($null -eq $row.AgentID -or [string]::IsNullOrWhiteSpace([string]$row.AgentID)) {
            continue
        }
        
        $agentID = if ($null -ne $row.AgentID) { $row.AgentID.ToString().Trim() } else { "" }
        $agentName = if ($null -ne $row.AgentName) { $row.AgentName.ToString().Trim() } else { "" }
        $category = if ($null -ne $row.Category) { $row.Category.ToString().Trim() } else { "" }
        
        # Skip if any required field is empty
        if ([string]::IsNullOrWhiteSpace($agentID) -or [string]::IsNullOrWhiteSpace($category)) {
            continue
        }
        
        $cleanedRow = [PSCustomObject]@{
            AgentID  = $agentID
            AgentName = $agentName
            Category = $category
        }
        $cleanedData += $cleanedRow
    }

    return $cleanedData
}

<# ===============================
   Module 2: Category & Agent Analysis
   =============================== #>

function Get-CategoryAnalysis {
    param(
        [array]$Data
    )

    # Extract category number and count occurrences
    $categoryCount = $Data | 
        ForEach-Object { 
            $catNum = $_.Category -replace 'Category\s+', ''
            [PSCustomObject]@{Category = $catNum}
        } | 
        Group-Object -Property Category | 
        Select-Object @{Name="Category";Expression={"Category $($_.Name)"}}, @{Name="Count";Expression={$_.Count}} | 
        Sort-Object Count -Descending

    return $categoryCount
}

function Get-AgentAnalysis {
    param(
        [array]$Data
    )

    # Count agents and their audit frequencies
    $agentCount = $Data | Group-Object -Property AgentID | Select-Object Name, @{Name="AgentName";Expression={$_.Group[0].AgentName}}, @{Name="AuditCount";Expression={$_.Count}} | Sort-Object AuditCount -Descending

    return $agentCount
}

<# ===============================
   Module 3: Reporting
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
        $OutputFile = Join-Path -Path $outputDir -ChildPath 'QADataReport_Summary.csv'
    }

    # Get category analysis
    $categoryAnalysis = Get-CategoryAnalysis -Data $Data
    
    # Get agent analysisclear
    $agentAnalysis = Get-AgentAnalysis -Data $Data
    
    # Calculate summary statistics
    $totalAudits = $Data.Count
    $totalAgents = ($Data | Select-Object -Unique AgentID).Count
    
    Write-Output "====================================="
    Write-Output "QUALITY ASSURANCE AUDIT REPORT"
    Write-Output "====================================="
    Write-Output ""
    Write-Output "SUMMARY STATISTICS"
    Write-Output "----- ----------"
    Write-Output "Total Audits Conducted: $totalAudits"
    Write-Output "Total Unique Agents: $totalAgents"
    Write-Output ""
    Write-Output ""
    Write-Output "CATEGORY BREAKDOWN"
    Write-Output "-------- ----------"
    Write-Output ""
    
    # Display category counts with better formatting
    $categoryAnalysis | Format-Table -Property Category, Count -AutoSize
    Write-Output ""
    
    Write-Output "AGENT AUDIT FREQUENCY"
    Write-Output "----- ----- ---------"
    Write-Output ""
    
    # Display agent counts sorted by audit count descending
    $agentAnalysis | Format-Table -Property Name, AgentName, AuditCount -AutoSize
    Write-Output ""
    
    # Export to CSV files (works without Excel installation)
    try {
        $categoryFile = $OutputFile -replace '\.[^.]*$', '_Category.csv'
        $agentFile = $OutputFile -replace '\.[^.]*$', '_Agent.csv'
        
        # Export category summary
        $categoryAnalysis | Export-Csv -Path $categoryFile -NoTypeInformation
        
        # Export agent summary
        $agentAnalysis | Export-Csv -Path $agentFile -NoTypeInformation
        
        Write-Output "====================================="
        Write-Output "EXPORT SUMMARY"
        Write-Output "====================================="
        Write-Output ""
        Write-Output "CSV Files Exported:"
        Write-Output "-------------------"
        Write-Output ""
        Write-Output "1. Category Analysis:"
        Write-Output "   File: $(Split-Path -Leaf $categoryFile)"
        Write-Output "   Path: $categoryFile"
        Write-Output ""
        Write-Output "2. Agent Analysis:"
        Write-Output "   File: $(Split-Path -Leaf $agentFile)"
        Write-Output "   Path: $agentFile"
        Write-Output ""
    }
    catch {
        Write-Warning "Could not export to CSV: $_"
        Write-Warning "Report displayed in console only."
    }
}

<# ===============================
   Module 3: Alerts
   =============================== #>

function Send-QAAlerts {
    param(
        [array]$Data,
        [int]$Threshold = 70,
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

$dataFile = Join-Path -Path $scriptFolder -ChildPath 'data\QADataReport.csv'
$outputFile = Join-Path -Path $outputDir -ChildPath 'QADataReport_Summary.csv'

$data = Import-AndValidateQAData -InputFile $dataFile
Generate-QAReport -Data $data -OutputFile $outputFile