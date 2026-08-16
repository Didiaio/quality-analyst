<#
.SYNOPSIS
	Send an Excel report via email, optionally including a small summary extracted from the sheet.

.DESCRIPTION
	Attaches the provided Excel file and sends it to the specified recipient(s). If the ImportExcel
	module is available the script will create a short summary (row count and numeric-score average
	if a "Score" column exists) and include that in the message body. The script will prompt for
	sender address when not supplied.

.PARAMETER ExcelPath
	Path to the Excel file to attach. Defaults to QAReport.xlsx in the current directory.

.PARAMETER To
	Recipient email address(es); comma-separated or an array. If omitted the script will prompt.

.PARAMETER From
	Sender email address. If omitted the script will prompt.

.PARAMETER SmtpServer
	SMTP server to use. If omitted the script will look for the global $PSEmailServer variable
	or the PSEMAILSERVER environment variable. If still not present, the script will open a
	browser compose window for Outlook.com as a manual fallback.

.EXAMPLE
	.\ReportData.ps1 -ExcelPath QAReport.xlsx -To supervisor@company.com -SmtpServer smtp.company.com
#>

param(
	[string]$ExcelPath = ".\Test\QAReport.xlsx",
	[Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
	[string[]]$To,
	# [string]$From,
	[string]$SmtpServer,
	[string]$Subject,
	[string]$Body,
	[switch]$AutoInstallImportExcel
)

function Resolve-SmtpServer {
	param([string]$SmtpServerParam)
	if ($SmtpServerParam -and -not [string]::IsNullOrWhiteSpace($SmtpServerParam)) { return $SmtpServerParam }
	if (Get-Variable -Name PSEmailServer -Scope Global -ErrorAction SilentlyContinue) {
		return (Get-Variable -Name PSEmailServer -Scope Global).Value
	}
	if ($env:PSEMAILSERVER) { return $env:PSEMAILSERVER }
	return $null
}

function Ensure-ImportExcel {
	param([switch]$AutoInstall)
	if (Get-Module -ListAvailable -Name ImportExcel) { return $true }
	Write-Host "ImportExcel module not found." -ForegroundColor Yellow
	if (-not $AutoInstall) {
		$ans = Read-Host "Install ImportExcel from PSGallery now? (Y/N)"
		if ($ans -notin @('Y','y','Yes','yes')) { return $false }
	}
	try {
		Write-Host "Installing ImportExcel (this may prompt for confirmation)..." -ForegroundColor Cyan
		Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
		Import-Module ImportExcel -ErrorAction Stop
		Write-Host "ImportExcel installed and imported." -ForegroundColor Green
		return $true
	}
	catch {
		Write-Warning "Failed to install/import ImportExcel: $_"
		return $false
	}
}

function Prompt-ForEmail([string]$PromptMessage) {
	while ($true) {
		$val = Read-Host $PromptMessage
		if (-not [string]::IsNullOrWhiteSpace($val)) { return $val }
		Write-Host "Please enter a non-empty email address." -ForegroundColor Yellow
	}
}

# Validate file exists
if (-not (Test-Path -Path $ExcelPath)) {
	Write-Error "Excel file not found: $ExcelPath"
	exit 1
}

# Prompt for To if not provided
if (-not $To -or $To.Count -eq 0) {
	$toInput = Prompt-ForEmail "Enter recipient email address(es) (comma-separated if multiple):"
	$To = $toInput -split '\s*,\s*'
}

# Prompt for From if not provided
# if (-not $From -or [string]::IsNullOrWhiteSpace($From)) {
# 	$From = Prompt-ForEmail "Enter sender (From) email address:"
# }

# Resolve SMTP server
$ResolvedSmtp = Resolve-SmtpServer -SmtpServerParam $SmtpServer

# Attempt to build a short summary using ImportExcel if available
$bodyLines = @()
$bodyLines += "Hello,"
$bodyLines += "`nPlease find the attached QA report for the following agents.`n"

if (-not $Body) {
	$importOk = $false
	if (Get-Module -ListAvailable -Name ImportExcel) {
		$importOk = $true
		Import-Module ImportExcel -ErrorAction SilentlyContinue
	}
	elseif ($AutoInstallImportExcel) {
		$importOk = Ensure-ImportExcel -AutoInstall
	}
	else {
		# Offer to install interactively
		$importOk = Ensure-ImportExcel
	}

	if ($importOk) {
		try {
			$rows = Import-Excel -Path $ExcelPath
			$rowCount = if ($rows) { $rows.Count } else { 0 }
			$bodyLines += "Rows in sheet: $rowCount"

			if ($rowCount -gt 0 -and ($rows | Get-Member -Name Score -MemberType NoteProperty -ErrorAction SilentlyContinue)) {
				# compute average if Score column exists and is numeric
				$numericScores = $rows | ForEach-Object {
					$n = $null
					if ([double]::TryParse([string]$_.Score, [ref]$n)) { $n }
				} | Where-Object { $_ -ne $null }
				if ($numericScores.Count -gt 0) {
					$avg = [Math]::Round((($numericScores | Measure-Object -Average).Average),2)
					$bodyLines += "Average Score: $avg"
				}
			}
		}
		catch {
			Write-Warning "Failed to read Excel using ImportExcel: $_. Falling back to attachment-only email."
		}
	}
	else {
		Write-Host "ImportExcel not available. The script will still attach the file; install ImportExcel to include a summary in the message body." -ForegroundColor Yellow
	}
}
else {
	# Body was provided by the caller; use it
	$bodyLines = @()
	$bodyLines += $Body
}

$bodyLines += "`nRegards,"
$bodyLines += "QA Team"

if (-not $Subject) { $Subject = "QA Report: $ExcelPath" }

$body = $bodyLines -join "`n"

if ($ResolvedSmtp) {
	Write-Host "Using SMTP server: $ResolvedSmtp" -ForegroundColor Green
	try {
		Send-MailMessage -From $From -To ($To -join ',') -Subject $Subject -Body $body -SmtpServer $ResolvedSmtp -Attachments $ExcelPath
		Write-Host "Email sent successfully to $($To -join ',')" -ForegroundColor Green
	}
	catch {
		Write-Error "Failed to send email via SMTP server $ResolvedSmtp $_"
		exit 2
	}
}
else {
	Write-Warning "No SMTP server configured. The script will open an Outlook.com compose window with To/Subject/Body prefilled. You will need to attach the file manually."
	$toEsc = [uri]::EscapeDataString(($To -join ','))
	$subjectEnc = [uri]::EscapeDataString("QA Report: $ExcelPath")
	$bodyEnc = [uri]::EscapeDataString($body)
	$composeUrl = "https://outlook.live.com/owa/?path=/mail/action/compose&to=$toEsc&subject=$subjectEnc&body=$bodyEnc"
	Start-Process $composeUrl
	Write-Host "Opened browser compose window. Attach the file manually: $ExcelPath" -ForegroundColor Yellow
}

