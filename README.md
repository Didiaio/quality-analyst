# QUALITY ASSURANCE TASKLIST

## Script Execution

- Source Files
  - QAData.ps1
  - QADataReport.ps1
  - QAReport_CategoryCount.ps1

## Quality Audits

Agent Call Listening Audits:

- Quality Assurance Audit Form
  - Category 1 - Fail (High Priority)
  - Category 2 - Fail (Medium Priority)
  - Category 3 - Fail (Low Priority)
  - Category 4 - Fail (Internal)
  - Acceptance - Pass

Quality Audit Checks:

- Excel Sheet Weekly Data Checks
  - Retentions
  - Amendments
  - Vulnerability
  - Complaints
  - Red Alerts
  - Dip Checks
  - Integrity Checks

## Quality Assurance Training

Agent Feedback and Coaching:

- Weekly Coaching Sessions
- Monthly Performance Reviews
- Individual Development Plans
- Real-time Feedback Mechanisms
- Care Bank for agent progress development

## Quality Assurance Reporting

Daily Checks:

- Daily Quality Assurance Report
  - Summary of Audits Conducted
  - Key Findings and Trends
  - Actionable Insights for Improvement

Weekly Reports:

- Weekly Quality Assurance Report
  - Comprehensive Analysis of Quality Metrics
  - Performance Trends and Patterns
  - Recommendations for Process Enhancements

Monthly Reports:

- Monthly Quality Assurance Report
  - In-depth Review of Quality Performance
  - Identification of Root Causes for Issues
  - Strategic Recommendations for Long-term Improvement

## Reporting and Analytics

- Quality Assurance Dashboard
  - Real-time Visualization of Key Metrics
  - Customizable Views for Different Stakeholders
  - Automated Alerts for Critical Issues

- Reporting Expectations
  - Timely Submission of Reports
  - Accuracy and Completeness of Data
  - Clear Communication of Findings and Recommendations

- Reporting Formats
  - Standardized Templates for Consistency
  - Clear and Concise Presentation of Information
  - Use of Visual Aids to Enhance Understanding
  - Email and Presentation Formats for Different Audiences

## REPORTS

| Report Name | Frequency | Description | Format |
| ----------- | --------- | ----------- | ------ |
| Daily Quality Assurance Report | Daily | Provides a summary of audits conducted, key findings, and actionable insights for improvement. | Email Relevant Departments and Direct Team |
| Weekly Quality Assurance Report | Weekly | Provides a comprehensive analysis of quality metrics, performance trends, and recommendations for process enhancements. | Email Relevant Department and Operations Team |
| Monthly Quality Assurance Report | Monthly | Offers an in-depth review of quality performance, identification of root causes for issues, and strategic recommendations for long-term improvement. | Email Relevant Department and Operations Team |

## Running the Export Scripts

Use the following commands from the repository root to generate outputs in the `output/` folder:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\export.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\script\QAData.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\script\QADataReport.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\script\QAReport_CategoryCount.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\QAData.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\QADataReport.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\QAReport_CategoryCount.ps1
```

Generated files go to `output/` and include Excel and CSV exports such as:
- `QADataCopy.xlsx`
- `QAReport.xlsx`
- `QAReport.csv`
- `QADataReport_Summary_*.csv`
- `QAReport_CategoryCount_*.csv`

> Note: Excel export requires the PowerShell `ImportExcel` module. If it's not installed, the scripts still create CSV output where supported.
