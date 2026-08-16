function Get-PropertyValue {
    param(
        $Row,
        [string[]]$Candidates
    )
    foreach ($name in $Candidates) {
        if ($null -ne $Row.PSObject.Properties[$name]) { return $Row.$name }
    }
    return $null
}

function Import-AndNormalize-QAData {
    param(
        [string]$InputFile
    )

    if (-not $InputFile) {
        $InputFile = Join-Path -Path $PSScriptRoot -ChildPath 'data\QAData.csv'
    }

    $raw = Import-Csv -Path $InputFile
    $normalized = @()

    foreach ($row in $raw) {
        $c = [PSCustomObject]@{
            ConsultantName = $null
            MonitorTypeChannel = $null
            TeamLeader = $null
            MonitorType = $null
            CallDigitalID = $null
            AssessmentOutcome = $null
            AgentID = $null
            AgentName = $null
            Score = $null
        }

        $c.ConsultantName = Get-PropertyValue -Row $row -Candidates @('ConsultantName','Consultant Name')
        $c.MonitorTypeChannel = Get-PropertyValue -Row $row -Candidates @('MonitorTypeChannel','Monitor Type Channel')
        $c.TeamLeader = Get-PropertyValue -Row $row -Candidates @('TeamLeader','Team Leader')
        $c.MonitorType = Get-PropertyValue -Row $row -Candidates @('MonitorType','Monitor Type')
        $c.CallDigitalID = Get-PropertyValue -Row $row -Candidates @('CallDigitalID','Call/Digital ID','Call Digital ID')
        $c.AssessmentOutcome = Get-PropertyValue -Row $row -Candidates @('AssessmentOutcome','Assessment Outcome')

        $c.AgentID = Get-PropertyValue -Row $row -Candidates @('AgentID','Agent ID')
        $c.AgentName = Get-PropertyValue -Row $row -Candidates @('AgentName','Agent Name')
        $c.Score = Get-PropertyValue -Row $row -Candidates @('Score')

        # Trim string fields
        foreach ($prop in $c.PSObject.Properties.Name) {
            if ($c.PSObject.Properties[$prop]) {
                $val = $c.$prop
                if ($val -ne $null) { $c.$prop = ($val -as [string]).Trim() }
            }
        }

        # If the dataset is agent-based, but ConsultantName is empty, map AgentName -> ConsultantName for downstream compatibility
        if (-not $c.ConsultantName -and $c.AgentName) { $c.ConsultantName = $c.AgentName }
        # If CallDigitalID empty but AgentID exists, set to AgentID
        if (-not $c.CallDigitalID -and $c.AgentID) { $c.CallDigitalID = $c.AgentID }

        $normalized += $c
    }

    return $normalized
}
