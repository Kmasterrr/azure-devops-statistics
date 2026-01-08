[CmdletBinding()]
param 
(
    [string]$DataPath = "",
    [string]$OutputFormat = "HTML", # Options: HTML, JSON, Excel (requires Excel module)
    [switch]$OpenReport
)

if (-not $DataPath) {
    # Try to find the most recent data collection
    $Organization = $env:ADOS_ORGANIZATION
    if (-not $Organization) {
        Write-Error "Either provide -DataPath or set ADOS_ORGANIZATION environment variable."
        exit 1
    }
    
    $BaseDataPath = Join-Path (Join-Path $PSScriptRoot "data") $Organization
    if (Test-Path $BaseDataPath) {
        $LatestDate = Get-ChildItem -Path $BaseDataPath -Directory | Sort-Object Name -Descending | Select-Object -First 1
        if ($LatestDate) {
            $DataPath = $LatestDate.FullName
            Write-Host "Using latest data collection: $DataPath" -ForegroundColor Yellow
        }
    }
    
    if (-not $DataPath -or -not (Test-Path $DataPath)) {
        Write-Error "No data path found. Please run Collect-Statistics-Local.ps1 first or specify -DataPath."
        exit 1
    }
}

Write-Host "Generating consolidated report from: $DataPath" -ForegroundColor Cyan

# Load JSON data files
$DataFiles = @{
    Summary = Join-Path $DataPath "Collection-Summary.json"
    OrganizationStats = Join-Path $DataPath "OrganizationStatistics.json"
    ProjectStats = Join-Path $DataPath "ProjectStatistics.json"
    Users = Join-Path $DataPath "Users.json"
    Groups = Join-Path $DataPath "Groups.json"
    GroupMemberships = Join-Path $DataPath "GroupMemberships.json"
    WorkItems = Join-Path $DataPath "WorkItems.json"
    GitPullRequests = Join-Path $DataPath "GitPullRequests.json"
    GitCommits = Join-Path $DataPath "GitCommits.json"
}

$LoadedData = @{}
foreach ($key in $DataFiles.Keys) {
    $filePath = $DataFiles[$key]
    if (Test-Path $filePath) {
        try {
            $jsonContent = Get-Content -Path $filePath -Raw | ConvertFrom-Json
            $LoadedData[$key] = $jsonContent
            Write-Host "Loaded: $key" -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not load $key from $filePath : $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning "File not found: $filePath"
    }
}

# Generate report timestamp
$ReportTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$Organization = if ($LoadedData.Summary) { $LoadedData.Summary.Organization } else { $env:ADOS_ORGANIZATION }

# Calculate summary statistics
$Stats = @{
    Organization = $Organization
    ReportGenerated = $ReportTimestamp
    DataCollectionDate = if ($LoadedData.Summary) { $LoadedData.Summary.CollectionTime } else { "Unknown" }
    TotalProjects = if ($LoadedData.OrganizationStats) { $LoadedData.OrganizationStats.data.Projects } else { 0 }
    TotalUsers = if ($LoadedData.Users) { $LoadedData.Users.data.Count } else { 0 }
    TotalGroups = if ($LoadedData.Groups) { $LoadedData.Groups.data.Count } else { 0 }
    TotalWorkItems = if ($LoadedData.WorkItems) { $LoadedData.WorkItems.data.Count } else { 0 }
    TotalPullRequests = if ($LoadedData.GitPullRequests) { $LoadedData.GitPullRequests.data.Count } else { 0 }
    TotalCommits = if ($LoadedData.GitCommits) { $LoadedData.GitCommits.data.Count } else { 0 }
    TotalBuilds = if ($LoadedData.OrganizationStats) { $LoadedData.OrganizationStats.data.Builds } else { 0 }
    TotalReleases = if ($LoadedData.OrganizationStats) { $LoadedData.OrganizationStats.data.Releases } else { 0 }
}

if ($OutputFormat -eq "HTML") {
    # Generate HTML report
    $ReportPath = Join-Path $DataPath "Azure-DevOps-Report.html"
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Azure DevOps Statistics Report - $Organization</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #323130; border-bottom: 1px solid #edebe9; padding-bottom: 5px; margin-top: 30px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }
        .stat-card { background: linear-gradient(135deg, #0078d4, #106ebe); color: white; padding: 20px; border-radius: 8px; text-align: center; }
        .stat-number { font-size: 2em; font-weight: bold; margin-bottom: 5px; }
        .stat-label { font-size: 0.9em; opacity: 0.9; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background-color: #0078d4; color: white; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .info-box { background-color: #f3f2f1; border-left: 4px solid #0078d4; padding: 15px; margin: 15px 0; }
        .no-data { color: #a19f9d; font-style: italic; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #edebe9; color: #605e5c; font-size: 0.9em; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Azure DevOps Statistics Report</h1>
        
        <div class="info-box">
            <strong>Organization:</strong> $($Stats.Organization)<br>
            <strong>Data Collection Date:</strong> $($Stats.DataCollectionDate)<br>
            <strong>Report Generated:</strong> $($Stats.ReportGenerated)
        </div>

        <h2>Overview Statistics</h2>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalProjects)</div>
                <div class="stat-label">Projects</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalUsers)</div>
                <div class="stat-label">Users</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalWorkItems)</div>
                <div class="stat-label">Work Items</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalBuilds)</div>
                <div class="stat-label">Builds</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalPullRequests)</div>
                <div class="stat-label">Pull Requests</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalCommits)</div>
                <div class="stat-label">Commits</div>
            </div>
        </div>
"@

    # Add Project Statistics table if available
    if ($LoadedData.ProjectStats) {
        $htmlContent += @"
        <h2>Project Statistics</h2>
        <table>
            <tr>
                <th>Project</th>
                <th>Repositories</th>
                <th>Work Items</th>
                <th>Builds</th>
                <th>Pull Requests</th>
                <th>Commits</th>
                <th>Build Success %</th>
            </tr>
"@
        foreach ($project in $LoadedData.ProjectStats.data) {
            $htmlContent += @"
            <tr>
                <td>$($project.Project)</td>
                <td>$($project.Repositories)</td>
                <td>$($project.WorkItems)</td>
                <td>$($project.Builds)</td>
                <td>$($project.PullRequests)</td>
                <td>$($project.Commits)</td>
                <td>$($project.BuildCompletionPercentage)</td>
            </tr>
"@
        }
        $htmlContent += "</table>"
    }

    # Add User License Summary if available
    if ($LoadedData.Users) {
        $LicenseGroups = $LoadedData.Users.data | Group-Object license | Sort-Object Count -Descending
        $htmlContent += @"
        <h2>User License Distribution</h2>
        <table>
            <tr><th>License Type</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($license in $LicenseGroups) {
            $percentage = [math]::Round(($license.Count / $Stats.TotalUsers) * 100, 1)
            $htmlContent += "<tr><td>$($license.Name)</td><td>$($license.Count)</td><td>$percentage%</td></tr>"
        }
        $htmlContent += "</table>"
    }

    # Add Work Item Type Summary if available
    if ($LoadedData.WorkItems) {
        $WorkItemTypes = $LoadedData.WorkItems.data | Group-Object workItemType | Sort-Object Count -Descending
        $htmlContent += @"
        <h2>Work Item Type Distribution</h2>
        <table>
            <tr><th>Work Item Type</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($wiType in $WorkItemTypes) {
            $percentage = [math]::Round(($wiType.Count / $Stats.TotalWorkItems) * 100, 1)
            $htmlContent += "<tr><td>$($wiType.Name)</td><td>$($wiType.Count)</td><td>$percentage%</td></tr>"
        }
        $htmlContent += "</table>"
    }

    $htmlContent += @"
        <div class="footer">
            Generated by Azure DevOps Statistics Local Collector<br>
            Data Path: $DataPath
        </div>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $ReportPath -Encoding UTF8
    Write-Host "HTML report generated: $ReportPath" -ForegroundColor Green
    
    if ($OpenReport) {
        Start-Process $ReportPath
    }
}
elseif ($OutputFormat -eq "JSON") {
    # Generate consolidated JSON report
    $ReportPath = Join-Path $DataPath "Azure-DevOps-Report.json"
    
    $ConsolidatedReport = @{
        metadata = @{
            organization = $Stats.Organization
            reportGenerated = $Stats.ReportGenerated
            dataCollectionDate = $Stats.DataCollectionDate
            reportType = "ConsolidatedStatistics"
        }
        summary = $Stats
        data = $LoadedData
    }
    
    $ConsolidatedReport | ConvertTo-Json -Depth 20 | Out-File -FilePath $ReportPath -Encoding UTF8
    Write-Host "JSON report generated: $ReportPath" -ForegroundColor Green
}

Write-Host "`nReport generation completed!" -ForegroundColor Green
Write-Host "Summary Statistics:" -ForegroundColor Yellow
Write-Host "- Projects: $($Stats.TotalProjects)" -ForegroundColor White
Write-Host "- Users: $($Stats.TotalUsers)" -ForegroundColor White  
Write-Host "- Work Items: $($Stats.TotalWorkItems)" -ForegroundColor White
Write-Host "- Pull Requests: $($Stats.TotalPullRequests)" -ForegroundColor White
Write-Host "- Commits: $($Stats.TotalCommits)" -ForegroundColor White