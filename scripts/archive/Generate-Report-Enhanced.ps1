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
        Write-Error "No data path found. Please run Collect-Statistics-Enhanced.ps1 first or specify -DataPath."
        exit 1
    }
}

Write-Host "🎯 Generating comprehensive report from: $DataPath" -ForegroundColor Cyan

# Load JSON data files
$DataFiles = @{
    Summary = Join-Path $DataPath "Collection-Summary.json"
    Users = Join-Path $DataPath "Users.json"
    GitRepositories = Join-Path $DataPath "GitRepositories.json"
    GitCommits = Join-Path $DataPath "GitCommits.json"
    BuildPipelines = Join-Path $DataPath "BuildPipelines.json"
    Builds = Join-Path $DataPath "Builds.json"
    WorkItems = Join-Path $DataPath "WorkItems.json"
}

$LoadedData = @{}
foreach ($key in $DataFiles.Keys) {
    $filePath = $DataFiles[$key]
    if (Test-Path $filePath) {
        try {
            $jsonContent = Get-Content -Path $filePath -Raw | ConvertFrom-Json
            $LoadedData[$key] = $jsonContent
            Write-Host "✅ Loaded: $key" -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not load $key from $filePath : $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning "⚠️ File not found: $filePath"
    }
}

# Generate report timestamp
$ReportTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$Organization = if ($LoadedData.Summary) { $LoadedData.Summary.Organization } else { $env:ADOS_ORGANIZATION }

# Calculate actual project statistics from collected data
$ProjectStats = @()
if ($LoadedData.GitRepositories -or $LoadedData.WorkItems -or $LoadedData.GitCommits -or $LoadedData.Builds) {
    # Get all unique projects
    $AllProjects = @()
    if ($LoadedData.GitRepositories) {
        $AllProjects += $LoadedData.GitRepositories.data | Select-Object -Property project, projectId -Unique
    }
    if ($LoadedData.WorkItems) {
        $AllProjects += $LoadedData.WorkItems.data | Select-Object -Property project, projectId -Unique
    }
    
    # Remove duplicates
    $UniqueProjects = $AllProjects | Sort-Object project -Unique
    
    foreach ($proj in $UniqueProjects) {
        # Count repositories for this project
        $repoCount = 0
        if ($LoadedData.GitRepositories) {
            $repoCount = ($LoadedData.GitRepositories.data | Where-Object { $_.project -eq $proj.project }).Count
        }
        
        # Count work items for this project  
        $workItemCount = 0
        if ($LoadedData.WorkItems) {
            $workItemCount = ($LoadedData.WorkItems.data | Where-Object { $_.project -eq $proj.project }).Count
        }
        
        # Count commits for this project
        $commitCount = 0
        if ($LoadedData.GitCommits) {
            $commitCount = ($LoadedData.GitCommits.data | Where-Object { $_.project -eq $proj.project }).Count
        }
        
        # Count pipelines for this project
        $pipelineCount = 0
        if ($LoadedData.BuildPipelines) {
            $pipelineCount = ($LoadedData.BuildPipelines.data | Where-Object { $_.project -eq $proj.project }).Count
        }
        
        # Count builds for this project
        $buildCount = 0
        $successfulBuilds = 0
        if ($LoadedData.Builds) {
            $projectBuilds = $LoadedData.Builds.data | Where-Object { $_.project -eq $proj.project }
            $buildCount = $projectBuilds.Count
            $successfulBuilds = ($projectBuilds | Where-Object { $_.result -eq "succeeded" }).Count
        }
        
        $buildSuccessRate = if ($buildCount -gt 0) { [math]::Round(($successfulBuilds / $buildCount) * 100, 1) } else { 0 }
        
        $ProjectStats += [PSCustomObject]@{
            Project = $proj.project
            ProjectId = $proj.projectId
            Repositories = $repoCount
            WorkItems = $workItemCount
            Commits = $commitCount
            Pipelines = $pipelineCount
            Builds = $buildCount
            SuccessfulBuilds = $successfulBuilds
            BuildSuccessRate = "$buildSuccessRate%"
        }
    }
}

# Calculate summary statistics
$Stats = @{
    Organization = $Organization
    ReportGenerated = $ReportTimestamp
    DataCollectionDate = if ($LoadedData.Summary) { $LoadedData.Summary.CollectionTime } else { "Unknown" }
    TotalProjects = $ProjectStats.Count
    TotalRepositories = if ($LoadedData.GitRepositories) { $LoadedData.GitRepositories.data.Count } else { 0 }
    TotalUsers = if ($LoadedData.Users) { $LoadedData.Users.data.Count } else { 0 }
    TotalCommits = if ($LoadedData.GitCommits) { $LoadedData.GitCommits.data.Count } else { 0 }
    TotalPipelines = if ($LoadedData.BuildPipelines) { $LoadedData.BuildPipelines.data.Count } else { 0 }
    TotalBuilds = if ($LoadedData.Builds) { $LoadedData.Builds.data.Count } else { 0 }
    TotalWorkItems = if ($LoadedData.WorkItems) { $LoadedData.WorkItems.data.Count } else { 0 }
}

if ($OutputFormat -eq "HTML") {
    # Generate HTML report
    $ReportPath = Join-Path $DataPath "Azure-DevOps-Comprehensive-Report.html"
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Azure DevOps Comprehensive Report - $Organization</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #323130; border-bottom: 2px solid #edebe9; padding-bottom: 8px; margin-top: 30px; }
        h3 { color: #605e5c; border-bottom: 1px solid #edebe9; padding-bottom: 5px; margin-top: 25px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin: 20px 0; }
        .stat-card { background: linear-gradient(135deg, #0078d4, #106ebe); color: white; padding: 20px; border-radius: 8px; text-align: center; }
        .stat-number { font-size: 2em; font-weight: bold; margin-bottom: 5px; }
        .stat-label { font-size: 0.9em; opacity: 0.9; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 0.9em; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background-color: #0078d4; color: white; font-weight: 600; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #f0f0f0; }
        .info-box { background-color: #f3f2f1; border-left: 4px solid #0078d4; padding: 15px; margin: 15px 0; }
        .success { color: #107c10; font-weight: bold; }
        .warning { color: #ff8c00; font-weight: bold; }
        .error { color: #d13438; font-weight: bold; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #edebe9; color: #605e5c; font-size: 0.9em; text-align: center; }
        .metric-highlight { background-color: #fff3cd; padding: 2px 6px; border-radius: 3px; }
        .table-container { overflow-x: auto; margin: 15px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🚀 Azure DevOps Comprehensive Report</h1>
        
        <div class="info-box">
            <strong>Organization:</strong> $($Stats.Organization)<br>
            <strong>Data Collection Date:</strong> $($Stats.DataCollectionDate)<br>
            <strong>Report Generated:</strong> $($Stats.ReportGenerated)
        </div>

        <h2>📊 Overview Statistics</h2>
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
                <div class="stat-number">$($Stats.TotalRepositories)</div>
                <div class="stat-label">Repositories</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalCommits)</div>
                <div class="stat-label">Commits</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalPipelines)</div>
                <div class="stat-label">Pipelines</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalBuilds)</div>
                <div class="stat-label">Builds</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalWorkItems)</div>
                <div class="stat-label">Work Items</div>
            </div>
        </div>
"@

    # Add Project Statistics table
    if ($ProjectStats.Count -gt 0) {
        $htmlContent += @"
        <h2>📁 Project Statistics</h2>
        <div class="table-container">
        <table>
            <tr>
                <th>Project</th>
                <th>Repositories</th>
                <th>Commits</th>
                <th>Pipelines</th>
                <th>Builds</th>
                <th>Successful Builds</th>
                <th>Build Success Rate</th>
                <th>Work Items</th>
            </tr>
"@
        foreach ($project in $ProjectStats) {
            $successClass = if ([int]($project.BuildSuccessRate -replace '%','') -gt 80) { "success" } 
                           elseif ([int]($project.BuildSuccessRate -replace '%','') -gt 50) { "warning" } 
                           else { "error" }
            
            $htmlContent += @"
            <tr>
                <td><strong>$($project.Project)</strong></td>
                <td>$($project.Repositories)</td>
                <td>$($project.Commits)</td>
                <td>$($project.Pipelines)</td>
                <td>$($project.Builds)</td>
                <td>$($project.SuccessfulBuilds)</td>
                <td><span class="$successClass">$($project.BuildSuccessRate)</span></td>
                <td>$($project.WorkItems)</td>
            </tr>
"@
        }
        $htmlContent += "</table></div>"
    }

    # Add User Statistics
    if ($LoadedData.Users) {
        $LicenseGroups = $LoadedData.Users.data | Group-Object license | Sort-Object Count -Descending
        $StatusGroups = $LoadedData.Users.data | Group-Object status | Sort-Object Count -Descending
        $htmlContent += @"
        <h2>👥 User Statistics</h2>
        
        <h3>License Distribution</h3>
        <div class="table-container">
        <table>
            <tr><th>License Type</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($license in $LicenseGroups) {
            $percentage = [math]::Round(($license.Count / $Stats.TotalUsers) * 100, 1)
            $htmlContent += "<tr><td>$($license.Name)</td><td>$($license.Count)</td><td><span class='metric-highlight'>$percentage%</span></td></tr>"
        }
        $htmlContent += "</table></div>"

        $htmlContent += @"
        <h3>User Status Distribution</h3>
        <div class="table-container">
        <table>
            <tr><th>Status</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($status in $StatusGroups) {
            $percentage = [math]::Round(($status.Count / $Stats.TotalUsers) * 100, 1)
            $statusClass = if ($status.Name -eq "active") { "success" } else { "warning" }
            $htmlContent += "<tr><td><span class='$statusClass'>$($status.Name)</span></td><td>$($status.Count)</td><td><span class='metric-highlight'>$percentage%</span></td></tr>"
        }
        $htmlContent += "</table></div>"
    }

    # Add Pipeline Details
    if ($LoadedData.BuildPipelines) {
        $htmlContent += @"
        <h2>🔧 Build Pipelines</h2>
        <div class="table-container">
        <table>
            <tr><th>Project</th><th>Pipeline Name</th><th>Type</th><th>Repository</th><th>Quality</th><th>Status</th><th>Created By</th></tr>
"@
        foreach ($pipeline in ($LoadedData.BuildPipelines.data | Sort-Object project, name)) {
            $qualityClass = if ($pipeline.quality -eq "definition") { "success" } else { "warning" }
            $statusClass = if ($pipeline.queueStatus -eq "enabled") { "success" } else { "error" }
            $htmlContent += @"
            <tr>
                <td>$($pipeline.project)</td>
                <td><strong>$($pipeline.name)</strong></td>
                <td>$($pipeline.type)</td>
                <td>$($pipeline.repository)</td>
                <td><span class="$qualityClass">$($pipeline.quality)</span></td>
                <td><span class="$statusClass">$($pipeline.queueStatus)</span></td>
                <td>$($pipeline.authoredBy)</td>
            </tr>
"@
        }
        $htmlContent += "</table></div>"
    }

    # Add Build Results Summary
    if ($LoadedData.Builds) {
        $BuildResults = $LoadedData.Builds.data | Group-Object result | Sort-Object Count -Descending
        $htmlContent += @"
        <h2>🏗️ Build Results Summary</h2>
        <div class="table-container">
        <table>
            <tr><th>Build Result</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($result in $BuildResults) {
            $percentage = [math]::Round(($result.Count / $Stats.TotalBuilds) * 100, 1)
            $resultClass = if ($result.Name -eq "succeeded") { "success" } 
                          elseif ($result.Name -eq "failed") { "error" } 
                          else { "warning" }
            $htmlContent += "<tr><td><span class='$resultClass'>$($result.Name)</span></td><td>$($result.Count)</td><td><span class='metric-highlight'>$percentage%</span></td></tr>"
        }
        $htmlContent += "</table></div>"
    }

    # Add Work Item Type Summary
    if ($LoadedData.WorkItems) {
        $WorkItemTypes = $LoadedData.WorkItems.data | Group-Object workItemType | Sort-Object Count -Descending
        $WorkItemStates = $LoadedData.WorkItems.data | Group-Object state | Sort-Object Count -Descending
        
        $htmlContent += @"
        <h2>📋 Work Item Analysis</h2>
        
        <h3>Work Item Types</h3>
        <div class="table-container">
        <table>
            <tr><th>Work Item Type</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($wiType in $WorkItemTypes) {
            $percentage = [math]::Round(($wiType.Count / $Stats.TotalWorkItems) * 100, 1)
            $htmlContent += "<tr><td>$($wiType.Name)</td><td>$($wiType.Count)</td><td><span class='metric-highlight'>$percentage%</span></td></tr>"
        }
        $htmlContent += "</table></div>"

        $htmlContent += @"
        <h3>Work Item States</h3>
        <div class="table-container">
        <table>
            <tr><th>State</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($state in $WorkItemStates) {
            $percentage = [math]::Round(($state.Count / $Stats.TotalWorkItems) * 100, 1)
            $stateClass = if ($state.Name -eq "Done" -or $state.Name -eq "Closed") { "success" } 
                         elseif ($state.Name -eq "Active" -or $state.Name -eq "In Progress") { "warning" } 
                         else { "" }
            $htmlContent += "<tr><td><span class='$stateClass'>$($state.Name)</span></td><td>$($state.Count)</td><td><span class='metric-highlight'>$percentage%</span></td></tr>"
        }
        $htmlContent += "</table></div>"
    }

    # Add Repository Details
    if ($LoadedData.GitRepositories) {
        $htmlContent += @"
        <h2>📚 Repository Details</h2>
        <div class="table-container">
        <table>
            <tr><th>Project</th><th>Repository Name</th><th>Default Branch</th><th>Size (bytes)</th><th>Commits</th></tr>
"@
        foreach ($repo in ($LoadedData.GitRepositories.data | Sort-Object project, name)) {
            $size = if ($repo.size -eq 0) { "Empty" } else { "{0:N0}" -f $repo.size }
            $branch = if ($repo.defaultBranch) { $repo.defaultBranch -replace "refs/heads/", "" } else { "N/A" }
            
            # Count commits for this repository
            $repoCommits = 0
            if ($LoadedData.GitCommits) {
                $repoCommits = ($LoadedData.GitCommits.data | Where-Object { $_.repositoryId -eq $repo.repositoryId }).Count
            }
            
            $htmlContent += "<tr><td>$($repo.project)</td><td><strong>$($repo.name)</strong></td><td>$branch</td><td>$size</td><td>$repoCommits</td></tr>"
        }
        $htmlContent += "</table></div>"
    }

    $htmlContent += @"
        <div class="footer">
            <span class="success">✅ Comprehensive Data Collection Successful</span><br>
            Generated by Azure DevOps Statistics Enhanced Collector<br>
            Data Path: $DataPath<br>
            Total Data Points Collected: $($Stats.TotalUsers + $Stats.TotalRepositories + $Stats.TotalCommits + $Stats.TotalPipelines + $Stats.TotalBuilds + $Stats.TotalWorkItems)
        </div>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $ReportPath -Encoding UTF8
    Write-Host "🎯 Comprehensive HTML report generated: $ReportPath" -ForegroundColor Green
    
    if ($OpenReport) {
        Start-Process $ReportPath
    }
}

Write-Host "`n🎉 Report generation completed!" -ForegroundColor Green
Write-Host "📊 Summary Statistics:" -ForegroundColor Yellow
Write-Host "  Projects: $($Stats.TotalProjects)" -ForegroundColor White
Write-Host "  Users: $($Stats.TotalUsers)" -ForegroundColor White  
Write-Host "  Repositories: $($Stats.TotalRepositories)" -ForegroundColor White
Write-Host "  Commits: $($Stats.TotalCommits)" -ForegroundColor White
Write-Host "  Pipelines: $($Stats.TotalPipelines)" -ForegroundColor White
Write-Host "  Builds: $($Stats.TotalBuilds)" -ForegroundColor White
Write-Host "  Work Items: $($Stats.TotalWorkItems)" -ForegroundColor White