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
    Users = Join-Path $DataPath "Users.json"
    GitRepositories = Join-Path $DataPath "GitRepositories.json"
    GitCommits = Join-Path $DataPath "GitCommits.json"
    BuildPipelines = Join-Path $DataPath "BuildPipelines.json"
    Builds = Join-Path $DataPath "Builds.json"
    WorkItems = Join-Path $DataPath "WorkItems.json"
    GitPullRequests = Join-Path $DataPath "GitPullRequests.json"
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

# Calculate actual project statistics from collected data
$ProjectStats = @()
if ($LoadedData.GitRepositories -or $LoadedData.WorkItems) {
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
        
        # Count pull requests for this project
        $prCount = 0
        if ($LoadedData.GitPullRequests) {
            $prCount = ($LoadedData.GitPullRequests.data | Where-Object { $_.project -eq $proj.project }).Count
        }
        
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
            PullRequests = $prCount
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
    TotalWorkItems = if ($LoadedData.WorkItems) { $LoadedData.WorkItems.data.Count } else { 0 }
    TotalCommits = if ($LoadedData.GitCommits) { $LoadedData.GitCommits.data.Count } else { 0 }
    TotalPipelines = if ($LoadedData.BuildPipelines) { $LoadedData.BuildPipelines.data.Count } else { 0 }
    TotalBuilds = if ($LoadedData.Builds) { $LoadedData.Builds.data.Count } else { 0 }
    TotalPullRequests = if ($LoadedData.GitPullRequests) { $LoadedData.GitPullRequests.data.Count } else { 0 }
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
        .success { color: #107c10; font-weight: bold; }
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
                <div class="stat-number">$($Stats.TotalRepositories)</div>
                <div class="stat-label">Repositories</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$($Stats.TotalWorkItems)</div>
                <div class="stat-label">Work Items</div>
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
        </div>
"@

    # Add Project Statistics table
    if ($ProjectStats.Count -gt 0) {
        $htmlContent += @"
        <h2>Project Statistics</h2>
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
                <th>Pull Requests</th>
            </tr>
"@
        foreach ($project in $ProjectStats) {
            $successClass = if ([int]($project.BuildSuccessRate -replace '%','') -gt 80) { 'style="color: #107c10; font-weight: bold;"' } 
                           elseif ([int]($project.BuildSuccessRate -replace '%','') -gt 50) { 'style="color: #ff8c00; font-weight: bold;"' } 
                           else { 'style="color: #d13438; font-weight: bold;"' }
            
            $htmlContent += @"
            <tr>
                <td><strong>$($project.Project)</strong></td>
                <td>$($project.Repositories)</td>
                <td>$($project.Commits)</td>
                <td>$($project.Pipelines)</td>
                <td>$($project.Builds)</td>
                <td>$($project.SuccessfulBuilds)</td>
                <td $successClass>$($project.BuildSuccessRate)</td>
                <td>$($project.WorkItems)</td>
                <td>$($project.PullRequests)</td>
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
        
        # Add User Status Distribution
        $StatusGroups = $LoadedData.Users.data | Group-Object status | Sort-Object Count -Descending
        $htmlContent += @"
        <h2>User Status Distribution</h2>
        <table>
            <tr><th>Status</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($status in $StatusGroups) {
            $percentage = [math]::Round(($status.Count / $Stats.TotalUsers) * 100, 1)
            $statusColor = if ($status.Name -eq "active") { 'style="color: #107c10; font-weight: bold;"' } else { 'style="color: #ff8c00; font-weight: bold;"' }
            $htmlContent += "<tr><td $statusColor>$($status.Name)</td><td>$($status.Count)</td><td>$percentage%</td></tr>"
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

    # Add Top Contributors section
    if ($LoadedData.GitCommits) {
        # Group by authorEmail to avoid duplicate contributors with same display name
        $CommitsByAuthor = $LoadedData.GitCommits.data | Group-Object authorEmail | Sort-Object Count -Descending | Select-Object -First 15
        $htmlContent += @"
        <h2>🏆 Top Contributors</h2>
        <table>
            <tr><th>Rank</th><th>Author</th><th>Email</th><th>Commits</th><th>Percentage</th><th>Latest Commit</th></tr>
"@
        $rank = 1
        foreach ($author in $CommitsByAuthor) {
            $percentage = [math]::Round(($author.Count / $Stats.TotalCommits) * 100, 1)
            # Determine a display name for this email (pick most frequent display name for that email)
            $displayName = ($author.Group | Group-Object author | Sort-Object Count -Descending | Select-Object -First 1).Name
            $latestCommit = ($author.Group | Sort-Object authorDate -Descending | Select-Object -First 1).authorDate
            $latestCommitFormatted = if ($latestCommit) { ([datetime]$latestCommit).ToString("yyyy-MM-dd") } else { "N/A" }
            
            $rankStyle = ""
            if ($rank -eq 1) { $rankStyle = 'style="color: #FFD700; font-weight: bold;"' }  # Gold
            elseif ($rank -eq 2) { $rankStyle = 'style="color: #C0C0C0; font-weight: bold;"' }  # Silver
            elseif ($rank -eq 3) { $rankStyle = 'style="color: #CD7F32; font-weight: bold;"' }  # Bronze
            
            $htmlContent += "<tr><td $rankStyle>$rank</td><td><strong>$displayName</strong></td><td>$($author.Name)</td><td>$($author.Count)</td><td>$percentage%</td><td>$latestCommitFormatted</td></tr>"
            $rank++
        }
        $htmlContent += "</table>"
        
        # Add Contributor Activity by Project
        $htmlContent += @"
        <h2>📊 Contributor Activity by Project</h2>
        <table>
            <tr><th>Author</th><th>Email</th><th>Project</th><th>Commits</th><th>Latest Commit</th><th>Repositories</th></tr>
"@
        
        # Group by authorEmail and project to avoid name-based duplication
        $contributorsByProject = $LoadedData.GitCommits.data | Group-Object authorEmail, project | Sort-Object { $_.Group.Count } -Descending | Select-Object -First 20
        foreach ($contributor in $contributorsByProject) {
            $email = $contributor.Name.Split(', ')[0]
            $projectName = $contributor.Name.Split(', ')[1]
            $commitCount = $contributor.Count
            # Determine a display name for this email within this project
            $authorName = ($contributor.Group | Group-Object author | Sort-Object Count -Descending | Select-Object -First 1).Name
            $latestCommit = ($contributor.Group | Sort-Object authorDate -Descending | Select-Object -First 1).authorDate
            $latestCommitFormatted = if ($latestCommit) { ([datetime]$latestCommit).ToString("yyyy-MM-dd") } else { "N/A" }
            $repositories = ($contributor.Group | Select-Object repository -Unique).Count
            
            $htmlContent += "<tr><td>$authorName</td><td>$email</td><td>$projectName</td><td>$commitCount</td><td>$latestCommitFormatted</td><td>$repositories</td></tr>"
        }
        $htmlContent += "</table>"
    }

    # Add Project Breakdown if available
    if ($LoadedData.GitRepositories) {
        $htmlContent += @"
        <h2>Repository Details by Project</h2>
        <table>
            <tr><th>Project</th><th>Repository Name</th><th>Default Branch</th><th>Size (bytes)</th><th>Commits</th><th>Created Date</th><th>Last Commit</th></tr>
"@
        foreach ($repo in ($LoadedData.GitRepositories.data | Sort-Object project, name)) {
            $size = if ($repo.size -eq 0) { "Empty" } else { "{0:N0}" -f $repo.size }
            $branch = if ($repo.defaultBranch) { $repo.defaultBranch -replace "refs/heads/", "" } else { "N/A" }
            
            # Count commits for this repository
            $repoCommits = 0
            $lastCommitDate = "N/A"
            if ($LoadedData.GitCommits) {
                $commits = $LoadedData.GitCommits.data | Where-Object { $_.repositoryId -eq $repo.repositoryId }
                $repoCommits = $commits.Count
                if ($commits.Count -gt 0) {
                    $latestCommit = $commits | Sort-Object authorDate -Descending | Select-Object -First 1
                    if ($latestCommit.authorDate) {
                        $lastCommitDate = ([datetime]$latestCommit.authorDate).ToString("yyyy-MM-dd HH:mm")
                    }
                }
            }
            
            # Get repository creation date from the earliest commit or set to N/A
            $createdDate = "N/A"
            if ($LoadedData.GitCommits) {
                $commits = $LoadedData.GitCommits.data | Where-Object { $_.repositoryId -eq $repo.repositoryId }
                if ($commits.Count -gt 0) {
                    $earliestCommit = $commits | Sort-Object authorDate | Select-Object -First 1
                    if ($earliestCommit.authorDate) {
                        $createdDate = ([datetime]$earliestCommit.authorDate).ToString("yyyy-MM-dd")
                    }
                }
            }
            
            $htmlContent += "<tr><td>$($repo.project)</td><td><strong>$($repo.name)</strong></td><td>$branch</td><td>$size</td><td>$repoCommits</td><td>$createdDate</td><td>$lastCommitDate</td></tr>"
        }
        $htmlContent += "</table>"
    }

    $htmlContent += @"
        <div class="footer">
            <span class="success">✓ Data Collection Successful</span><br>
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

Write-Host "`nReport generation completed!" -ForegroundColor Green
Write-Host "Summary Statistics:" -ForegroundColor Yellow
Write-Host "- Projects: $($Stats.TotalProjects)" -ForegroundColor White
Write-Host "- Users: $($Stats.TotalUsers)" -ForegroundColor White  
Write-Host "- Repositories: $($Stats.TotalRepositories)" -ForegroundColor White
Write-Host "- Commits: $($Stats.TotalCommits)" -ForegroundColor White
Write-Host "- Pipelines: $($Stats.TotalPipelines)" -ForegroundColor White
Write-Host "- Builds: $($Stats.TotalBuilds)" -ForegroundColor White
Write-Host "- Work Items: $($Stats.TotalWorkItems)" -ForegroundColor White
Write-Host "- Pull Requests: $($Stats.TotalPullRequests)" -ForegroundColor White