[CmdletBinding()]
param 
(
    [string]$DataPath = "",
    [string]$OutputFormat = "HTML",
    [switch]$OpenReport
)

# =============================================================================
# Load centralized scoring configuration
# =============================================================================
$configPath = Join-Path $PSScriptRoot "Config-ScoringWeights.ps1"
if (Test-Path $configPath) {
    . $configPath
    $ScoringWeights = $Global:ScoringWeights
    Write-Host "Loaded scoring configuration from Config-ScoringWeights.ps1" -ForegroundColor Green
} else {
    # Fallback if config file not found
    Write-Warning "Config-ScoringWeights.ps1 not found, using default weights"
    $ScoringWeights = @{
        PRsMerged       = 5
        PRsCreated      = 3
        CodeReviews     = 2
        WorkItems       = 2
        Commits         = 1
    }
}
# =============================================================================

if (-not $DataPath) {
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
        Write-Error "No data path found. Please run Collect-Statistics-Fixed.ps1 first or specify -DataPath."
        exit 1
    }
}

Write-Host "Generating enhanced report from: $DataPath" -ForegroundColor Cyan

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
}

$ReportTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$Organization = if ($LoadedData.Summary) { $LoadedData.Summary.Organization } else { $env:ADOS_ORGANIZATION }

# Calculate project statistics
$ProjectStats = @()
$AllProjects = @()

if ($LoadedData.GitRepositories) {
    $AllProjects += $LoadedData.GitRepositories.data | Select-Object -Property project, projectId -Unique
}
if ($LoadedData.WorkItems) {
    $AllProjects += $LoadedData.WorkItems.data | Select-Object -Property project, projectId -Unique
}
if ($LoadedData.BuildPipelines) {
    $AllProjects += $LoadedData.BuildPipelines.data | Select-Object -Property project, projectId -Unique
}

$UniqueProjects = $AllProjects | Sort-Object project -Unique

foreach ($proj in $UniqueProjects) {
    $repoCount = if ($LoadedData.GitRepositories) { ($LoadedData.GitRepositories.data | Where-Object { $_.project -eq $proj.project }).Count } else { 0 }
    $workItemCount = if ($LoadedData.WorkItems) { ($LoadedData.WorkItems.data | Where-Object { $_.project -eq $proj.project }).Count } else { 0 }
    $commitCount = if ($LoadedData.GitCommits) { ($LoadedData.GitCommits.data | Where-Object { $_.project -eq $proj.project }).Count } else { 0 }
    $pipelineCount = if ($LoadedData.BuildPipelines) { ($LoadedData.BuildPipelines.data | Where-Object { $_.project -eq $proj.project }).Count } else { 0 }
    
    $buildCount = 0
    $successfulBuilds = 0
    $failedBuilds = 0
    if ($LoadedData.Builds) {
        $projectBuilds = $LoadedData.Builds.data | Where-Object { $_.project -eq $proj.project }
        $buildCount = $projectBuilds.Count
        $successfulBuilds = ($projectBuilds | Where-Object { $_.result -eq "succeeded" }).Count
        $failedBuilds = ($projectBuilds | Where-Object { $_.result -eq "failed" }).Count
    }
    
    $buildSuccessRate = if ($buildCount -gt 0) { [math]::Round(($successfulBuilds / $buildCount) * 100, 1) } else { 0 }
    $prCount = if ($LoadedData.GitPullRequests) { ($LoadedData.GitPullRequests.data | Where-Object { $_.project -eq $proj.project }).Count } else { 0 }
    
    $ProjectStats += [PSCustomObject]@{
        Project = $proj.project
        ProjectId = $proj.projectId
        Repositories = $repoCount
        WorkItems = $workItemCount
        Commits = $commitCount
        Pipelines = $pipelineCount
        Builds = $buildCount
        SuccessfulBuilds = $successfulBuilds
        FailedBuilds = $failedBuilds
        BuildSuccessRate = $buildSuccessRate
        PullRequests = $prCount
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

# Calculate additional metrics
$ActiveUsers = if ($LoadedData.Users) { ($LoadedData.Users.data | Where-Object { $_.status -eq "active" }).Count } else { 0 }
$InactiveUsers = $Stats.TotalUsers - $ActiveUsers

$OverallBuildSuccessRate = 0
if ($LoadedData.Builds -and $LoadedData.Builds.data.Count -gt 0) {
    $successBuilds = ($LoadedData.Builds.data | Where-Object { $_.result -eq "succeeded" }).Count
    $OverallBuildSuccessRate = [math]::Round(($successBuilds / $LoadedData.Builds.data.Count) * 100, 1)
}

# Calculate commit activity by day of week
$CommitsByDayOfWeek = @{}
if ($LoadedData.GitCommits) {
    foreach ($commit in $LoadedData.GitCommits.data) {
        if ($commit.authorDate) {
            $dayOfWeek = ([datetime]$commit.authorDate).DayOfWeek.ToString()
            if (-not $CommitsByDayOfWeek.ContainsKey($dayOfWeek)) {
                $CommitsByDayOfWeek[$dayOfWeek] = 0
            }
            $CommitsByDayOfWeek[$dayOfWeek]++
        }
    }
}

# Calculate recent activity (last 30 days)
$thirtyDaysAgo = (Get-Date).AddDays(-30)
$RecentCommits = 0
$RecentBuilds = 0
if ($LoadedData.GitCommits) {
    $RecentCommits = ($LoadedData.GitCommits.data | Where-Object { $_.authorDate -and ([datetime]$_.authorDate) -gt $thirtyDaysAgo }).Count
}
if ($LoadedData.Builds) {
    $RecentBuilds = ($LoadedData.Builds.data | Where-Object { $_.finishTime -and ([datetime]$_.finishTime) -gt $thirtyDaysAgo }).Count
}

if ($OutputFormat -eq "HTML") {
    $ReportPath = Join-Path $DataPath "Azure-DevOps-Report.html"
    
    # Prepare chart data
    $projectLabels = ($ProjectStats | ForEach-Object { "'$($_.Project)'" }) -join ","
    $projectCommits = ($ProjectStats | ForEach-Object { $_.Commits }) -join ","
    $projectRepos = ($ProjectStats | ForEach-Object { $_.Repositories }) -join ","
    $projectBuilds = ($ProjectStats | ForEach-Object { $_.Builds }) -join ","
    $projectWorkItems = ($ProjectStats | ForEach-Object { $_.WorkItems }) -join ","
    
    # Day of week data
    $daysOrder = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')
    $dayLabels = ($daysOrder | ForEach-Object { "'$_'" }) -join ","
    $dayValues = ($daysOrder | ForEach-Object { if ($CommitsByDayOfWeek[$_]) { $CommitsByDayOfWeek[$_] } else { 0 } }) -join ","
    
    # License distribution data
    $licenseLabels = ""
    $licenseValues = ""
    $licenseColors = ""
    if ($LoadedData.Users) {
        $LicenseGroups = $LoadedData.Users.data | Group-Object license | Sort-Object Count -Descending
        $licenseLabels = ($LicenseGroups | ForEach-Object { "'$($_.Name)'" }) -join ","
        $licenseValues = ($LicenseGroups | ForEach-Object { $_.Count }) -join ","
        $colors = @("'#0078d4'", "'#00bcf2'", "'#8764b8'", "'#e81123'", "'#ff8c00'", "'#107c10'", "'#ffb900'", "'#00188f'")
        $licenseColors = ($LicenseGroups | ForEach-Object -Begin { $i = 0 } -Process { $colors[$i % $colors.Count]; $i++ }) -join ","
    }
    
    # Build status data
    $buildSucceeded = if ($LoadedData.Builds) { ($LoadedData.Builds.data | Where-Object { $_.result -eq "succeeded" }).Count } else { 0 }
    $buildFailed = if ($LoadedData.Builds) { ($LoadedData.Builds.data | Where-Object { $_.result -eq "failed" }).Count } else { 0 }
    $buildCanceled = if ($LoadedData.Builds) { ($LoadedData.Builds.data | Where-Object { $_.result -eq "canceled" }).Count } else { 0 }
    $buildOther = $Stats.TotalBuilds - $buildSucceeded - $buildFailed - $buildCanceled
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Azure DevOps Statistics - $Organization</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --azure-blue: #0078d4;
            --azure-dark: #106ebe;
            --success: #107c10;
            --warning: #ff8c00;
            --danger: #d13438;
            --gray-10: #faf9f8;
            --gray-20: #f3f2f1;
            --gray-30: #edebe9;
            --gray-90: #323130;
            --gray-130: #605e5c;
        }
        
        * { box-sizing: border-box; margin: 0; padding: 0; }
        
        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .dashboard {
            max-width: 1400px;
            margin: 0 auto;
        }
        
        .header {
            background: white;
            border-radius: 16px;
            padding: 30px;
            margin-bottom: 20px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        }
        
        .header h1 {
            color: var(--azure-blue);
            font-size: 2.5em;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 15px;
        }
        
        .header h1 i {
            color: var(--azure-blue);
        }
        
        .header-meta {
            display: flex;
            gap: 30px;
            color: var(--gray-130);
            font-size: 0.95em;
            flex-wrap: wrap;
        }
        
        .header-meta span {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .header-meta i {
            color: var(--azure-blue);
        }
        
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }
        
        .metric-card {
            background: white;
            border-radius: 12px;
            padding: 20px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.08);
            transition: transform 0.2s, box-shadow 0.2s;
        }
        
        .metric-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }
        
        .metric-icon {
            font-size: 2em;
            margin-bottom: 10px;
            color: var(--azure-blue);
        }
        
        .metric-value {
            font-size: 2.5em;
            font-weight: 700;
            color: var(--azure-blue);
            line-height: 1;
        }
        
        .metric-label {
            color: var(--gray-130);
            font-size: 0.9em;
            margin-top: 5px;
        }
        
        .metric-trend {
            font-size: 0.85em;
            margin-top: 8px;
            padding: 4px 8px;
            border-radius: 12px;
            display: inline-block;
        }
        
        .trend-up { background: #dff6dd; color: var(--success); }
        .trend-neutral { background: var(--gray-20); color: var(--gray-130); }
        
        .charts-row {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }
        
        .chart-card {
            background: white;
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        }
        
        .chart-card h3 {
            color: var(--gray-90);
            margin-bottom: 20px;
            font-size: 1.1em;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .chart-card h3 i {
            color: var(--azure-blue);
        }
        
        .chart-container {
            position: relative;
            height: 300px;
        }
        
        .section {
            background: white;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 20px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        }
        
        .section h2 {
            color: var(--gray-90);
            font-size: 1.3em;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid var(--gray-30);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .section h2 i {
            color: var(--azure-blue);
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.95em;
        }
        
        th {
            background: var(--azure-blue);
            color: white;
            padding: 14px 12px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
        }
        
        td {
            padding: 12px;
            border-bottom: 1px solid var(--gray-30);
        }
        
        tr:hover {
            background: var(--gray-10);
        }
        
        .badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.85em;
            font-weight: 600;
        }
        
        .badge-success { background: #dff6dd; color: var(--success); }
        .badge-warning { background: #fff4ce; color: #8a6914; }
        .badge-danger { background: #fde7e9; color: var(--danger); }
        .badge-info { background: #cce4f7; color: #004578; }
        
        .progress-bar {
            background: var(--gray-20);
            border-radius: 10px;
            height: 20px;
            overflow: hidden;
            position: relative;
        }
        
        .progress-fill {
            height: 100%;
            border-radius: 10px;
            transition: width 0.5s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 0.75em;
            font-weight: 600;
        }
        
        .leaderboard {
            display: grid;
            gap: 10px;
        }
        
        .leader-item {
            display: flex;
            align-items: center;
            padding: 15px;
            background: var(--gray-10);
            border-radius: 10px;
            gap: 15px;
        }
        
        .leader-rank {
            font-size: 1.5em;
            width: 50px;
            text-align: center;
        }
        
        .leader-rank.gold { color: #FFD700; }
        .leader-rank.silver { color: #C0C0C0; }
        .leader-rank.bronze { color: #CD7F32; }
        
        .leader-info {
            flex: 1;
        }
        
        .leader-name {
            font-weight: 600;
            color: var(--gray-90);
        }
        
        .leader-stats {
            font-size: 0.85em;
            color: var(--gray-130);
        }
        
        .leader-commits {
            font-size: 1.3em;
            font-weight: 700;
            color: var(--azure-blue);
        }
        
        .footer {
            text-align: center;
            padding: 20px;
            color: white;
            font-size: 0.9em;
            opacity: 0.9;
        }
        
        .two-col {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 20px;
        }
        
        @media (max-width: 768px) {
            .header h1 { font-size: 1.8em; }
            .metrics-grid { grid-template-columns: repeat(2, 1fr); }
            .charts-row { grid-template-columns: 1fr; }
            .chart-container { height: 250px; }
        }
    </style>
</head>
<body>
    <div class="dashboard">
        <div class="header">
            <h1><i class="fas fa-chart-bar"></i> Azure DevOps Statistics Report</h1>
            <div class="header-meta">
                <span><i class="fas fa-building"></i> <strong>$($Stats.Organization)</strong></span>
                <span><i class="fas fa-calendar-alt"></i> Data Collected: <strong>$($Stats.DataCollectionDate)</strong></span>
                <span><i class="fas fa-clock"></i> Report Generated: <strong>$($Stats.ReportGenerated)</strong></span>
            </div>
        </div>

        <div class="metrics-grid">
            <div class="metric-card">
                <div class="metric-icon"><i class="fas fa-folder-open"></i></div>
                <div class="metric-value">$($Stats.TotalProjects)</div>
                <div class="metric-label">Projects</div>
            </div>
            <div class="metric-card">
                <div class="metric-icon"><i class="fas fa-users"></i></div>
                <div class="metric-value">$($Stats.TotalUsers)</div>
                <div class="metric-label">Users</div>
                <div class="metric-trend trend-neutral">$ActiveUsers active</div>
            </div>
            <div class="metric-card">
                <div class="metric-icon"><i class="fas fa-code-branch"></i></div>
                <div class="metric-value">$($Stats.TotalRepositories)</div>
                <div class="metric-label">Repositories</div>
            </div>
            <div class="metric-card">
                <div class="metric-icon"><i class="fas fa-code-commit"></i></div>
                <div class="metric-value">$($Stats.TotalCommits)</div>
                <div class="metric-label">Commits</div>
                <div class="metric-trend trend-up">$RecentCommits last 30d</div>
            </div>
            <div class="metric-card">
                <div class="metric-icon"><i class="fas fa-cogs"></i></div>
                <div class="metric-value">$($Stats.TotalPipelines)</div>
                <div class="metric-label">Pipelines</div>
            </div>
            <div class="metric-card">
                <div class="metric-icon"><i class="fas fa-hammer"></i></div>
                <div class="metric-value">$($Stats.TotalBuilds)</div>
                <div class="metric-label">Builds</div>
                <div class="metric-trend trend-up">$RecentBuilds last 30d</div>
            </div>
            <div class="metric-card">
                <div class="metric-icon"><i class="fas fa-tasks"></i></div>
                <div class="metric-value">$($Stats.TotalWorkItems)</div>
                <div class="metric-label">Work Items</div>
            </div>
            <div class="metric-card">
                <div class="metric-icon"><i class="fas fa-code-pull-request"></i></div>
                <div class="metric-value">$($Stats.TotalPullRequests)</div>
                <div class="metric-label">Pull Requests</div>
            </div>
        </div>

        <div class="charts-row">
            <div class="chart-card">
                <h3><i class="fas fa-chart-bar"></i> Activity by Project</h3>
                <div class="chart-container">
                    <canvas id="projectChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3><i class="fas fa-circle-check"></i> Build Status Distribution</h3>
                <div class="chart-container">
                    <canvas id="buildChart"></canvas>
                </div>
            </div>
        </div>

        <div class="charts-row">
            <div class="chart-card">
                <h3><i class="fas fa-calendar-week"></i> Commits by Day of Week</h3>
                <div class="chart-container">
                    <canvas id="dayChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3><i class="fas fa-id-card"></i> License Distribution</h3>
                <div class="chart-container">
                    <canvas id="licenseChart"></canvas>
                </div>
            </div>
        </div>
"@

    # Team Activity Section - Multiple meaningful metrics
    # Build comprehensive activity data per person
    $teamActivity = @{}
    
    # Count commits per person
    if ($LoadedData.GitCommits) {
        foreach ($commit in $LoadedData.GitCommits.data) {
            $name = $commit.author
            if (-not $teamActivity.ContainsKey($name)) {
                $teamActivity[$name] = @{ Name = $name; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0; WorkItemsCreated = 0; LastActivity = $null }
            }
            $teamActivity[$name].Commits++
            $commitDate = if ($commit.authorDate) { [datetime]$commit.authorDate } else { $null }
            if ($commitDate -and (-not $teamActivity[$name].LastActivity -or $commitDate -gt $teamActivity[$name].LastActivity)) {
                $teamActivity[$name].LastActivity = $commitDate
            }
        }
    }
    
    # Count PRs created and merged per person
    if ($LoadedData.GitPullRequests) {
        foreach ($pr in $LoadedData.GitPullRequests.data) {
            $creator = $pr.createdBy
            if ($creator) {
                if (-not $teamActivity.ContainsKey($creator)) {
                    $teamActivity[$creator] = @{ Name = $creator; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0; WorkItemsCreated = 0; LastActivity = $null }
                }
                $teamActivity[$creator].PRsCreated++
                if ($pr.status -eq "completed") {
                    $teamActivity[$creator].PRsMerged++
                }
                $prDate = if ($pr.creationDate) { [datetime]$pr.creationDate } else { $null }
                if ($prDate -and (-not $teamActivity[$creator].LastActivity -or $prDate -gt $teamActivity[$creator].LastActivity)) {
                    $teamActivity[$creator].LastActivity = $prDate
                }
            }
            # Count reviewers
            if ($pr.reviewers) {
                $reviewerList = $pr.reviewers -split "; "
                foreach ($reviewer in $reviewerList) {
                    if ($reviewer -and $reviewer.Trim()) {
                        $reviewerName = $reviewer.Trim()
                        if (-not $teamActivity.ContainsKey($reviewerName)) {
                            $teamActivity[$reviewerName] = @{ Name = $reviewerName; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0; WorkItemsCreated = 0; LastActivity = $null }
                        }
                        $teamActivity[$reviewerName].PRsReviewed++
                    }
                }
            }
        }
    }
    
    # Count work items assigned and created per person
    if ($LoadedData.WorkItems) {
        foreach ($wi in $LoadedData.WorkItems.data) {
            if ($wi.assignedTo) {
                $assignee = $wi.assignedTo
                if (-not $teamActivity.ContainsKey($assignee)) {
                    $teamActivity[$assignee] = @{ Name = $assignee; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0; WorkItemsCreated = 0; LastActivity = $null }
                }
                $teamActivity[$assignee].WorkItemsAssigned++
            }
            if ($wi.createdBy) {
                $creator = $wi.createdBy
                if (-not $teamActivity.ContainsKey($creator)) {
                    $teamActivity[$creator] = @{ Name = $creator; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0; WorkItemsCreated = 0; LastActivity = $null }
                }
                $teamActivity[$creator].WorkItemsCreated++
            }
        }
    }
    
    # Calculate activity score using configurable weights
    foreach ($person in $teamActivity.Keys) {
        $activity = $teamActivity[$person]
        $activity.Score = ($activity.PRsMerged * $ScoringWeights.PRsMerged) + 
                          ($activity.PRsCreated * $ScoringWeights.PRsCreated) + 
                          ($activity.PRsReviewed * $ScoringWeights.CodeReviews) + 
                          ($activity.WorkItemsAssigned * $ScoringWeights.WorkItems) + 
                          ($activity.Commits * $ScoringWeights.Commits)
    }
    
    # Sort by activity score
    $sortedActivity = $teamActivity.Values | Sort-Object { $_.Score } -Descending | Select-Object -First 15
    
    if ($sortedActivity.Count -gt 0) {
        $htmlContent += @"
        <div class="section">
            <h2><i class="fas fa-users-cog"></i> Team Activity</h2>
            <p style="color: var(--gray-130); margin-bottom: 20px; font-size: 0.9em;">
                Activity score based on: PRs Merged ($($ScoringWeights.PRsMerged)pts) + PRs Created ($($ScoringWeights.PRsCreated)pts) + Code Reviews ($($ScoringWeights.CodeReviews)pts) + Work Items ($($ScoringWeights.WorkItems)pts) + Commits ($($ScoringWeights.Commits)pt)
            </p>
            <table>
                <tr>
                    <th style="width: 40px;">Rank</th>
                    <th>Team Member</th>
                    <th style="text-align: center;">PRs Merged</th>
                    <th style="text-align: center;">PRs Created</th>
                    <th style="text-align: center;">Reviews</th>
                    <th style="text-align: center;">Work Items</th>
                    <th style="text-align: center;">Commits</th>
                    <th style="text-align: center;">Activity Score</th>
                    <th>Last Active</th>
                </tr>
"@
        $rank = 1
        foreach ($person in $sortedActivity) {
            $rankClass = switch ($rank) { 1 { "gold" } 2 { "silver" } 3 { "bronze" } default { "" } }
            $rankDisplay = switch ($rank) { 1 { "<i class='fas fa-trophy' style='color: #FFD700;'></i>" } 2 { "<i class='fas fa-medal' style='color: #C0C0C0;'></i>" } 3 { "<i class='fas fa-medal' style='color: #CD7F32;'></i>" } default { $rank } }
            $lastActive = if ($person.LastActivity) { $person.LastActivity.ToString("MMM dd, yyyy") } else { "-" }
            $scoreColor = if ($person.Score -ge 50) { "var(--success)" } elseif ($person.Score -ge 20) { "var(--azure-blue)" } else { "var(--gray-90)" }
            
            $htmlContent += @"
                <tr>
                    <td style="text-align: center;">$rankDisplay</td>
                    <td><strong>$($person.Name)</strong></td>
                    <td style="text-align: center;"><span class="badge badge-success">$($person.PRsMerged)</span></td>
                    <td style="text-align: center;"><span class="badge badge-info">$($person.PRsCreated)</span></td>
                    <td style="text-align: center;">$($person.PRsReviewed)</td>
                    <td style="text-align: center;">$($person.WorkItemsAssigned)</td>
                    <td style="text-align: center;">$($person.Commits)</td>
                    <td style="text-align: center; font-weight: 700; color: $scoreColor;">$($person.Score)</td>
                    <td>$lastActive</td>
                </tr>
"@
            $rank++
        }
        $htmlContent += "</table></div>"
    }

    # Project Statistics Table
    if ($ProjectStats.Count -gt 0) {
        $htmlContent += @"
        <div class="section">
            <h2><i class="fas fa-folder"></i> Project Statistics</h2>
            <table>
                <tr>
                    <th>Project</th>
                    <th>Repos</th>
                    <th>Commits</th>
                    <th>Pipelines</th>
                    <th>Builds</th>
                    <th>Success Rate</th>
                    <th>Work Items</th>
                    <th>PRs</th>
                </tr>
"@
        foreach ($project in ($ProjectStats | Sort-Object Commits -Descending)) {
            $successBadge = if ($project.BuildSuccessRate -ge 80) { "badge-success" } 
                           elseif ($project.BuildSuccessRate -ge 50) { "badge-warning" } 
                           else { "badge-danger" }
            
            $htmlContent += @"
                <tr>
                    <td><strong>$($project.Project)</strong></td>
                    <td>$($project.Repositories)</td>
                    <td>$($project.Commits)</td>
                    <td>$($project.Pipelines)</td>
                    <td>$($project.Builds)</td>
                    <td><span class="badge $successBadge">$($project.BuildSuccessRate)%</span></td>
                    <td>$($project.WorkItems)</td>
                    <td>$($project.PullRequests)</td>
                </tr>
"@
        }
        $htmlContent += "</table></div>"
    }

    # User Statistics
    if ($LoadedData.Users) {
        $LicenseGroups = $LoadedData.Users.data | Group-Object license | Sort-Object Count -Descending
        $StatusGroups = $LoadedData.Users.data | Group-Object status | Sort-Object Count -Descending
        
        $htmlContent += @"
        <div class="section">
            <h2><i class="fas fa-user-group"></i> User Analytics</h2>
            <div class="two-col">
                <div>
                    <h4 style="margin-bottom: 15px; color: var(--gray-90);">License Distribution</h4>
                    <table>
                        <tr><th>License Type</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($license in $LicenseGroups) {
            $percentage = [math]::Round(($license.Count / $Stats.TotalUsers) * 100, 1)
            $htmlContent += "<tr><td>$($license.Name)</td><td>$($license.Count)</td><td>$percentage%</td></tr>"
        }
        $htmlContent += @"
                    </table>
                </div>
                <div>
                    <h4 style="margin-bottom: 15px; color: var(--gray-90);">User Status</h4>
                    <table>
                        <tr><th>Status</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($status in $StatusGroups) {
            $percentage = [math]::Round(($status.Count / $Stats.TotalUsers) * 100, 1)
            $badgeClass = if ($status.Name -eq "active") { "badge-success" } else { "badge-warning" }
            $htmlContent += "<tr><td><span class='badge $badgeClass'>$($status.Name)</span></td><td>$($status.Count)</td><td>$percentage%</td></tr>"
        }
        $htmlContent += "</table></div></div></div>"
    }

    # Work Items
    if ($LoadedData.WorkItems -and $LoadedData.WorkItems.data.Count -gt 0) {
        $WorkItemTypes = $LoadedData.WorkItems.data | Group-Object workItemType | Sort-Object Count -Descending
        $WorkItemStates = $LoadedData.WorkItems.data | Group-Object state | Sort-Object Count -Descending | Select-Object -First 8
        
        $htmlContent += @"
        <div class="section">
            <h2><i class="fas fa-clipboard-list"></i> Work Items Analysis</h2>
            <div class="two-col">
                <div>
                    <h4 style="margin-bottom: 15px; color: var(--gray-90);">By Type</h4>
                    <table>
                        <tr><th>Type</th><th>Count</th><th>Distribution</th></tr>
"@
        foreach ($wiType in $WorkItemTypes) {
            $percentage = [math]::Round(($wiType.Count / $Stats.TotalWorkItems) * 100, 1)
            $htmlContent += @"
                        <tr>
                            <td>$($wiType.Name)</td>
                            <td>$($wiType.Count)</td>
                            <td>
                                <div class="progress-bar">
                                    <div class="progress-fill" style="width: $percentage%; background: var(--azure-blue);">$percentage%</div>
                                </div>
                            </td>
                        </tr>
"@
        }
        $htmlContent += @"
                    </table>
                </div>
                <div>
                    <h4 style="margin-bottom: 15px; color: var(--gray-90);">By State</h4>
                    <table>
                        <tr><th>State</th><th>Count</th><th>Percentage</th></tr>
"@
        foreach ($state in $WorkItemStates) {
            $percentage = [math]::Round(($state.Count / $Stats.TotalWorkItems) * 100, 1)
            $htmlContent += "<tr><td>$($state.Name)</td><td>$($state.Count)</td><td>$percentage%</td></tr>"
        }
        $htmlContent += "</table></div></div></div>"
    }

    # Repository Details
    if ($LoadedData.GitRepositories -and $LoadedData.GitRepositories.data.Count -gt 0) {
        $htmlContent += @"
        <div class="section">
            <h2><i class="fas fa-code-branch"></i> Repository Details</h2>
            <table>
                <tr>
                    <th>Project</th>
                    <th>Repository</th>
                    <th>Default Branch</th>
                    <th>Size</th>
                    <th>Commits</th>
                    <th>Last Activity</th>
                </tr>
"@
        foreach ($repo in ($LoadedData.GitRepositories.data | Sort-Object project, name)) {
            $size = if ($repo.size -eq 0) { "-" } elseif ($repo.size -lt 1024) { "$($repo.size) B" } elseif ($repo.size -lt 1048576) { "$([math]::Round($repo.size/1024, 1)) KB" } else { "$([math]::Round($repo.size/1048576, 1)) MB" }
            $branch = if ($repo.defaultBranch) { $repo.defaultBranch -replace "refs/heads/", "" } else { "-" }
            
            $repoCommits = 0
            $lastCommitDate = "-"
            if ($LoadedData.GitCommits) {
                $commits = $LoadedData.GitCommits.data | Where-Object { $_.repositoryId -eq $repo.repositoryId }
                $repoCommits = $commits.Count
                if ($commits.Count -gt 0) {
                    $latestCommit = $commits | Sort-Object authorDate -Descending | Select-Object -First 1
                    if ($latestCommit.authorDate) {
                        $lastCommitDate = ([datetime]$latestCommit.authorDate).ToString("MMM dd, yyyy")
                    }
                }
            }
            
            $htmlContent += @"
                <tr>
                    <td>$($repo.project)</td>
                    <td><strong>$($repo.name)</strong></td>
                    <td><span class="badge badge-info">$branch</span></td>
                    <td>$size</td>
                    <td>$repoCommits</td>
                    <td>$lastCommitDate</td>
                </tr>
"@
        }
        $htmlContent += "</table></div>"
    }

    # Footer and Chart Scripts
    $htmlContent += @"
        <div class="footer">
            <p>Generated by Azure DevOps Statistics Pipeline</p>
            <p>Data Path: $DataPath</p>
        </div>
    </div>

    <script>
        // Project Activity Chart
        new Chart(document.getElementById('projectChart'), {
            type: 'bar',
            data: {
                labels: [$projectLabels],
                datasets: [
                    { label: 'Commits', data: [$projectCommits], backgroundColor: '#0078d4' },
                    { label: 'Builds', data: [$projectBuilds], backgroundColor: '#00bcf2' },
                    { label: 'Work Items', data: [$projectWorkItems], backgroundColor: '#8764b8' }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom' } },
                scales: { y: { beginAtZero: true } }
            }
        });

        // Build Status Chart
        new Chart(document.getElementById('buildChart'), {
            type: 'doughnut',
            data: {
                labels: ['Succeeded', 'Failed', 'Canceled', 'Other'],
                datasets: [{
                    data: [$buildSucceeded, $buildFailed, $buildCanceled, $buildOther],
                    backgroundColor: ['#107c10', '#d13438', '#ff8c00', '#8a8886']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { 
                    legend: { position: 'bottom' },
                    title: { display: true, text: '$OverallBuildSuccessRate% Success Rate', font: { size: 16 } }
                }
            }
        });

        // Day of Week Chart
        new Chart(document.getElementById('dayChart'), {
            type: 'bar',
            data: {
                labels: [$dayLabels],
                datasets: [{
                    label: 'Commits',
                    data: [$dayValues],
                    backgroundColor: '#667eea'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true } }
            }
        });

        // License Chart
        new Chart(document.getElementById('licenseChart'), {
            type: 'pie',
            data: {
                labels: [$licenseLabels],
                datasets: [{
                    data: [$licenseValues],
                    backgroundColor: [$licenseColors]
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom' } }
            }
        });
    </script>
</body>
</html>
"@

    # Write the file with proper UTF-8 encoding (with BOM for browser compatibility)
    $Utf8BomEncoding = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($ReportPath, $htmlContent, $Utf8BomEncoding)
    Write-Host "Enhanced report generated: $ReportPath" -ForegroundColor Green
    
    if ($OpenReport) {
        Start-Process $ReportPath
    }
}

Write-Host "`nReport generation completed!" -ForegroundColor Green
