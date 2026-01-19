# =============================================================================
# Generate-PipelineSummary.ps1
# =============================================================================
# Generates a Markdown summary for Azure DevOps Pipeline display.
# Uses the centralized scoring weights from Config-ScoringWeights.ps1
# =============================================================================

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$DataPath = "",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "",
    
    [Parameter(Mandatory = $false)]
    [string]$Organization = ""
)

# Load centralized scoring configuration
$configPath = Join-Path $PSScriptRoot "Config-ScoringWeights.ps1"
if (Test-Path $configPath) {
    . $configPath
    Write-Host "Loaded scoring configuration" -ForegroundColor Green
} else {
    Write-Error "Config-ScoringWeights.ps1 not found at: $configPath"
    exit 1
}

# Determine data path
if (-not $DataPath) {
    $org = if ($Organization) { $Organization } else { $env:ADOS_ORGANIZATION }
    if (-not $org) {
        Write-Error "Either provide -DataPath, -Organization, or set ADOS_ORGANIZATION environment variable."
        exit 1
    }
    
    $BaseDataPath = Join-Path (Join-Path $PSScriptRoot "data") $org
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

# Determine output path
if (-not $OutputPath) {
    $OutputPath = Join-Path $DataPath "pipeline-summary.md"
}

Write-Host "Generating pipeline summary from: $DataPath" -ForegroundColor Cyan

# Load summary data
$summaryPath = Join-Path $DataPath "Collection-Summary.json"
$summary = $null
if (Test-Path $summaryPath) {
    $summary = Get-Content $summaryPath | ConvertFrom-Json
    $Organization = if ($Organization) { $Organization } elseif ($summary.Organization) { $summary.Organization } else { $env:ADOS_ORGANIZATION }
}

# Start building markdown
$md = @"
# Azure DevOps Statistics Report

**Organization:** $Organization  
**Collection Date:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

---

"@

# Overview section
if ($summary) {
    $md += @"
## Overview

| Metric | Count |
|--------|-------|
| Users | $($summary.UsersCollected) |
| Repositories | $($summary.RepositoriesCollected) |
| Commits | $($summary.CommitsCollected) |
| Pull Requests | $($summary.PullRequestsCollected) |
| Pipelines | $($summary.PipelinesCollected) |
| Builds | $($summary.BuildsCollected) |
| Work Items | $($summary.WorkItemsCollected) |
| Projects | $($summary.ProjectsProcessed) |

**Duration:** $($summary.Duration)

"@
}

# Build Team Activity data (aggregate by email when possible)
# Build user index from Users.json to resolve display names to emails
$teamActivity = @{}
$userIndex = @{}

$usersPath = Join-Path $DataPath "Users.json"
if (Test-Path $usersPath) {
    try {
        $usersData = Get-Content $usersPath | ConvertFrom-Json
        foreach ($u in $usersData.data) {
            if ($u.displayName -and $u.email) { $userIndex[$u.displayName] = $u.email }
            if ($u.principalName -and $u.email) { $userIndex[$u.principalName] = $u.email }
        }
        Write-Host "Loaded users index" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to load users for index: $($_.Exception.Message)"
    }
}

function Resolve-KeyAndEmail([string]$name, [string]$email) {
    $resolvedEmail = $email
    if (-not $resolvedEmail -and $name -and $userIndex.ContainsKey($name)) { $resolvedEmail = $userIndex[$name] }
    $key = if ($resolvedEmail) { $resolvedEmail } elseif ($name) { $name } else { 'Unknown' }
    return @{ Key = $key; Email = $resolvedEmail }
}

# Load commits
$commitsPath = Join-Path $DataPath "GitCommits.json"
if (Test-Path $commitsPath) {
    $commitsData = Get-Content $commitsPath | ConvertFrom-Json
    foreach ($commit in $commitsData.data) {
        $name = $commit.author
        $email = $commit.authorEmail
        $r = Resolve-KeyAndEmail -name $name -email $email
        $key = $r.Key
        if (-not $teamActivity.ContainsKey($key)) {
            $teamActivity[$key] = @{ Key = $key; Name = $name; Email = $r.Email; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItems = 0 }
        } else {
            if (-not $teamActivity[$key].Name -and $name) { $teamActivity[$key].Name = $name }
            if (-not $teamActivity[$key].Email -and $r.Email) { $teamActivity[$key].Email = $r.Email }
        }
        $teamActivity[$key].Commits++
    }
    Write-Host "Loaded commits data" -ForegroundColor Green
}

# Load PRs
$prsPath = Join-Path $DataPath "GitPullRequests.json"
if (Test-Path $prsPath) {
    $prsData = Get-Content $prsPath | ConvertFrom-Json
    foreach ($pr in $prsData.data) {
        $creatorName = $pr.createdBy
        $creatorEmail = $pr.createdByEmail
        $r = Resolve-KeyAndEmail -name $creatorName -email $creatorEmail
        $key = $r.Key
        if ($creatorName -or $creatorEmail) {
            if (-not $teamActivity.ContainsKey($key)) {
                $teamActivity[$key] = @{ Key = $key; Name = $creatorName; Email = $r.Email; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItems = 0 }
            } else {
                if (-not $teamActivity[$key].Name -and $creatorName) { $teamActivity[$key].Name = $creatorName }
                if (-not $teamActivity[$key].Email -and $r.Email) { $teamActivity[$key].Email = $r.Email }
            }
            $teamActivity[$key].PRsCreated++
            if ($pr.status -eq "completed") { $teamActivity[$key].PRsMerged++ }
        }
        # Count reviewers
        if ($pr.reviewers) {
            $reviewerList = $pr.reviewers -split "; "
            foreach ($reviewer in $reviewerList) {
                if ($reviewer -and $reviewer.Trim()) {
                    $reviewerName = $reviewer.Trim()
                    $rr = Resolve-KeyAndEmail -name $reviewerName -email $null
                    $revKey = $rr.Key
                    if (-not $teamActivity.ContainsKey($revKey)) {
                        $teamActivity[$revKey] = @{ Key = $revKey; Name = $reviewerName; Email = $rr.Email; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItems = 0 }
                    } else {
                        if (-not $teamActivity[$revKey].Name -and $reviewerName) { $teamActivity[$revKey].Name = $reviewerName }
                        if (-not $teamActivity[$revKey].Email -and $rr.Email) { $teamActivity[$revKey].Email = $rr.Email }
                    }
                    $teamActivity[$revKey].PRsReviewed++
                }
            }
        }
    }
    Write-Host "Loaded pull requests data" -ForegroundColor Green
}

# Load work items
$wiPath = Join-Path $DataPath "WorkItems.json"
if (Test-Path $wiPath) {
    $wiData = Get-Content $wiPath | ConvertFrom-Json
    foreach ($wi in $wiData.data) {
        if ($wi.assignedTo) {
            $assigneeName = $wi.assignedTo
            $r = Resolve-KeyAndEmail -name $assigneeName -email $null
            $key = $r.Key
            if (-not $teamActivity.ContainsKey($key)) {
                $teamActivity[$key] = @{ Key = $key; Name = $assigneeName; Email = $r.Email; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItems = 0 }
            }
            $teamActivity[$key].WorkItems++
        }
    }
    Write-Host "Loaded work items data" -ForegroundColor Green
}

# Calculate scores using centralized function
foreach ($person in $teamActivity.Keys) {
    $a = $teamActivity[$person]
    $a.Score = Get-ActivityScore -PRsMerged $a.PRsMerged -PRsCreated $a.PRsCreated -CodeReviews $a.PRsReviewed -WorkItems $a.WorkItems -Commits $a.Commits
}

# Sort and limit
$sortedActivity = $teamActivity.Values | Sort-Object { $_.Score } -Descending | Select-Object -First 10

# Team Activity section
if ($sortedActivity.Count -gt 0) {
    $formulaText = Get-ScoringFormulaText -Markdown
    $md += @"

## Team Activity

*Score = $formulaText*

| Rank | Team Member | Email | PRs Merged | PRs Created | Reviews | Work Items | Commits | Score |
|------|-------------|-------|------------|-------------|---------|------------|---------|-------|
"@
    $rank = 1
    foreach ($person in $sortedActivity) {
        $medal = switch ($rank) { 1 { "1st" } 2 { "2nd" } 3 { "3rd" } default { "${rank}th" } }
        $md += "`n| $medal | $($person.Name) | $($person.Email) | $($person.PRsMerged) | $($person.PRsCreated) | $($person.PRsReviewed) | $($person.WorkItems) | $($person.Commits) | **$($person.Score)** |"
        $rank++
    }
}

# Repositories by Project section
$reposPath = Join-Path $DataPath "GitRepositories.json"
if (Test-Path $reposPath) {
    $reposData = Get-Content $reposPath | ConvertFrom-Json
    $projectStats = $reposData.data | Group-Object project | Sort-Object Count -Descending
    
    if ($projectStats.Count -gt 0) {
        $md += @"


## Repositories by Project

| Project | Repositories |
|---------|--------------
"@
        foreach ($p in $projectStats) {
            $md += "`n| $($p.Name) | $($p.Count) |"
        }
    }
}

# Footer
$md += @"


---

**Download the full HTML report from Artifacts > AzureDevOpsStatisticsReport**
"@

# Write output file with UTF-8 BOM encoding for proper emoji display
$Utf8BomEncoding = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($OutputPath, $md, $Utf8BomEncoding)
Write-Host "Pipeline summary generated: $OutputPath" -ForegroundColor Green

# Return the path for pipeline use
return $OutputPath
