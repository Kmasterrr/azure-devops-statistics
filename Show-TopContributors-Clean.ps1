# Top Contributors Analysis
Write-Host "Top Contributors Analysis" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Yellow

# Load the commits data
$commitsFile = "C:\Users\george.karsas\Documents\Dev\sandbox\playground\azure-devops-statistics\data\C4AI-ANZ\2025-11-05\GitCommits.json"
if (Test-Path $commitsFile) {
    $commits = Get-Content $commitsFile | ConvertFrom-Json
    
    Write-Host "`nTop 10 Contributors by Commit Count:" -ForegroundColor Green
    $topContributors = $commits.data | Group-Object author | Sort-Object Count -Descending | Select-Object -First 10
    
    $rank = 1
    foreach ($contributor in $topContributors) {
        $medal = switch ($rank) {
            1 { "[GOLD]" }
            2 { "[SILVER]" }
            3 { "[BRONZE]" }
            default { "      " }
        }
        
        $percentage = [math]::Round(($contributor.Count / $commits.data.Count) * 100, 1)
        $latestCommit = ($commits.data | Where-Object { $_.author -eq $contributor.Name } | Sort-Object authorDate -Descending | Select-Object -First 1).authorDate
        $latestDate = if ($latestCommit) { ([datetime]$latestCommit).ToString("yyyy-MM-dd") } else { "N/A" }
        
        Write-Host "$medal $rank. $($contributor.Name)" -ForegroundColor White
        Write-Host "    Commits: $($contributor.Count) ($percentage percent of total)" -ForegroundColor Gray
        Write-Host "    Latest: $latestDate" -ForegroundColor Gray
        Write-Host ""
        $rank++
    }
    
    Write-Host "`nActivity Summary:" -ForegroundColor Green
    Write-Host "Total Commits: $($commits.data.Count)" -ForegroundColor White
    Write-Host "Active Contributors: $($topContributors.Count)" -ForegroundColor White
    Write-Host "Date Range: $(($commits.data | Sort-Object authorDate | Select-Object -First 1).authorDate.Split('T')[0]) to $(($commits.data | Sort-Object authorDate -Descending | Select-Object -First 1).authorDate.Split('T')[0])" -ForegroundColor White
    
    Write-Host "`nContributors by Project:" -ForegroundColor Green
    $contributorsByProject = $commits.data | Group-Object author, project | Sort-Object { $_.Group.Count } -Descending | Select-Object -First 15
    foreach ($cp in $contributorsByProject) {
        $parts = $cp.Name.Split(', ')
        $author = $parts[0]
        $project = $parts[1]
        Write-Host "  $author -> $project : $($cp.Count) commits" -ForegroundColor Gray
    }
    
    Write-Host "`nMost Active Repositories:" -ForegroundColor Green
    $repoActivity = $commits.data | Group-Object repository | Sort-Object Count -Descending | Select-Object -First 10
    foreach ($repo in $repoActivity) {
        Write-Host "  $($repo.Name): $($repo.Count) commits" -ForegroundColor Gray
    }
    
} else {
    Write-Host "Commits data not found. Please run data collection first." -ForegroundColor Red
}