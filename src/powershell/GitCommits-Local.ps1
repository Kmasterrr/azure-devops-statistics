[CmdletBinding()]
Param
(
    [string]$OutputPath = ".\data",
    [string]$Format = "Both" # Options: JSON, CSV, Both
)

<###[Environment Variables]#####################>
$Organization = $env:ADOS_ORGANIZATION
$PAT          = $env:ADOS_PAT

if (-not $Organization -or -not $PAT) {
    Write-Error "Environment variables ADOS_ORGANIZATION and ADOS_PAT must be set. Please run Set-Environment-Local.ps1 first."
    return
}

<###[Set Paths]#################################>
$ModulePath = Join-Path $PSScriptRoot "modules"

<###[Load Modules]##############################>
Import-Module (Join-Path $ModulePath "ADOS") -Force
Import-Module (Join-Path $ModulePath "AzDevOps") -Force

<###[Script Variables]##########################>
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm K"

# Ensure output path exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$BaseFilename = Join-Path $OutputPath "GitCommits"

Write-Host "Collecting Git Commits data..." -ForegroundColor Yellow

try {
    # Create new Azure DevOps session
    Write-Verbose "Creating AzDevOps session."
    $AzConfig = @{
        SessionName         = 'ADOS'
        ApiVersion          = '7.1-preview.3'
        Collection          = $Organization
        PersonalAccessToken = $PAT
    }
    $AzSession = New-AzDevOpsSession @AzConfig

    # Get project list
    Write-Verbose "Getting project list."
    $allProjects = Get-AzDevOpsProjectList -Session $AzSession
    # Use the direct array (not .value property)
    $projects = $allProjects
    Write-Host "Found $($projects.Count) projects to process." -ForegroundColor Green

    $allCommits = @()
    $stepCounter = 0

    foreach ($project in $projects) {
        Write-ProgressHelper -Message "Processing Git Commits" -Steps $projects.Count -StepNumber ($stepCounter++)
        
        Write-Host "Processing project: $($project.name)" -ForegroundColor Cyan
        
        try {
            # Get repositories for this project
            $repositories = Get-AzDevOpsGitRepositoryList -Session $AzSession -Project $project.name
            
            foreach ($repo in $repositories.value) {
                Write-Verbose "Processing repository: $($repo.name)"
                
                # Get commits for this repository (last 1000 commits)
                $commits = Get-AzDevOpsGitCommitList `
                    -Session $AzSession `
                    -Project $project.name `
                    -Repository $repo.id `
                    -Top 1000

                foreach ($commit in $commits.value) {
                    $allCommits += [PSCustomObject]@{
                        project           = $project.name
                        projectId         = $project.id
                        repository        = $repo.name
                        repositoryId      = $repo.id
                        commitId          = $commit.commitId
                        author            = if ($commit.author) { $commit.author.name } else { "" }
                        authorEmail       = if ($commit.author) { $commit.author.email } else { "" }
                        authorDate        = if ($commit.author) { $commit.author.date } else { $null }
                        committer         = if ($commit.committer) { $commit.committer.name } else { "" }
                        committerEmail    = if ($commit.committer) { $commit.committer.email } else { "" }
                        committerDate     = if ($commit.committer) { $commit.committer.date } else { $null }
                        comment           = if ($commit.comment) { $commit.comment.replace("`n", " ").replace("`r", " ") } else { "" }
                        changeCounts      = if ($commit.changeCounts) { "$($commit.changeCounts.Add)+/$($commit.changeCounts.Edit)~/$($commit.changeCounts.Delete)-" } else { "" }
                        url               = $commit.url
                        timeStamp         = $Timestamp
                    }
                }
            }
        }
        catch {
            Write-Warning "Error processing commits for project '$($project.name)': $($_.Exception.Message)"
        }
    }

    # Export data based on format preference
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $CsvPath = "$BaseFilename.csv"
        Write-Verbose "Writing to CSV file: $CsvPath"
        $allCommits | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "Git Commits data saved to: $CsvPath" -ForegroundColor Green
    }

    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonPath = "$BaseFilename.json"
        Write-Verbose "Writing to JSON file: $JsonPath"
        
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $Timestamp
                recordCount = $allCommits.Count
                projectsProcessed = $allProjects.value.Count
                dataType = "GitCommits"
            }
            data = $allCommits
        }
        
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $JsonPath -Encoding UTF8
        Write-Host "Git Commits data saved to: $JsonPath" -ForegroundColor Green
    }

    Write-Host "Git Commits collection completed. Total commits: $($allCommits.Count)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error collecting git commits data: $($_.Exception.Message)"
}
finally {
    # Clean up session
    if ($AzSession) {
        Write-Verbose "Removing AzDevOps session."
        Remove-AzDevOpsSession $AzSession.Id
    }
}