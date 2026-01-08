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

$BaseFilename = Join-Path $OutputPath "ProjectStatistics"

Write-Host "Collecting Project Statistics data..." -ForegroundColor Yellow

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

    $projectStats = @()
    $stepCounter = 0

    foreach ($project in $projects) {
        Write-ProgressHelper -Message "Processing Project Statistics" -Steps $projects.Count -StepNumber ($stepCounter++)
        
        Write-Host "Processing project: $($project.name)" -ForegroundColor Cyan
        
        $stats = [ordered]@{
            Project                     = $project.name
            ProjectId                   = $project.id
            ProjectDescription          = $project.description
            ProjectState                = $project.state
            ProjectVisibility           = $project.visibility
            TimeStamp                   = $Timestamp
            Repositories                = 0
            BuildPipelines              = 0
            Builds                      = 0
            BuildsCompleted             = 0
            BuildCompletionPercentage   = "0%"
            ReleasePipelines            = 0
            Releases                    = 0
            ReleasesCompleted           = 0
            ReleaseCompletionPercentage = "0%"
            WorkItems                   = 0
            PullRequests                = 0
            Commits                     = 0
        }
        
        try {
            # Get repositories
            Write-Verbose " |- Getting repositories."
            $repositories = Get-AzDevOpsGitRepositoryList -Session $AzSession -Project $project.name
            $stats['Repositories'] = $repositories.value.Count

            # Get build definitions
            Write-Verbose " |- Getting build pipelines."
            $buildDefinitions = Get-AzDevOpsBuildDefinitionList `
                -Session $AzSession `
                -Project "$($project.name)" `
                -Top 1000
            $stats['BuildPipelines'] = $buildDefinitions.Count

            # Get builds
            Write-Verbose " |- Getting builds."
            $builds = Get-AzDevOpsBuildList `
                -Session $AzSession `
                -Project "$($project.name)" `
                -Top 1000
            $stats['Builds'] = $builds.Count
            
            $completedBuilds = ($builds | Where-Object { $_.status -eq 'completed' }).Count
            $stats['BuildsCompleted'] = $completedBuilds
            
            if ($stats['Builds'] -gt 0) {
                $buildPercentage = ($completedBuilds / $stats['Builds'] * 100).ToString("F1") + "%"
                $stats['BuildCompletionPercentage'] = $buildPercentage
            }

            # Get release definitions
            Write-Verbose " |- Getting release pipelines."
            $releaseDefinitions = Get-AzDevOpsReleaseDefinitionList `
                -Session $AzSession `
                -Project "$($project.name)" `
                -Top 1000
            $stats['ReleasePipelines'] = $releaseDefinitions.Count

            # Get releases (simplified version)
            Write-Verbose " |- Getting releases."
            $releases = Get-AzDevOpsReleaseList `
                -Session $AzSession `
                -Project "$($project.name)" `
                -Top 500
            $stats['Releases'] = $releases.Count
            
            $completedReleases = ($releases | Where-Object { $_.status -eq 'active' }).Count
            $stats['ReleasesCompleted'] = $completedReleases
            
            if ($stats['Releases'] -gt 0) {
                $releasePercentage = ($completedReleases / $stats['Releases'] * 100).ToString("F1") + "%"
                $stats['ReleaseCompletionPercentage'] = $releasePercentage
            }

            # Get work items count (simplified query)
            Write-Verbose " |- Getting work items count."
            try {
                $wiqlQuery = "Select [System.Id] From WorkItems Where [System.TeamProject] = @project AND [State] <> 'Removed'"
                $workItemsResult = Invoke-AzDevOpsQuery `
                    -Session $AzSession `
                    -Project $project.id `
                    -Query $wiqlQuery
                $stats['WorkItems'] = $workItemsResult.WorkItems.Count
            }
            catch {
                Write-Warning "Could not get work items count for project '$($project.name)'"
                $stats['WorkItems'] = 0
            }

            # Get pull requests count
            Write-Verbose " |- Getting pull requests count."
            $totalPRs = 0
            $totalCommits = 0
            foreach ($repo in $repositories.value) {
                try {
                    $prs = Get-AzDevOpsGitPullRequestList `
                        -Session $AzSession `
                        -Project $project.name `
                        -Repository $repo.id `
                        -Top 500
                    $totalPRs += $prs.value.Count

                    $commits = Get-AzDevOpsGitCommitList `
                        -Session $AzSession `
                        -Project $project.name `
                        -Repository $repo.id `
                        -Top 500
                    $totalCommits += $commits.value.Count
                }
                catch {
                    Write-Warning "Could not get Git data for repository '$($repo.name)' in project '$($project.name)'"
                }
            }
            $stats['PullRequests'] = $totalPRs
            $stats['Commits'] = $totalCommits

        }
        catch {
            Write-Warning "Error collecting statistics for project '$($project.name)': $($_.Exception.Message)"
        }

        $projectStats += [PSCustomObject]$stats
    }

    # Export data based on format preference
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $CsvPath = "$BaseFilename.csv"
        Write-Verbose "Writing to CSV file: $CsvPath"
        $projectStats | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "Project Statistics data saved to: $CsvPath" -ForegroundColor Green
    }

    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonPath = "$BaseFilename.json"
        Write-Verbose "Writing to JSON file: $JsonPath"
        
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $Timestamp
                recordCount = $projectStats.Count
                dataType = "ProjectStatistics"
            }
            data = $projectStats
        }
        
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $JsonPath -Encoding UTF8
        Write-Host "Project Statistics data saved to: $JsonPath" -ForegroundColor Green
    }

    Write-Host "Project Statistics collection completed. Total projects: $($projectStats.Count)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error collecting project statistics: $($_.Exception.Message)"
}
finally {
    # Clean up session
    if ($AzSession) {
        Write-Verbose "Removing AzDevOps session."
        Remove-AzDevOpsSession $AzSession.Id
    }
}