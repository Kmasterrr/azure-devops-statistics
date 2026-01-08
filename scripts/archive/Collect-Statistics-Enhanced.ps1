[CmdletBinding()]
param 
(
    [string]$OutputPath = (Join-Path $PSScriptRoot "data"),
    [string]$Format = "Both" # Options: JSON, CSV, Both
)

<###[Environment Variables]#####################>
$Organization = $env:ADOS_ORGANIZATION
$PAT = $env:ADOS_PAT
$OutputBasePath = if ($env:LOCAL_DATA_PATH) { $env:LOCAL_DATA_PATH } else { $OutputPath }

if (-not $Organization -or -not $PAT) {
    Write-Error "Environment variables ADOS_ORGANIZATION and ADOS_PAT must be set. Please run Set-Environment-Local.ps1 first."
    exit 1
}

<###[Set Paths]#################################>
$ModulePath = Join-Path $PSScriptRoot "src\powershell\modules"

<###[Load Modules]##############################>
Import-Module (Join-Path $ModulePath "ADOS") -Force
Import-Module (Join-Path $ModulePath "AzDevOps") -Force

<###[Script Variables]##########################>
$FileDate = Get-Date -Format "yyyy-MM-dd"
$TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$DataPath = Join-Path (Join-Path $OutputBasePath $Organization) $FileDate

# Create data path if it does not exist
if (-not (Test-Path $DataPath)) {
    New-Item -ItemType Directory -Path $DataPath -Force | Out-Null
    Write-Host "Created output directory: $DataPath" -ForegroundColor Green
}

# Create authorization header for direct API calls
$encodedPAT = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$PAT"))
$headers = @{
    'Authorization' = "Basic $encodedPAT"
    'Content-Type' = 'application/json'
}

Write-Host "🚀 Collecting Comprehensive Azure DevOps Statistics from: $Organization" -ForegroundColor Cyan
Write-Host "Output Directory: $DataPath" -ForegroundColor Yellow
Write-Host "Output Format: $Format" -ForegroundColor Yellow
Write-Host "Using Direct REST API calls for reliability" -ForegroundColor Green

$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    # Create session for user data (this works)
    $AzConfig = @{
        SessionName         = 'ADOS'
        ApiVersion          = '7.1-preview.3'
        Collection          = $Organization
        PersonalAccessToken = $PAT
    }
    $AzSession = New-AzDevOpsSession @AzConfig

    # Get projects using existing module (this works)
    Write-Host "`n📁 Collecting Projects..." -ForegroundColor Yellow
    $projects = Get-AzDevOpsProjectList -Session $AzSession
    Write-Host "Found $($projects.Count) projects" -ForegroundColor Green

    # Collect Users (this works)
    Write-Host "`n👥 Collecting Users..." -ForegroundColor Yellow
    $users = Get-AzDevOpsUserEntitlementList -Session $AzSession
    $userData = @()
    foreach($user in $users) {
        $userData += [PSCustomObject]@{
            id               = $user.id
            descriptor       = $user.user.descriptor
            principalName    = $user.user.principalName
            displayName      = $user.user.displayName
            email            = $user.user.mailAddress
            origin           = $user.user.origin
            originId         = $user.user.originId
            kind             = $user.user.subjectKind
            type             = $user.user.metaType
            domain           = $user.user.domain
            status           = $user.accessLevel.status
            license          = $user.accessLevel.licenseDisplayName
            licenseType      = $user.accessLevel.accountLicenseType
            source           = $user.accessLevel.assignmentSource
            dateCreated      = $user.dateCreated
            lastAccessedDate = $user.lastAccessedDate
            timeStamp        = $TimeStamp
        }
    }

    # Save Users data
    $BaseFilename = Join-Path $DataPath "Users"
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $userData | Export-Csv -Path "$BaseFilename.csv" -NoTypeInformation
        Write-Host "Users data saved to: $BaseFilename.csv" -ForegroundColor Green
    }
    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $TimeStamp
                recordCount = $userData.Count
                dataType = "Users"
            }
            data = $userData
        }
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$BaseFilename.json" -Encoding UTF8
        Write-Host "Users data saved to: $BaseFilename.json" -ForegroundColor Green
    }
    Write-Host "Users collection completed. Total users: $($userData.Count)" -ForegroundColor Cyan

    # Collect Git Repositories using direct API
    Write-Host "`n📚 Collecting Git Repositories..." -ForegroundColor Yellow
    $allRepos = @()
    foreach ($project in $projects) {
        Write-Host "Processing repositories for: $($project.name)" -ForegroundColor Cyan
        $gitUrl = "https://dev.azure.com/$Organization/$($project.id)/_apis/git/repositories?api-version=7.1-preview.1"
        try {
            $gitResponse = Invoke-RestMethod -Uri $gitUrl -Headers $headers -Method Get
            foreach ($repo in $gitResponse.value) {
                $allRepos += [PSCustomObject]@{
                    project      = $project.name
                    projectId    = $project.id
                    repositoryId = $repo.id
                    name         = $repo.name
                    url          = $repo.url
                    defaultBranch = $repo.defaultBranch
                    size         = $repo.size
                    timeStamp    = $TimeStamp
                }
            }
        }
        catch {
            Write-Warning "Error getting repositories for project '$($project.name)': $($_.Exception.Message)"
        }
    }

    # Save Git Repositories data
    $BaseFilename = Join-Path $DataPath "GitRepositories"
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $allRepos | Export-Csv -Path "$BaseFilename.csv" -NoTypeInformation
        Write-Host "Git Repositories data saved to: $BaseFilename.csv" -ForegroundColor Green
    }
    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $TimeStamp
                recordCount = $allRepos.Count
                dataType = "GitRepositories"
            }
            data = $allRepos
        }
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$BaseFilename.json" -Encoding UTF8
        Write-Host "Git Repositories data saved to: $BaseFilename.json" -ForegroundColor Green
    }
    Write-Host "Git Repositories collection completed. Total repositories: $($allRepos.Count)" -ForegroundColor Cyan

    # Collect Git Commits using direct API
    Write-Host "`n📝 Collecting Git Commits..." -ForegroundColor Yellow
    $allCommits = @()
    foreach ($repo in $allRepos) {
        Write-Host "Processing commits for repository: $($repo.name) in $($repo.project)" -ForegroundColor Cyan
        $commitsUrl = "https://dev.azure.com/$Organization/$($repo.projectId)/_apis/git/repositories/$($repo.repositoryId)/commits?`$top=100&api-version=7.1-preview.1"
        try {
            $commitsResponse = Invoke-RestMethod -Uri $commitsUrl -Headers $headers -Method Get
            foreach ($commit in $commitsResponse.value) {
                $allCommits += [PSCustomObject]@{
                    project       = $repo.project
                    projectId     = $repo.projectId
                    repository    = $repo.name
                    repositoryId  = $repo.repositoryId
                    commitId      = $commit.commitId
                    author        = $commit.author.name
                    authorEmail   = $commit.author.email
                    authorDate    = $commit.author.date
                    committer     = $commit.committer.name
                    committerDate = $commit.committer.date
                    comment       = $commit.comment
                    url           = $commit.url
                    timeStamp     = $TimeStamp
                }
            }
        }
        catch {
            Write-Warning "Error getting commits for repository '$($repo.name)': $($_.Exception.Message)"
        }
    }

    # Save Git Commits data
    $BaseFilename = Join-Path $DataPath "GitCommits"
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $allCommits | Export-Csv -Path "$BaseFilename.csv" -NoTypeInformation
        Write-Host "Git Commits data saved to: $BaseFilename.csv" -ForegroundColor Green
    }
    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $TimeStamp
                recordCount = $allCommits.Count
                dataType = "GitCommits"
            }
            data = $allCommits
        }
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$BaseFilename.json" -Encoding UTF8
        Write-Host "Git Commits data saved to: $BaseFilename.json" -ForegroundColor Green
    }
    Write-Host "Git Commits collection completed. Total commits: $($allCommits.Count)" -ForegroundColor Cyan

    # Collect Build Definitions (Pipelines) using direct API
    Write-Host "`n🔧 Collecting Build Pipelines..." -ForegroundColor Yellow
    $allPipelines = @()
    foreach ($project in $projects) {
        Write-Host "Processing build pipelines for: $($project.name)" -ForegroundColor Cyan
        $pipelinesUrl = "https://dev.azure.com/$Organization/$($project.id)/_apis/build/definitions?api-version=7.1-preview.7"
        try {
            $pipelinesResponse = Invoke-RestMethod -Uri $pipelinesUrl -Headers $headers -Method Get
            foreach ($pipeline in $pipelinesResponse.value) {
                $allPipelines += [PSCustomObject]@{
                    project      = $project.name
                    projectId    = $project.id
                    pipelineId   = $pipeline.id
                    name         = $pipeline.name
                    path         = $pipeline.path
                    type         = $pipeline.type
                    quality      = $pipeline.quality
                    repository   = if ($pipeline.repository) { $pipeline.repository.name } else { "N/A" }
                    repositoryId = if ($pipeline.repository) { $pipeline.repository.id } else { "N/A" }
                    queueStatus  = $pipeline.queueStatus
                    revision     = $pipeline.revision
                    createdDate  = $pipeline.createdDate
                    authoredBy   = if ($pipeline.authoredBy) { $pipeline.authoredBy.displayName } else { "N/A" }
                    timeStamp    = $TimeStamp
                }
            }
        }
        catch {
            Write-Warning "Error getting pipelines for project '$($project.name)': $($_.Exception.Message)"
        }
    }

    # Save Build Pipelines data
    $BaseFilename = Join-Path $DataPath "BuildPipelines"
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $allPipelines | Export-Csv -Path "$BaseFilename.csv" -NoTypeInformation
        Write-Host "Build Pipelines data saved to: $BaseFilename.csv" -ForegroundColor Green
    }
    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $TimeStamp
                recordCount = $allPipelines.Count
                dataType = "BuildPipelines"
            }
            data = $allPipelines
        }
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$BaseFilename.json" -Encoding UTF8
        Write-Host "Build Pipelines data saved to: $BaseFilename.json" -ForegroundColor Green
    }
    Write-Host "Build Pipelines collection completed. Total pipelines: $($allPipelines.Count)" -ForegroundColor Cyan

    # Collect Builds using direct API
    Write-Host "`n🏗️ Collecting Builds..." -ForegroundColor Yellow
    $allBuilds = @()
    foreach ($project in $projects) {
        Write-Host "Processing builds for: $($project.name)" -ForegroundColor Cyan
        $buildsUrl = "https://dev.azure.com/$Organization/$($project.id)/_apis/build/builds?`$top=100&api-version=7.1-preview.7"
        try {
            $buildsResponse = Invoke-RestMethod -Uri $buildsUrl -Headers $headers -Method Get
            foreach ($build in $buildsResponse.value) {
                $allBuilds += [PSCustomObject]@{
                    project        = $project.name
                    projectId      = $project.id
                    buildId        = $build.id
                    buildNumber    = $build.buildNumber
                    definition     = $build.definition.name
                    definitionId   = $build.definition.id
                    status         = $build.status
                    result         = $build.result
                    queueTime      = $build.queueTime
                    startTime      = $build.startTime
                    finishTime     = $build.finishTime
                    reason         = $build.reason
                    requestedBy    = if ($build.requestedBy) { $build.requestedBy.displayName } else { "N/A" }
                    requestedFor   = if ($build.requestedFor) { $build.requestedFor.displayName } else { "N/A" }
                    sourceBranch   = $build.sourceBranch
                    sourceVersion  = $build.sourceVersion
                    timeStamp      = $TimeStamp
                }
            }
        }
        catch {
            Write-Warning "Error getting builds for project '$($project.name)': $($_.Exception.Message)"
        }
    }

    # Save Builds data
    $BaseFilename = Join-Path $DataPath "Builds"
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $allBuilds | Export-Csv -Path "$BaseFilename.csv" -NoTypeInformation
        Write-Host "Builds data saved to: $BaseFilename.csv" -ForegroundColor Green
    }
    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $TimeStamp
                recordCount = $allBuilds.Count
                dataType = "Builds"
            }
            data = $allBuilds
        }
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$BaseFilename.json" -Encoding UTF8
        Write-Host "Builds data saved to: $BaseFilename.json" -ForegroundColor Green
    }
    Write-Host "Builds collection completed. Total builds: $($allBuilds.Count)" -ForegroundColor Cyan

    # Collect Work Items using direct API
    Write-Host "`n📋 Collecting Work Items..." -ForegroundColor Yellow
    $allWorkItems = @()
    foreach ($project in $projects) {
        Write-Host "Processing work items for: $($project.name)" -ForegroundColor Cyan
        
        # Get work item IDs
        $wiqlUrl = "https://dev.azure.com/$Organization/$($project.id)/_apis/wit/wiql?api-version=7.1-preview.2"
        $wiqlBody = @{
            query = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$($project.name)'"
        } | ConvertTo-Json
        
        try {
            $wiqlResponse = Invoke-RestMethod -Uri $wiqlUrl -Headers $headers -Method Post -Body $wiqlBody
            
            if ($wiqlResponse.workItems.Count -gt 0) {
                # Get work item details in batches
                $ids = $wiqlResponse.workItems | ForEach-Object { $_.id }
                $idsString = $ids -join ","
                $workItemsUrl = "https://dev.azure.com/$Organization/$($project.id)/_apis/wit/workitems?ids=$idsString&api-version=7.1-preview.3"
                
                $workItemsResponse = Invoke-RestMethod -Uri $workItemsUrl -Headers $headers -Method Get
                
                foreach ($wi in $workItemsResponse.value) {
                    $allWorkItems += [PSCustomObject]@{
                        project          = $project.name
                        projectId        = $project.id
                        workItemId       = $wi.id
                        workItemType     = $wi.fields."System.WorkItemType"
                        title            = $wi.fields."System.Title"
                        state            = $wi.fields."System.State"
                        assignedTo       = if ($wi.fields."System.AssignedTo") { $wi.fields."System.AssignedTo".displayName } else { $null }
                        createdBy        = if ($wi.fields."System.CreatedBy") { $wi.fields."System.CreatedBy".displayName } else { $null }
                        createdDate      = $wi.fields."System.CreatedDate"
                        changedDate      = $wi.fields."System.ChangedDate"
                        priority         = $wi.fields."Microsoft.VSTS.Common.Priority"
                        storyPoints      = $wi.fields."Microsoft.VSTS.Scheduling.StoryPoints"
                        timeStamp        = $TimeStamp
                    }
                }
            }
        }
        catch {
            Write-Warning "Error getting work items for project '$($project.name)': $($_.Exception.Message)"
        }
    }

    # Save Work Items data
    $BaseFilename = Join-Path $DataPath "WorkItems"
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $allWorkItems | Export-Csv -Path "$BaseFilename.csv" -NoTypeInformation
        Write-Host "Work Items data saved to: $BaseFilename.csv" -ForegroundColor Green
    }
    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $TimeStamp
                recordCount = $allWorkItems.Count
                dataType = "WorkItems"
            }
            data = $allWorkItems
        }
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$BaseFilename.json" -Encoding UTF8
        Write-Host "Work Items data saved to: $BaseFilename.json" -ForegroundColor Green
    }
    Write-Host "Work Items collection completed. Total work items: $($allWorkItems.Count)" -ForegroundColor Cyan

    # Clean up session
    Remove-AzDevOpsSession $AzSession.Id

    $StopWatch.Stop()

    # Create summary
    $SummaryData = @{
        Organization = $Organization
        CollectionDate = $FileDate
        CollectionTime = $TimeStamp
        OutputPath = $DataPath
        Format = $Format
        Duration = $StopWatch.Elapsed.ToString()
        UsersCollected = $userData.Count
        RepositoriesCollected = $allRepos.Count
        CommitsCollected = $allCommits.Count
        PipelinesCollected = $allPipelines.Count
        BuildsCollected = $allBuilds.Count
        WorkItemsCollected = $allWorkItems.Count
        ProjectsProcessed = $projects.Count
        FilesGenerated = @()
    }

    # List all generated files
    $GeneratedFiles = Get-ChildItem -Path $DataPath -File | Select-Object Name, Length, LastWriteTime
    $SummaryData.FilesGenerated = $GeneratedFiles

    # Save summary
    $SummaryPath = Join-Path $DataPath "Collection-Summary.json"
    $SummaryData | ConvertTo-Json -Depth 3 | Out-File -FilePath $SummaryPath -Encoding UTF8

    Write-Host "`n🎉 Comprehensive data collection completed successfully!" -ForegroundColor Green
    Write-Host "Duration: $($StopWatch.Elapsed)" -ForegroundColor Yellow
    Write-Host "📊 Summary:" -ForegroundColor White
    Write-Host "  Users: $($userData.Count)" -ForegroundColor White
    Write-Host "  Repositories: $($allRepos.Count)" -ForegroundColor White
    Write-Host "  Commits: $($allCommits.Count)" -ForegroundColor White
    Write-Host "  Pipelines: $($allPipelines.Count)" -ForegroundColor White
    Write-Host "  Builds: $($allBuilds.Count)" -ForegroundColor White
    Write-Host "  Work Items: $($allWorkItems.Count)" -ForegroundColor White
    Write-Host "  Files generated: $($GeneratedFiles.Count)" -ForegroundColor Yellow
    Write-Host "Summary saved to: $SummaryPath" -ForegroundColor Cyan

}
catch {
    Write-Error "Collection failed: $($_.Exception.Message)"
}
finally {
    $StopWatch.Stop()
}