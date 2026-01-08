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

$BaseFilename = Join-Path $OutputPath "GitPullRequests"

Write-Host "Collecting Git Pull Requests data..." -ForegroundColor Yellow

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

    $allPullRequests = @()
    $stepCounter = 0

    foreach ($project in $projects) {
        Write-ProgressHelper -Message "Processing Git Pull Requests" -Steps $projects.Count -StepNumber ($stepCounter++)
        
        Write-Host "Processing project: $($project.name)" -ForegroundColor Cyan
        
        try {
            # Get repositories for this project
            $repositories = Get-AzDevOpsGitRepositoryList -Session $AzSession -Project $project.name
            
            foreach ($repo in $repositories.value) {
                Write-Verbose "Processing repository: $($repo.name)"
                
                # Get pull requests for this repository
                $pullRequests = Get-AzDevOpsGitPullRequestList `
                    -Session $AzSession `
                    -Project $project.name `
                    -Repository $repo.id `
                    -Top 1000

                foreach ($pr in $pullRequests.value) {
                    $allPullRequests += [PSCustomObject]@{
                        project           = $project.name
                        projectId         = $project.id
                        repository        = $repo.name
                        repositoryId      = $repo.id
                        pullRequestId     = $pr.pullRequestId
                        title             = $pr.title
                        description       = if ($pr.description) { $pr.description.replace("`n", " ").replace("`r", " ") } else { "" }
                        status            = $pr.status
                        createdBy         = if ($pr.createdBy) { $pr.createdBy.displayName } else { "" }
                        createdDate       = $pr.creationDate
                        closedDate        = $pr.closedDate
                        sourceRefName     = $pr.sourceRefName
                        targetRefName     = $pr.targetRefName
                        mergeStatus       = $pr.mergeStatus
                        isDraft           = $pr.isDraft
                        url               = $pr.url
                        timeStamp         = $Timestamp
                    }
                }
            }
        }
        catch {
            Write-Warning "Error processing pull requests for project '$($project.name)': $($_.Exception.Message)"
        }
    }

    # Export data based on format preference
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $CsvPath = "$BaseFilename.csv"
        Write-Verbose "Writing to CSV file: $CsvPath"
        $allPullRequests | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "Git Pull Requests data saved to: $CsvPath" -ForegroundColor Green
    }

    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonPath = "$BaseFilename.json"
        Write-Verbose "Writing to JSON file: $JsonPath"
        
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $Timestamp
                recordCount = $allPullRequests.Count
                projectsProcessed = $allProjects.value.Count
                dataType = "GitPullRequests"
            }
            data = $allPullRequests
        }
        
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $JsonPath -Encoding UTF8
        Write-Host "Git Pull Requests data saved to: $JsonPath" -ForegroundColor Green
    }

    Write-Host "Git Pull Requests collection completed. Total pull requests: $($allPullRequests.Count)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error collecting git pull requests data: $($_.Exception.Message)"
}
finally {
    # Clean up session
    if ($AzSession) {
        Write-Verbose "Removing AzDevOps session."
        Remove-AzDevOpsSession $AzSession.Id
    }
}