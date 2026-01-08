[CmdletBinding()]
param()

Write-Host "Detailed API Diagnostics for Data Collection Issues..." -ForegroundColor Yellow

# Check environment variables
$Organization = $env:ADOS_ORGANIZATION
$PAT = $env:ADOS_PAT

if (-not $Organization -or -not $PAT) {
    Write-Error "Environment variables not set. Please run Set-Environment-Local.ps1 first."
    exit 1
}

# Load modules
$ModulePath = Join-Path $PSScriptRoot "src\powershell\modules"
Import-Module (Join-Path $ModulePath "ADOS") -Force
Import-Module (Join-Path $ModulePath "AzDevOps") -Force

try {
    # Create session
    $AzConfig = @{
        SessionName         = 'DiagnosticSession'
        ApiVersion          = '7.1-preview.3'
        Collection          = $Organization
        PersonalAccessToken = $PAT
    }
    $AzSession = New-AzDevOpsSession @AzConfig

    # Get projects
    Write-Host "`n=== PROJECT ANALYSIS ===" -ForegroundColor Green
    $projects = Get-AzDevOpsProjectList -Session $AzSession
    Write-Host "Found $($projects.Count) projects" -ForegroundColor Cyan
    
    foreach ($project in $projects) {
        Write-Host "`n--- Project: $($project.name) ---" -ForegroundColor Yellow
        Write-Host "ID: $($project.id)" -ForegroundColor White
        Write-Host "State: $($project.state)" -ForegroundColor White
        
        # Test Git repositories for this project
        Write-Host "`nTesting Git Repositories..." -ForegroundColor Cyan
        try {
            $repos = Get-AzDevOpsGitRepositoryList -Session $AzSession -Project $project.name
            Write-Host "Found $($repos.value.Count) repositories" -ForegroundColor Green
            foreach ($repo in $repos.value) {
                Write-Host "  - $($repo.name) (ID: $($repo.id))" -ForegroundColor White
            }
        }
        catch {
            Write-Host "Git Repository Error: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Test Work Items for this project
        Write-Host "`nTesting Work Items..." -ForegroundColor Cyan
        try {
            $wiqlQuery = "Select [System.Id] From WorkItems Where [System.TeamProject] = @project"
            $workItems = Invoke-AzDevOpsQuery -Session $AzSession -Project $project.id -Query $wiqlQuery
            Write-Host "Found $($workItems.WorkItems.Count) work items" -ForegroundColor Green
        }
        catch {
            Write-Host "Work Items Error: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Test Build Definitions for this project  
        Write-Host "`nTesting Build Definitions..." -ForegroundColor Cyan
        try {
            $buildDefs = Get-AzDevOpsBuildDefinitionList -Session $AzSession -Project $project.name -Top 100
            Write-Host "Found $($buildDefs.Count) build definitions" -ForegroundColor Green
            foreach ($buildDef in $buildDefs) {
                Write-Host "  - $($buildDef.name) (ID: $($buildDef.id))" -ForegroundColor White
            }
        }
        catch {
            Write-Host "Build Definitions Error: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Test Builds for this project
        Write-Host "`nTesting Builds..." -ForegroundColor Cyan
        try {
            $builds = Get-AzDevOpsBuildList -Session $AzSession -Project $project.name -Top 10
            Write-Host "Found $($builds.Count) recent builds" -ForegroundColor Green
            foreach ($build in $builds) {
                Write-Host "  - Build #$($build.buildNumber) - $($build.status)" -ForegroundColor White
            }
        }
        catch {
            Write-Host "Builds Error: $($_.Exception.Message)" -ForegroundColor Red
        }

        Write-Host "`n$('='*50)" -ForegroundColor Gray
    }

    # Clean up
    Remove-AzDevOpsSession $AzSession.Id
    Write-Host "`nDiagnostic completed!" -ForegroundColor Green

}
catch {
    Write-Host "Diagnostic failed: $($_.Exception.Message)" -ForegroundColor Red
}