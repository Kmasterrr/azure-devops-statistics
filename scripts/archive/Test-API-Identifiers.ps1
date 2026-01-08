[CmdletBinding()]
param()

Write-Host "Testing Git Repository API with different project identifiers..." -ForegroundColor Yellow

$Organization = $env:ADOS_ORGANIZATION
$PAT = $env:ADOS_PAT

# Load modules
$ModulePath = Join-Path $PSScriptRoot "src\powershell\modules"
Import-Module (Join-Path $ModulePath "ADOS") -Force
Import-Module (Join-Path $ModulePath "AzDevOps") -Force

try {
    $AzConfig = @{
        SessionName         = 'TestSession'
        ApiVersion          = '7.1-preview.3'
        Collection          = $Organization
        PersonalAccessToken = $PAT
    }
    $AzSession = New-AzDevOpsSession @AzConfig

    $projects = Get-AzDevOpsProjectList -Session $AzSession
    
    foreach ($project in $projects) {
        Write-Host "`n=== Testing: $($project.name) ===" -ForegroundColor Green
        
        # Test with project name
        Write-Host "Testing Git repos with Project NAME: '$($project.name)'" -ForegroundColor Cyan
        try {
            $reposByName = Get-AzDevOpsGitRepositoryList -Session $AzSession -Project $project.name
            Write-Host "SUCCESS with name: Found $($reposByName.value.Count) repositories" -ForegroundColor Green
        }
        catch {
            Write-Host "FAILED with name: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Test with project ID
        Write-Host "Testing Git repos with Project ID: '$($project.id)'" -ForegroundColor Cyan
        try {
            $reposByID = Get-AzDevOpsGitRepositoryList -Session $AzSession -Project $project.id
            Write-Host "SUCCESS with ID: Found $($reposByID.value.Count) repositories" -ForegroundColor Green
        }
        catch {
            Write-Host "FAILED with ID: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Test work items with project name
        Write-Host "Testing Work Items with Project NAME: '$($project.name)'" -ForegroundColor Cyan
        try {
            $wiqlQuery = "Select [System.Id] From WorkItems Where [System.TeamProject] = '$($project.name)'"
            $workItemsByName = Invoke-AzDevOpsQuery -Session $AzSession -Project $project.name -Query $wiqlQuery
            Write-Host "SUCCESS with name: Found $($workItemsByName.WorkItems.Count) work items" -ForegroundColor Green
        }
        catch {
            Write-Host "FAILED with name: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Test work items with project ID
        Write-Host "Testing Work Items with Project ID: '$($project.id)'" -ForegroundColor Cyan
        try {
            $wiqlQuery = "Select [System.Id] From WorkItems Where [System.TeamProject] = '$($project.name)'"
            $workItemsByID = Invoke-AzDevOpsQuery -Session $AzSession -Project $project.id -Query $wiqlQuery
            Write-Host "SUCCESS with ID: Found $($workItemsByID.WorkItems.Count) work items" -ForegroundColor Green
        }
        catch {
            Write-Host "FAILED with ID: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Remove-AzDevOpsSession $AzSession.Id

}
catch {
    Write-Host "Test failed: $($_.Exception.Message)" -ForegroundColor Red
}