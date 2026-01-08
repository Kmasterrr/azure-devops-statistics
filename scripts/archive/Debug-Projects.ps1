[CmdletBinding()]
param()

Write-Host "Debugging Project Data Structure..." -ForegroundColor Yellow

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
        SessionName         = 'DebugSession'
        ApiVersion          = '7.1-preview.3'
        Collection          = $Organization
        PersonalAccessToken = $PAT
    }
    $AzSession = New-AzDevOpsSession @AzConfig

    # Get project list and analyze structure
    Write-Host "`nGetting project list..." -ForegroundColor Yellow
    $projectData = Get-AzDevOpsProjectList -Session $AzSession
    
    Write-Host "`nRaw project data type:" -ForegroundColor Cyan
    Write-Host $projectData.GetType().FullName -ForegroundColor White
    
    Write-Host "`nProject data structure:" -ForegroundColor Cyan
    Write-Host "Has .value property: $($null -ne $projectData.value)" -ForegroundColor White
    Write-Host "Has .Count property: $($null -ne $projectData.Count)" -ForegroundColor White
    Write-Host "Is Array: $($projectData -is [Array])" -ForegroundColor White
    
    if ($projectData.value) {
        Write-Host "`nUsing .value property:" -ForegroundColor Green
        $projects = $projectData.value
        Write-Host "Projects count: $($projects.Count)" -ForegroundColor White
        
        Write-Host "`nFirst project (.value):" -ForegroundColor Cyan
        $firstProject = $projects[0]
        Write-Host "Name: '$($firstProject.name)'" -ForegroundColor White
        Write-Host "ID: '$($firstProject.id)'" -ForegroundColor White
        Write-Host "State: '$($firstProject.state)'" -ForegroundColor White
        
        Write-Host "`nAll project names (.value):" -ForegroundColor Cyan
        foreach ($proj in $projects) {
            Write-Host "- '$($proj.name)' (ID: $($proj.id))" -ForegroundColor White
        }
    }
    else {
        Write-Host "`nUsing direct array:" -ForegroundColor Green
        $projects = $projectData
        Write-Host "Projects count: $($projects.Count)" -ForegroundColor White
        
        Write-Host "`nFirst project (direct):" -ForegroundColor Cyan
        $firstProject = $projects[0]
        Write-Host "Name: '$($firstProject.name)'" -ForegroundColor White
        Write-Host "ID: '$($firstProject.id)'" -ForegroundColor White
        Write-Host "State: '$($firstProject.state)'" -ForegroundColor White
        
        Write-Host "`nAll project names (direct):" -ForegroundColor Cyan
        foreach ($proj in $projects) {
            Write-Host "- '$($proj.name)' (ID: $($proj.id))" -ForegroundColor White
        }
    }

    # Debug the raw JSON response if possible
    Write-Host "`nDebugging raw response..." -ForegroundColor Yellow
    $projectData | ConvertTo-Json -Depth 3 | Out-File "debug-projects.json"
    Write-Host "Raw project data saved to: debug-projects.json" -ForegroundColor Cyan

    # Clean up
    Remove-AzDevOpsSession $AzSession.Id

}
catch {
    Write-Host "❌ Debug failed: $($_.Exception.Message)" -ForegroundColor Red
}