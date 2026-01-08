[CmdletBinding()]
param()

Write-Host "Direct REST API Testing..." -ForegroundColor Yellow

$Organization = $env:ADOS_ORGANIZATION
$PAT = $env:ADOS_PAT

# Create authorization header
$encodedPAT = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$PAT"))
$headers = @{
    'Authorization' = "Basic $encodedPAT"
    'Content-Type' = 'application/json'
}

# Test direct API calls
$projects = @(
    @{name="C4AI_ANZ"; id="79870d30-e26f-4521-8591-eaa61a41bb56"},
    @{name="C4AAI_Playground"; id="10cdc211-f7db-4149-84b0-65a44a2be8b8"}
)

foreach ($project in $projects) {
    Write-Host "`n=== Testing: $($project.name) ===" -ForegroundColor Green
    
    # Test Git repositories
    Write-Host "Testing Git Repositories (direct REST API)..." -ForegroundColor Cyan
    $gitUrl = "https://dev.azure.com/$Organization/$($project.id)/_apis/git/repositories?api-version=7.1-preview.1"
    try {
        $gitResponse = Invoke-RestMethod -Uri $gitUrl -Headers $headers -Method Get
        Write-Host "SUCCESS: Found $($gitResponse.count) repositories" -ForegroundColor Green
        foreach ($repo in $gitResponse.value) {
            Write-Host "  - $($repo.name)" -ForegroundColor White
        }
    }
    catch {
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "URL: $gitUrl" -ForegroundColor Gray
    }

    # Test work items using WIQL
    Write-Host "Testing Work Items (direct REST API)..." -ForegroundColor Cyan
    $wiqlUrl = "https://dev.azure.com/$Organization/$($project.id)/_apis/wit/wiql?api-version=7.1-preview.2"
    $wiqlBody = @{
        query = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$($project.name)'"
    } | ConvertTo-Json
    
    try {
        $wiqlResponse = Invoke-RestMethod -Uri $wiqlUrl -Headers $headers -Method Post -Body $wiqlBody
        Write-Host "SUCCESS: Found $($wiqlResponse.workItems.Count) work items" -ForegroundColor Green
    }
    catch {
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "URL: $wiqlUrl" -ForegroundColor Gray
    }

    # Test builds
    Write-Host "Testing Builds (direct REST API)..." -ForegroundColor Cyan
    $buildsUrl = "https://dev.azure.com/$Organization/$($project.id)/_apis/build/builds?api-version=7.1-preview.7&`$top=5"
    try {
        $buildsResponse = Invoke-RestMethod -Uri $buildsUrl -Headers $headers -Method Get
        Write-Host "SUCCESS: Found $($buildsResponse.count) builds" -ForegroundColor Green
        foreach ($build in $buildsResponse.value) {
            Write-Host "  - Build #$($build.buildNumber) - $($build.status)" -ForegroundColor White
        }
    }
    catch {
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "URL: $buildsUrl" -ForegroundColor Gray
    }
}