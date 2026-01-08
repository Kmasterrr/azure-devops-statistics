[CmdletBinding()]
param()

Write-Host "Testing Azure DevOps Connection..." -ForegroundColor Yellow

# Check environment variables
$Organization = $env:ADOS_ORGANIZATION
$PAT = $env:ADOS_PAT

if (-not $Organization -or -not $PAT) {
    Write-Error "Environment variables not set. Please run Set-Environment-Local.ps1 first."
    exit 1
}

Write-Host "Organization: $Organization" -ForegroundColor Cyan
Write-Host "PAT: $($PAT.Substring(0,8))..." -ForegroundColor Cyan

# Load modules
$ModulePath = Join-Path $PSScriptRoot "src\powershell\modules"
Import-Module (Join-Path $ModulePath "ADOS") -Force
Import-Module (Join-Path $ModulePath "AzDevOps") -Force

try {
    # Create session
    Write-Host "`nCreating Azure DevOps session..." -ForegroundColor Yellow
    $AzConfig = @{
        SessionName         = 'TestSession'
        ApiVersion          = '7.1-preview.3'
        Collection          = $Organization
        PersonalAccessToken = $PAT
    }
    $AzSession = New-AzDevOpsSession @AzConfig
    Write-Host "✅ Session created successfully" -ForegroundColor Green

    # Test project access
    Write-Host "`nTesting project access..." -ForegroundColor Yellow
    $projects = Get-AzDevOpsProjectList -Session $AzSession
    Write-Host "✅ Found $($projects.Count) projects:" -ForegroundColor Green
    
    foreach ($project in $projects) {
        Write-Host "  - $($project.name) (ID: $($project.id))" -ForegroundColor White
        Write-Host "    State: $($project.state), Visibility: $($project.visibility)" -ForegroundColor Gray
    }

    # Test user access  
    Write-Host "`nTesting user access..." -ForegroundColor Yellow
    try {
        $users = Get-AzDevOpsUserEntitlementList -Session $AzSession
        $userCount = if ($users) { $users.Count } else { 0 }
        Write-Host "✅ Found $userCount users" -ForegroundColor Green
        if ($users -and $users.Count -gt 0) {
            $displayUsers = $users | Select-Object -First 3
            foreach ($user in $displayUsers) {
                Write-Host "  - $($user.user.displayName) ($($user.user.principalName))" -ForegroundColor White
            }
            if ($users.Count -gt 3) {
                Write-Host "  ... and $($users.Count - 3) more" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Host "❌ User access failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "   This usually means missing 'Graph' or 'Identity' permissions" -ForegroundColor Yellow
    }

    # Test group access
    Write-Host "`nTesting group access..." -ForegroundColor Yellow
    try {
        $groups = Get-AzDevOpsGroupList -Session $AzSession
        Write-Host "✅ Found $($groups.Count) groups" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Group access failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "   This usually means missing 'Graph' permissions" -ForegroundColor Yellow
    }

    # Clean up
    Remove-AzDevOpsSession $AzSession.Id
    Write-Host "`n🎉 Connection test completed!" -ForegroundColor Green

}
catch {
    Write-Host "❌ Connection failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nCommon issues:" -ForegroundColor Yellow
    Write-Host "1. PAT expired or invalid" -ForegroundColor White
    Write-Host "2. Missing required permissions (see README-Local.md)" -ForegroundColor White
    Write-Host "3. Organization name incorrect" -ForegroundColor White
}