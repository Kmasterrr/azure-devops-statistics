# Azure DevOps Statistics - Local Environment Setup
# Sets up environment variables for local data collection without Azure Storage dependencies

[CmdletBinding()]
param()

Write-Host "🔧 Azure DevOps Statistics - Local Environment Setup" -ForegroundColor Cyan

# Get Organization Name
$Organization = Read-Host "Enter your Azure DevOps Organization name"
if (-not $Organization) {
    Write-Error "Organization name is required"
    exit 1
}

# Get Personal Access Token
$PAT = Read-Host "Enter your Personal Access Token" -AsSecureString
if (-not $PAT -or $PAT.Length -eq 0) {
    Write-Error "Personal Access Token is required"
    exit 1
}

# Convert secure string to plain text for environment variable
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PAT)
$PlainPAT = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Set environment variables for current session
$env:ADOS_ORGANIZATION = $Organization
$env:ADOS_PAT = $PlainPAT

Write-Host "✅ Environment variables set for local Azure DevOps data collection." -ForegroundColor Green
Write-Host "Organization: $Organization" -ForegroundColor Yellow
Write-Host ""
Write-Host "🚀 You can now run:" -ForegroundColor Cyan
Write-Host "  .\scripts\Collect-Statistics.ps1  - Collect all data" -ForegroundColor White
Write-Host "  .\scripts\Generate-Report.ps1     - Generate HTML report" -ForegroundColor White
Write-Host "  .\scripts\analysis\Show-TopContributors.ps1 - View top contributors" -ForegroundColor White

# Clean up sensitive data
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
Remove-Variable PlainPAT -ErrorAction SilentlyContinue