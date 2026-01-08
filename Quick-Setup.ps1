# Quick Setup for Azure DevOps Statistics Collection
# This script sets up your environment variables

Write-Host "Azure DevOps Statistics - Quick Setup" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Get Organization Name
$Organization = Read-Host "Enter your Azure DevOps Organization name"

# Get PAT
$PAT = Read-Host "Enter your Personal Access Token (PAT)" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PAT)
$PlainPAT = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Set environment variables
$env:ADOS_ORGANIZATION = $Organization
$env:ADOS_PAT = $PlainPAT

Write-Host ""
Write-Host "Environment variables set successfully!" -ForegroundColor Green
Write-Host "Organization: $Organization" -ForegroundColor Yellow
Write-Host ""
Write-Host "You can now run:" -ForegroundColor Cyan
Write-Host "  .\Collect-Statistics-Fixed.ps1   - Collect all data" -ForegroundColor White
Write-Host "  .\Generate-Report-Fixed.ps1      - Generate HTML report" -ForegroundColor White
Write-Host ""
