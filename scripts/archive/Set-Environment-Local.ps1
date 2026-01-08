# Azure DevOps Configuration for Local Data Collection
# This simplified version only requires Azure DevOps credentials
# No Azure Storage Account or SQL Database needed

# Azure DevOps Organization and Personal Access Token
$env:ADOS_ORGANIZATION = 'C4AI-ANZ'
$env:ADOS_PAT = '79SRaqcvxZU3uJsPoXtrKXJXruw1wICukKUvna2SdOAxgfYzU1FKJQQJ99BKACAAAAAkGEQJAAASAZDObD87'

# Optional: Set custom output directory (defaults to .\data)
# $env:LOCAL_DATA_PATH = 'C:\path\to\your\data\folder'

Write-Host "Environment variables set for local Azure DevOps data collection." -ForegroundColor Green
Write-Host "Organization: $env:ADOS_ORGANIZATION" -ForegroundColor Yellow
Write-Host "Make sure to replace 'your-organization-name' and 'your-personal-access-token' with actual values." -ForegroundColor Cyan