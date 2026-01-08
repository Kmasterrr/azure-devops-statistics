[CmdletBinding()]
param 
(
    [string]$OutputPath = (Join-Path $PSScriptRoot "data"),
    [string]$Format = "Both" # Options: JSON, CSV, Both
)

<###[Environment Variables]#####################>
$Organization = $env:ADOS_ORGANIZATION
$OutputBasePath = if ($env:LOCAL_DATA_PATH) { $env:LOCAL_DATA_PATH } else { $OutputPath }

if (-not $Organization) {
    Write-Error "Environment variable ADOS_ORGANIZATION is not set. Please run Set-Environment-Local.ps1 first."
    exit 1
}

<###[Set Paths]#################################>
$ScriptPath = Join-Path $PSScriptRoot "src\powershell"
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

# Scripts that work with basic project permissions
$ScriptFiles = @(
    "ProjectStatistics-Local.ps1",
    "OrganizationStatistics-Local.ps1",
    "GitPullRequests-Local.ps1",
    "GitCommits-Local.ps1",
    "WorkItems-Local.ps1"
)

# Scripts that need Graph permissions (skip for now)
$SkippedScripts = @(
    "Users-Local.ps1", 
    "Groups-Local.ps1",
    "GroupMemberships-Local.ps1"
)

Write-Host "Collecting Azure DevOps Statistics from: $Organization" -ForegroundColor Cyan
Write-Host "Output Directory: $DataPath" -ForegroundColor Yellow
Write-Host "Output Format: $Format" -ForegroundColor Yellow
Write-Host ""
Write-Host "Note: Skipping user/group collection (requires Graph permissions)" -ForegroundColor Yellow
Write-Host "Collecting: Projects, Work Items, Git Data, Build/Release Stats" -ForegroundColor Green

$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$x = 0

foreach ($file in $ScriptFiles) {
    Write-Debug "Step Value: $x"
    Write-ProgressHelper -Message "Get Azure DevOps Stats" -ProcessId 1 -Steps $ScriptFiles.Count -StepNumber ($x++)

    $FilePath = Join-Path $ScriptPath $file
    
    if (Test-Path $FilePath) {
        Write-Host "Executing: $file" -ForegroundColor Green
        try {
            # Pass parameters to child scripts
            . $FilePath -OutputPath $DataPath -Format $Format
        }
        catch {
            Write-Warning "Error executing $file : $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning "Script file not found: $FilePath"
    }
}

$StopWatch.Stop()

# Create a summary report
$SummaryData = @{
    Organization = $Organization
    CollectionDate = $FileDate
    CollectionTime = $TimeStamp
    OutputPath = $DataPath
    Format = $Format
    Duration = $StopWatch.Elapsed.ToString()
    CollectedDataTypes = $ScriptFiles -replace "-Local.ps1", ""
    SkippedDataTypes = $SkippedScripts -replace "-Local.ps1", ""
    Note = "User and Group data skipped - requires Graph permissions in PAT"
    FilesGenerated = @()
}

# List all generated files
$GeneratedFiles = Get-ChildItem -Path $DataPath -File | Select-Object Name, Length, LastWriteTime
$SummaryData.FilesGenerated = $GeneratedFiles

# Save summary
$SummaryPath = Join-Path $DataPath "Collection-Summary.json"
$SummaryData | ConvertTo-Json -Depth 3 | Out-File -FilePath $SummaryPath -Encoding UTF8

Write-Host "`nData collection completed!" -ForegroundColor Green
Write-Host "Duration: $($StopWatch.Elapsed)" -ForegroundColor Yellow
Write-Host "Files generated: $($GeneratedFiles.Count)" -ForegroundColor Yellow
Write-Host "Summary saved to: $SummaryPath" -ForegroundColor Cyan

# Display file summary
Write-Host "`nGenerated Files:" -ForegroundColor White
$GeneratedFiles | Format-Table Name, @{Name="Size (KB)"; Expression={[math]::Round($_.Length/1KB, 2)}}, LastWriteTime -AutoSize

Write-Host "`nSkipped (need Graph permissions in PAT):" -ForegroundColor Yellow
$SkippedScripts | ForEach-Object { Write-Host "- $_" -ForegroundColor Gray }

Write-Host "`nTo collect user/group data:" -ForegroundColor Cyan
Write-Host "1. Update PAT with 'Graph' and 'Identity' permissions" -ForegroundColor White
Write-Host "2. Run: .\Collect-Statistics-Local.ps1" -ForegroundColor White