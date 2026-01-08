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

$BaseFilename = Join-Path $OutputPath "OrganizationStatistics"

$StatProperties = [ordered]@{
    Organization                = $Organization
    TimeStamp                   = $Timestamp
    Projects                    = 0
    BuildPipelines              = 0
    Builds                      = 0
    BuildsCompleted             = 0
    BuildCompletionPercentage   = 0
    ReleasePipelines            = 0
    Releases                    = 0
    ReleasesToProduction        = 0
    ReleasesCompleted           = 0
    ReleaseCompletionPercentage = 0
}

Write-Host "Collecting Organization Statistics..." -ForegroundColor Yellow

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

    # Get list of projects
    Write-Verbose "Getting list of projects."
    $ProjectsResult = Get-AzDevOpsProjectList -Session $AzSession
    # Use the direct array (not .value property)
    $projects = $ProjectsResult
    Write-Host "Found $($projects.Count) projects." -ForegroundColor Green

    $StatProperties['Projects'] = $projects.Count

    # Loop through all projects to collect stats
    $stepCounter = 0
    Foreach ($project in $projects) {
        Write-ProgressHelper -Message "Organization Stats" -Steps $projects.Count -StepNumber ($stepCounter++)
        
        Write-Host "Processing project: $($project.name)" -ForegroundColor Cyan
        
        try {
            Write-Verbose " |- Getting list of Build Pipelines."
            $BuildDefinitionsResult = Get-AzDevOpsBuildDefinitionList `
                -Session $AzSession `
                -Project "$($project.name)" `
                -Top 5000
            $StatProperties['BuildPipelines'] += $BuildDefinitionsResult.Count;

            Write-Verbose " |- Getting list of Builds."
            $BuildsResult = Get-AzDevOpsBuildList `
                -Session $AzSession `
                -Project "$($project.name)" `
                -Top 5000
            $StatProperties['Builds'] += $BuildsResult.Count;

            Foreach ($build in $BuildsResult) {
                switch ($build.status) {
                    completed { $StatProperties['BuildsCompleted'] += 1; }
                    Default {}
                }
            }

            Write-Verbose " |- Getting list of Release Pipelines."
            $ReleaseDefinitionsResult = Get-AzDevOpsReleaseDefinitionList `
                -Session $AzSession `
                -Project "$($project.name)" `
                -Top 5000
            $StatProperties['ReleasePipelines'] += $ReleaseDefinitionsResult.Count;

            Write-Verbose " |- Getting list of Releases."
            $ReleasesResult = Get-AzDevOpsReleaseList `
                -Session $AzSession `
                -Project "$($project.name)" `
                -Top 5000
            
            Foreach ($release in $ReleasesResult) {
                Write-Debug "|- ($($release.id)) $($release.name)"
                $releaseDetails = Get-AzDevOpsRelease `
                    -Session $AzSession `
                    -Project "$($project.name)" `
                    -ReleaseId "$($release.id)"
                
                Foreach ($e in $releaseDetails.environments) {
                    Write-Debug "    |-- Stage: $($e.name) - $($e.status)"
                    # A release could have several environments and each should count as a release.
                    # So we add them up in this loop to count up all the environment releases.
                    $StatProperties['Releases'] += 1;

                    # Total up completed releases
                    switch ($e.status) {
                        succeeded { $StatProperties['ReleasesCompleted'] += 1; }
                        partiallySucceeded { $StatProperties['ReleasesCompleted'] += 1; }
                        Default {}
                    }

                    # Total up releases to production when the environment/stage starts with prod*
                    if ($e.name.StartsWith('prod', "CurrentCultureIgnoreCase")) {
                        $StatProperties['ReleasesToProduction'] += 1;
                    }
                }
            }
        }
        catch {
            Write-Warning "Error processing project '$($project.name)': $($_.Exception.Message)"
        }
    }

    # Calculate completed builds percentage
    if ($StatProperties['Builds'] -ne 0) {
        $BuildCompletionPercentage = ($StatProperties['BuildsCompleted'] / $StatProperties['Builds']).ToString("P");
        $StatProperties['BuildCompletionPercentage'] = $BuildCompletionPercentage;    
    }

    # Calculate completed releases percentage
    if ($StatProperties['Releases'] -ne 0) {
        $ReleaseCompletionPercentage = ($StatProperties['ReleasesCompleted'] / $StatProperties['Releases']).ToString("P");
        $StatProperties['ReleaseCompletionPercentage'] = $ReleaseCompletionPercentage;
    }

    # Export data based on format preference
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $CsvPath = "$BaseFilename.csv"
        Write-Verbose "Writing to CSV file: $CsvPath"
        $StatProperties | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "Organization statistics saved to: $CsvPath" -ForegroundColor Green
    }

    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonPath = "$BaseFilename.json"
        Write-Verbose "Writing to JSON file: $JsonPath"
        
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $Timestamp
                dataType = "OrganizationStatistics"
            }
            data = $StatProperties
        }
        
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $JsonPath -Encoding UTF8
        Write-Host "Organization statistics saved to: $JsonPath" -ForegroundColor Green
    }

    Write-Host "Organization statistics collection completed." -ForegroundColor Cyan
    Write-Host "Projects: $($StatProperties['Projects']), Builds: $($StatProperties['Builds']), Releases: $($StatProperties['Releases'])" -ForegroundColor Yellow

}
catch {
    Write-Error "Error collecting organization statistics: $($_.Exception.Message)"
}
finally {
    # Clean up session
    if ($AzSession) {
        Write-Verbose "Removing AzDevOps session."
        Remove-AzDevOpsSession $AzSession.Id
    }
}