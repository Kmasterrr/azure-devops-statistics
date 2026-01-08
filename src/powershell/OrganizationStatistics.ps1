[CmdletBinding()]
Param
(
)

<###[Environment Variables]#####################>
$Organization = $env:ADOS_ORGANIZATION
$PAT          = $env:ADOS_PAT

<###[Set Paths]#################################>
$ModulePath = Join-Path $PSScriptRoot "\modules"
$DataPath   = Join-Path $PSScriptRoot "..\..\data"

<###[Load Modules]##############################>
Import-Module (Join-Path $ModulePath "ADOS") -Force
Import-Module (Join-Path $ModulePath "AzDevOps") -Force

<###[Script Variables]##########################>
$Timestamp    = Get-Date -Format "yyyy-MM-dd HH:mm K"
$FileDate     = Get-Date -Format "yyyy-MM-dd"
$DataPath     = Join-Path $DataPath "$($Organization)" "$($FileDate)"
$Filename     = Join-Path $DataPath "OrganizationStatistics.csv"

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

# Create data path if it does not exist.
Initialize-Path -Path $DataPath

# Remove file if it already exists.
if (Test-Path $Filename) {
    Write-Verbose "Removing existing data file: [$($Filename)]"
    Remove-Item $Filename
}

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
Write-Verbose "Found $($ProjectsResult.value.Count) projects."

$StatProperties['Projects'] = $ProjectsResult.value.Count

# Loop through all projects to collect stats
$stepCounter = 0
Foreach ($project in $ProjectsResult)
{
    Write-ProgressHelper -Message "Organization Stats" -Steps $ProjectsResult.Count -StepNumber ($stepCounter++)
    
    Write-Verbose "[*] $($project.name)"
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

    Foreach ($build in $BuildsResult)
    {
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
    
    Foreach ($release in $ReleasesResult)
    {
        Write-Debug "|- ($($release.id)) $($release.name)"
        $releaseDetails = Get-AzDevOpsRelease `
            -Session $AzSession `
            -Project "$($project.name)" `
            -ReleaseId "$($release.id)"
        
        Foreach ($e in $releaseDetails.environments)
        {
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

# Calculate completed builds percentage
if ($StatProperties['Builds'] -ne 0)
{
    $BuildCompletionPercentage = ($StatProperties['BuildsCompleted'] / $StatProperties['Builds']).ToString("P");
    $StatProperties['BuildCompletionPercentage'] = $BuildCompletionPercentage;    
}

# Calculate completed releases percentage
if ($StatProperties['Releases'] -ne 0)
{
    $ReleaseCompletionPercentage = ($StatProperties['ReleasesCompleted'] / $StatProperties['Releases']).ToString("P");
    $StatProperties['ReleaseCompletionPercentage'] = $ReleaseCompletionPercentage;
}

Write-Verbose "Writing to file."
$StatProperties | Export-Csv -Path $Filename -NoTypeInformation

# Destroy Azure DevOps session
Write-Verbose "Removing AzDevOps session."
Remove-AzDevOpsSession $AzSession.Id