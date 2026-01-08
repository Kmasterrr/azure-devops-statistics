[CmdletBinding()]
Param
(
    [string]$ProjectName
)

<###[Environment Variables]#####################>
$Organization = $env:ADOS_ORGANIZATION
$PAT          = $env:ADOS_PAT

<###[Set Paths]#################################>
$ModulePath = Join-Path $PSScriptRoot "\modules"
$DataPath   = Join-Path $PSScriptRoot "..\..\data"

<###[Load Modules]##############################>
Import-Module (Join-Path $ModulePath "ADOS") -Force

<###[Script Variables]##########################>
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm K"
$FileDate  = Get-Date -Format "yyyy-MM-dd"
$DataPath  = Join-Path $DataPath "$($Organization)" "$($FileDate)"
$Filename  = Join-Path $DataPath "ProjectStatistics.csv"

$AzureDevOpsAuthenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PAT)")) } + @{"Content-Type"="application/json"; "Accept"="application/json"}
$UriOrganization                = "https://dev.azure.com/$($Organization)/"
$UriOrganizationRM              = "https://vsrm.dev.azure.com/$($Organization)/"
$monthAgo                       = (Get-Date).AddMonths(-1).ToString("yyyy-MM-dd")

# Create data path if it does not exist.
Initialize-Path -Path $DataPath

# Remove file if it already exists.
if (Test-Path $Filename) {
    Write-Verbose "Removing existing data file: [$($Filename)]"
    Remove-Item $Filename
}

# Get list of projects
Write-Verbose "Getting list of projects."
$uriProject = $UriOrganization + "_apis/projects?`$top=500&api-version=6.1-preview.4"
$ProjectsResult = Invoke-RestMethod -Uri $uriProject -Method Get -Headers $AzureDevOpsAuthenicationHeader

if ($ProjectName.Length -gt 0)
{
    Write-Verbose "Filtering project list by name: [$($ProjectName)]."
    $ProjectsResult = $ProjectsResult | Where-Object name -EQ $ProjectName
}

Write-Verbose "Found $($ProjectsResult.Count) projects."

$ProjectStats  = @()

# Loop through all projects to collect stats
$stepCounter = 0
Foreach ($project in $ProjectsResult.value)
{
    Write-ProgressHelper -Message "$($project.name)" -Steps $ProjectsResult.Count -StepNumber ($stepCounter++)

    Write-Verbose "[*] $($project.name)"

    $uriProjectStats = $UriOrganization + "_apis/Contribution/HierarchyQuery/project/$($project.id)?api-version=6.1-preview.1"   
    $projectStatsBody = @{
        "contributionIds"= @(
            "ms.vss-work-web.work-item-metrics-data-provider-verticals", 
            "ms.vss-code-web.code-metrics-data-provider-verticals", 
            "ms.vss-code-web.build-metrics-data-provider-verticals")
        "dataProviderContext" = @{
            "properties" =@{
                "numOfDays"=30
                "sourcePage"=@{
                    "url"=($UriOrganization + $project.name)
                    "routeId"="ms.vss-tfs-web.project-overview-route"
                    "routeValues" =@{
                        "project" = $project.id
                        "controller"="Apps"
                        "action"="ContributedHub"
                        "serviceHost"=$Organization
                        }
                    }          
                }
            }
        }  | ConvertTo-Json -Depth 5

    $projectStatsResult = Invoke-WebRequest -Uri $uriProjectStats -Headers $AzureDevOpsAuthenicationHeader -Method Post -Body $projectStatsBody 
    $projectStatsJson = ConvertFrom-Json $projectStatsResult.Content

    if (Test-Debug)
    {
        $projectStatsResult.Content | Out-File -FilePath (Join-Path $DataPath "ProjectStats-$($project.name)-raw.json")
    }

    $workItemsCreated      = 0
    $workItemsCompleted    = 0
    $commitsPushed         = 0
    $pullRequestsCreated   = 0
    $pullRequestsCompleted = 0

    $workItemsCreated = $projectStatsJson.dataProviders.'ms.vss-work-web.work-item-metrics-data-provider-verticals'.workMetrics.workItemsCreated
    $workItemsCompleted = $projectStatsJson.dataProviders.'ms.vss-work-web.work-item-metrics-data-provider-verticals'.workMetrics.workItemsCompleted
    $commitsPushed = $projectStatsJson.dataProviders.'ms.vss-code-web.code-metrics-data-provider-verticals'.gitmetrics.commitsPushedCount
    if (!$commitsPushed) { $commitsPushed = 0 }

    $pullRequestsCreated = $projectStatsJson.dataProviders.'ms.vss-code-web.code-metrics-data-provider-verticals'.gitmetrics.pullRequestsCreatedCount
    if (!$pullRequestsCreated) { $pullRequestsCreated = 0 }

    $pullRequestsCompleted = $projectStatsJson.dataProviders.'ms.vss-code-web.code-metrics-data-provider-verticals'.gitmetrics.pullRequestsCompletedCount
    if (!$pullRequestsCompleted) { $pullRequestsCompleted = 0 }
           
    $uriBuildMetrics = $UriOrganization + "$($project.id)/_apis/build/Metrics/Daily?minMetricsTime=$($monthAgo)" 
    $buildMetricsResult = Invoke-RestMethod -Uri $uriBuildMetrics -Method get -Headers $AzureDevOpsAuthenicationHeader
    $totalBuilds = 0
    $buildMetricsResult.value | Where-Object { $_.name -eq 'TotalBuilds'} | ForEach-Object { $totalBuilds+= $_.intValue }

    $totalReleases = 0
    $UriReleaseMetrics = $UriOrganizationRM + "$($project.id)/_apis/Release/metrics?minMetricsTime=minMetricsTime=$($monthAgo)"
    $releaseMetricsResult = Invoke-RestMethod -Uri $UriReleaseMetrics -Method get -Headers $AzureDevOpsAuthenicationHeader
    $releaseMetricsResult.value | ForEach-Object { $totalReleases+= $_.value }

    $ProjectStats += [ordered]@{
        timeStamp             = $TimeStamp
        projectId             = $project.id
        projectName           = $project.name
        workItemsCreated      = $workItemsCreated
        workItemsCompleted    = $workItemsCompleted
        commitsPushed         = $commitsPushed
        pullRequestsCreated   = $pullRequestsCreated
        pullRequestsCompleted = $pullRequestsCompleted
        builds                = $totalBuilds
        releases              = $totalReleases
    }
}

Write-Verbose "Writing to file."
$ProjectStats | Export-Csv -Path $Filename -NoTypeInformation -Append