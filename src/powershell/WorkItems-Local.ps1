[CmdletBinding()]
Param
(
    [string]$OutputPath = ".\data",
    [string]$Format = "Both", # Options: JSON, CSV, Both
    [string]$ProjectName = ""
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
$TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm K"
$batchSize = 200

# Ensure output path exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$BaseFilename = Join-Path $OutputPath "WorkItems"

Write-Host "Collecting Work Items data..." -ForegroundColor Yellow

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

    # Get project list
    Write-Verbose "Getting project list."
    $allProjects = Get-AzDevOpsProjectList -Session $AzSession

    if ($ProjectName.Length -gt 0) {
        Write-Verbose "Filtering project list by name: [$($ProjectName)]."
        # Use the direct array (not .value property)
        $projects = $allProjects | Where-Object name -EQ $ProjectName
    } else {
        # Use the direct array (not .value property)
        $projects = $allProjects
    }

    Write-Host "Found $($projects.Count) projects to process." -ForegroundColor Green

    $allWorkItems = @()
    $stepCounter = 0
    
    foreach ($project in $projects) {
        Write-ProgressHelper -Message "Processing Work Items" -Steps $projects.Count -StepNumber ($stepCounter++)
        
        Write-Host "Processing project: $($project.name)" -ForegroundColor Cyan
        Write-Verbose "Getting list of work items."
        
        try {
            # Updated WIQL query for better performance and broader coverage
            $wiqlQuery = @"
Select [System.Id] From WorkItems 
Where [System.TeamProject] = @project 
AND [System.WorkItemType] IN ('User Story', 'Task', 'Bug', 'Feature', 'Epic') 
AND [State] <> 'Removed' 
AND [System.CreatedDate] >= '2020-01-01'
"@
            
            $wiqlResults = Invoke-AzDevOpsQuery `
                -Session $AzSession `
                -Project $project.id `
                -Query $wiqlQuery
            Write-Verbose "Found $($wiqlResults.WorkItems.Count) query results."

            if ($wiqlResults.WorkItems.Count -gt 0) {
                $wiList = New-Object System.Collections.ArrayList(, $wiqlResults.WorkItems)
                
                Write-Verbose "Generating batches."
                $batches = Get-BatchRanges `
                    -TotalItems $wiList.Count `
                    -BatchSize $batchSize `
                    -ZeroIndex
                Write-Verbose "Batches: $($batches.Count), Batch Size: $($batchSize)"

                foreach ($b in $batches) {
                    Write-Verbose "Batch[$($b.idx)] - range: $($b.range[0]), $($b.range[1])"

                    $batch = $wiList.GetRange($b.range[0], $b.range[1])
                    $ids = $batch.id

                    Write-Verbose "Getting Work Item details."
                    $wiBatch = Get-AzDevOpsWorkItemList `
                        -Session $AzSession `
                        -Project $project.id `
                        -Ids $ids
                    Write-Verbose "Found $($wiBatch.Count) work item details."

                    foreach ($wi in $wiBatch) {
                        $allWorkItems += [PSCustomObject]@{
                            timeStamp        = $TimeStamp
                            project          = $project.name
                            projectId        = $project.id
                            workItemId       = $wi.id
                            areaPath         = $wi.fields."System.AreaPath"
                            iterationPath    = $wi.fields."System.IterationPath"
                            valueArea        = $wi.fields."Microsoft.VSTS.Common.ValueArea"
                            workItemType     = $wi.fields."System.WorkItemType"
                            title            = $wi.fields."System.Title"
                            priority         = $wi.fields."Microsoft.VSTS.Common.Priority"
                            state            = $wi.fields."System.State"
                            reason           = $wi.fields."System.Reason"
                            storyPoints      = $wi.fields."Microsoft.VSTS.Scheduling.StoryPoints"
                            originalEstimate = $wi.fields."Microsoft.VSTS.Scheduling.OriginalEstimate"
                            completedWork    = $wi.fields."Microsoft.VSTS.Scheduling.CompletedWork"
                            assignedTo       = if ($wi.fields."System.AssignedTo") { $wi.fields."System.AssignedTo".id } else { $null }
                            createdBy        = if ($wi.fields."System.CreatedBy") { $wi.fields."System.CreatedBy".id } else { $null }
                            createdDate      = $wi.fields."System.CreatedDate"
                            changedBy        = if ($wi.fields."System.ChangedBy") { $wi.fields."System.ChangedBy".id } else { $null }
                            changedDate      = $wi.fields."System.ChangedDate"
                            closedBy         = if ($wi.fields."Microsoft.VSTS.Common.ClosedBy") { $wi.fields."Microsoft.VSTS.Common.ClosedBy".id } else { $null }
                            closedDate       = $wi.fields."Microsoft.VSTS.Common.ClosedDate"
                            resolvedBy       = if ($wi.fields."Microsoft.VSTS.Common.ResolvedBy") { $wi.fields."Microsoft.VSTS.Common.ResolvedBy".id } else { $null }
                            resolvedDate     = $wi.fields."Microsoft.VSTS.Common.ResolvedDate"
                            stateChangeDate  = $wi.fields."Microsoft.VSTS.Common.StateChangeDate"
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Error processing work items for project '$($project.name)': $($_.Exception.Message)"
        }
    }

    # Export data based on format preference
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $CsvPath = "$BaseFilename.csv"
        Write-Verbose "Writing to CSV file: $CsvPath"
        $allWorkItems | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "Work Items data saved to: $CsvPath" -ForegroundColor Green
    }

    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonPath = "$BaseFilename.json"
        Write-Verbose "Writing to JSON file: $JsonPath"
        
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $TimeStamp
                recordCount = $allWorkItems.Count
                projectsProcessed = $allProjects.Count
                dataType = "WorkItems"
            }
            data = $allWorkItems
        }
        
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $JsonPath -Encoding UTF8
        Write-Host "Work Items data saved to: $JsonPath" -ForegroundColor Green
    }

    Write-Host "Work Items collection completed. Total work items: $($allWorkItems.Count)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error collecting work items data: $($_.Exception.Message)"
}
finally {
    # Clean up session
    if ($AzSession) {
        Write-Verbose "Removing AzDevOps session."
        Remove-AzDevOpsSession $AzSession.Id
    }
}