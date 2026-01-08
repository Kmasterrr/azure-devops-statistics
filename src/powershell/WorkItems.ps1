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
Import-Module (Join-Path $ModulePath "AzDevOps") -Force

<###[Script Variables]##########################>
$TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm K"
$FileDate  = Get-Date -Format "yyyy-MM-dd"
$DataPath  = Join-Path $DataPath "$($Organization)" "$($FileDate)"
$Filename  = Join-Path $DataPath "WorkItems.csv"
$batchSize = 200


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

# Get project list
Write-Verbose "Getting project list."
$allProjects = Get-AzDevOpsProjectList -Session $AzSession

if ($ProjectName.Length -gt 0)
{
    Write-Verbose "Filtering project list by name: [$($ProjectName)]."
    $allProjects = $allProjects | Where-Object name -EQ $ProjectName
}

Write-Verbose "Found $($allProjects.Count) projects."
$stepCounter = 0
foreach ($project in $allProjects)
{
    Write-ProgressHelper -Message "Progress" -Steps $allProjects.Count -StepNumber ($stepCounter++)
    
    Write-Verbose "[*] $($project.name)"
    Write-Verbose "    Gettting list of work items."
    $wiqlResults = Invoke-AzDevOpsQuery `
        -Session $AzSession `
        -Project $project.id `
        -Query "Select [System.Id] From WorkItems Where [System.TeamProject] = @project AND [System.WorkItemType] IN ('User Story', 'Task', 'Bug') AND [State] <> 'Removed' AND [System.CreatedDate] >= '01-01-2020'"
    Write-Verbose "    Found $($wiqlResults.WorkItems.Count) query results."

    if (Test-Debug) 
    {
        $DebugFile = (Join-Path $DataPath "wiql-$($project.name)-raw.json")
        Write-Debug "Writing debug file: [$($DebugFile)]"
        $wiqlResults | ConvertTo-Json -Depth 20 | Out-File -FilePath $DebugFile
    }

    if ($wiqlResults.WorkItems.Count -gt 0) 
    {
        $wiList = New-Object System.Collections.ArrayList(, $wiqlResults.WorkItems)
        
        Write-Verbose "    Generating batches."
        $batches = Get-BatchRanges `
        -TotalItems $wiList.Count `
        -BatchSize $batchSize `
        -ZeroIndex
        Write-Verbose "    Batches: $($batches.Count), Batch Size: $($batchSize)"

        foreach ($b in $batches)
        {
            Write-Verbose "    Batch[$($b.idx)] - range: $($b.range[0]), $($b.range[1])"

            $batch = $wiList.GetRange($b.range[0], $b.range[1])
            $ids = $batch.id

            Write-Verbose "    Getting Work Item details."
            $wiBatch = Get-AzDevOpsWorkItemList `
                -Session $AzSession `
                -Project $project.id `
                -Ids $ids
            Write-Verbose "    Found $($wiBatch.Count) work item details."

            if (Test-Debug) 
            {
                $DebugFile = (Join-Path $DataPath "wiql-$($project.name)-batch$($b.idx)-raw.json")
                Write-Debug "Writting debug file: [$($DebugFile)]"
                $wiBatch | ConvertTo-Json -Depth 20 | Out-File -FilePath $DebugFile
            }

            $workItems = @()

            foreach ($wi in $wiBatch)
            {
                $workItems += [PSCustomObject]@{
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
                    assignedTo       = $wi.fields."System.AssignedTo".id
                    createdBy        = $wi.fields."System.CreatedBy".id
                    createdDate      = $wi.fields."System.CreatedDate"
                    changedBy        = $wi.fields."System.ChangedBy".id
                    changedDate      = $wi.fields."System.ChangedDate"
                    closedBy         = $wi.fields."Microsoft.VSTS.Common.ClosedBy".id
                    ClosedDate       = $wi.fields."Microsoft.VSTS.Common.ClosedDate"
                    resolvedBy       = $wi.fields."Microsoft.VSTS.Common.ResolvedBy".id
                    resolvedDate     = $wi.fields."Microsoft.VSTS.Common.ResolvedDate"
                    stateChangeDate  = $wi.fields."Microsoft.VSTS.Common.StateChangeDate"
                }
            }

            Write-Debug "Writing to output file: [$($Filename)]"
            $workItems | Export-Csv -Path $Filename -NoTypeInformation -Append
        }
    }
}


Write-Verbose "Removing AzDevOps session."
Remove-AzDevOpsSession $AzSession.Id