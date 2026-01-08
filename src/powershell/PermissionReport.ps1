[CmdletBinding()]
param
(
    [string]$ProjectName = "",
    [string]$RepoName    = ''  # Ability to narrow to specific list or regex. Update Line 20 for list with -in
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
$FileName     = Join-Path $DataPath "PermissionsReport.csv"
$ReportDate   = Get-Date -Format "yyyyMMddhhmm"
$PollingDelay = 5
$PollingMax   = 6

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
$projectList = Get-AzDevOpsProjectList -Session $AzSession

if ($ProjectName.Length -gt 0)
{
    Write-Verbose "Filtering project list by name: [$($ProjectName)]."
    $projectList = $projectList | Where-Object name -EQ $ProjectName
}

Write-Verbose "Found $($projectList.Count) projects."

$stepCounter = 0

foreach ($project in $projectList)
{
    Write-ProgressHelper -Message "Progress" -Steps $projectList.Count -StepNumber ($stepCounter++)

    Write-Verbose "[*] $($project.name)"
    Write-Verbose "    Getting list of repositories."
    $repoList = Get-AzDevOpsGitRepositoryList `
        -Session $AzSession `
        -Project "$($project.id)"

    # Checking if RepoName was set in the params
    if ($RepoName)
    {
        Write-Verbose "    Filtering repository list by name: [$($RepoName)]."
        $repoList = $repoList | Where-Object name -EQ $RepoName 
    }

    Write-Verbose "    Found $($repoList.Count) repositories."

    Foreach ($repo in $repoList)
    {
        if (!$repo.isDisabled)
        {
            Write-Verbose "    Currently Working on $($Repo.name)"
        
            # Create an object for create action and for reuse in script
            $requestBody = [pscustomobject]@{
                descriptors = @() 
                reportname  = "$($repo.name)-$($ReportDate)"
                resources   = @(
                    @{
                        resourceId   = $repo.id
                        resourceName = $repo.name
                        resourceType = 'repo'
                    }
                )  
            }
            
            # Creating report here and storing in var for less noise
            New-AzDevOpsPermissionReport -Session $AzSession -InputObject $requestBody
            
            # Set CheckReport to null so that the while loop works
            $CheckReport = $null
            
            # Checking to see if the report is ready or not
            Write-Verbose "    Waiting for report to return status: completedSuccessfully"
            
            $CurrentPoll = 0
            while (($CheckReport.reportStatus -notmatch "completedSuccessfully") -or ($CurrentPoll -lt $PollingMax))
            {
                # Write-Host '#' -NoNewline
                Start-Sleep -Seconds $PollingDelay
                $CurrentPoll++
                $CheckReport = Get-AzDevOpsPermissionReportList -Session $AzSession | Where-Object { $_.reportname -Match $requestBody.reportname }
            }
            
            # Report is ready for Download at this point so the Var report is storing that report
            Write-Verbose "    Report is ready for download. Starting Download now."
            $report = Save-AzDevOpsPermissionReport -Session $AzSession -ReportId $CheckReport.id -PassThru -OutputPath (Join-Path $DataPath "PermissionsReport-$($requestBody.reportname)-raw.json")
    
            if (Test-Debug) 
            {
                $report | ConvertTo-Json -Depth 20 | Out-File -FilePath (Join-Path $DataPath "PermissionsReport-$($requestBody.reportname).json")
            }
            
            $repoPermissions = @()
    
            foreach($i in $report | Where-Object { ($_.Descriptor -Match 'vssgp.') -or ($_.Descriptor -Match 'aadgp.') })
            {
                foreach($p in $i.Permissions | Where-Object EffectivePermission -Match 'Allow')
                {
                    $repoPermissions += [PSCustomObject]@{
                        Timestamp           = $Timestamp
                        ProjectId           = $project.id
                        ProjectName         = $project.name
                        RepositoryId        = $repo.id
                        RepositoryName      = $repo.name
                        Id                  = $i.Id
                        Descriptor          = $i.Descriptor
                        PrincipalName       = $i.AccountName
                        DisplayName         = $i.DisplayName
                        PermissionName      = $p.PermissionName
                        EffectivePermission = $p.EffectivePermission
                    }
                }
            }
    
            Write-Verbose "    Writing to file."
            $repoPermissions | Export-Csv -Path $Filename -NoTypeInformation -Append
        }
    }
}

Write-Verbose "Removing AzDevOps session."
Remove-AzDevOpsSession $AzSession.Id