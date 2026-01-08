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
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm K"
$FileDate  = Get-Date -Format "yyyy-MM-dd"
$DataPath  = Join-Path $DataPath "$($Organization)" "$($FileDate)"
$Filename  = Join-Path $DataPath "GitPullRequests.csv"

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
Write-Verbose "Getting list of projects."
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
    Write-ProgressHelper -Message "Git Pull Requests" -Steps $allProjects.Count -StepNumber ($stepCounter++)

    Write-Verbose "[*] $($project.name)"
    Write-Verbose "    Getting list of repositories."
    $allRepos = Get-AzDevOpsGitRepositoryList `
        -Session $AzSession `
        -Project "$($project.id)"
    Write-Verbose "    Found $($allRepos.Count) repositories."

    foreach ($repo in $allRepos)
    {
        if ($repo.defaultBranch -and !$repo.isDisabled)
        {
            Write-Verbose "    Getting list of pull requests."
            $allPRs = Get-AzDevOpsGitPullRequestList `
                -Session $AzSession `
                -Project "$($project.id)" `
                -RepositoryId "$($repo.id)" `
                -SearchCriteria_TargetRefName "$($repo.defaultBranch)" `
                -SearchCriteria_Status "all" `
                -Top 5000
            Write-Verbose "    $($repo.name) [$($repo.defaultBranch)] - $($allPRs.Count) pull requests."

            if ($allPRs.Count -gt 0)
            {
                $GitPullRequests = @()

                foreach ($pr in $allPRs)
                {
                    $GitPullRequests += [ordered]@{
                        pullRequestId  = $pr.pullRequestId
                        status         = $pr.status
                        title          = $pr.title
                        sourceBranch   = $pr.sourceRefName
                        targetBranch   = $pr.targetRefName
                        projectId      = $repo.project.id
                        projectName    = $repo.project.name
                        repositoryId   = $repo.id
                        repositoryName = $repo.name
                        userId         = $pr.createdBy.id
                        user           = $pr.createdBy.displayName
                        creationDate   = $pr.creationDate
                        closedDate     = ($null -eq $pr.closedDate) ? '' : $pr.closedDate
                        timestamp      = $Timestamp
                    }
                }

                Write-Verbose "    Writing to file."
                $GitPullRequests | Export-Csv -Path $Filename -NoTypeInformation -Append
            }
        }
    }
}

Write-Verbose "Removing AzDevOps session."
Remove-AzDevOpsSession $AzSession.Id