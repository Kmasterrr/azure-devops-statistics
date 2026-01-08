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
$Filename  = Join-Path $DataPath "GitCommits.csv"

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
Foreach ($project in $allProjects)
{
    Write-ProgressHelper -Message "Git Commits" -Steps $allProjects.Count -StepNumber ($stepCounter++)
    
    Write-Verbose "[*] $($project.name)"
    Write-Verbose "    Getting list of repositories."
    $allRepos = Get-AzDevOpsGitRepositoryList `
        -Session $AzSession `
        -Project "$($project.id)"
    Write-Verbose "    Found $($allRepos.Count) repositories."

    Foreach ($repo in $allRepos)
    {
        if ($repo.defaultBranch -and !$repo.isDisabled)
        {
            Write-Verbose "    Getting list of commits."
            $CommitResults = Get-AzDevOpsGitCommitList `
            -Session $AzSession `
            -Project "$($project.id)" `
            -RepositoryId "$($repo.id)" `
            -Top 9000
            # -SearchCriteria_FromDate "1/1/2022 00:00:00" `
            # -SearchCriteria_ToDate "6/1/2022 00:00:00" `
            Write-Verbose "    $($repo.name) [$($repo.defaultBranch)] - $($CommitResults.Count) commits."

            $GitCommits = @()

            Foreach ($commit in $CommitResults)
            {
                $GitCommits += [ordered]@{
                    commitId       = $commit.commitId
                    date           = $commit.author.date
                    author         = $commit.author.name
                    email          = $commit.author.email
                    repositoryId   = $repo.id
                    repositoryName = $repo.name
                    defaultBranch  = $repo.defaultBranch
                    projectId      = $repo.project.id
                    projectName    = $repo.project.name
                    comment        = $commit.comment
                    timestamp      = $Timestamp
                }
            }

            Write-Verbose "    Writing to file."
            $GitCommits | Export-Csv -Path $Filename -NoTypeInformation -Append
        }
    }
}

Write-Verbose "Removing AzDevOps session."
Remove-AzDevOpsSession $AzSession.Id