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
$Filename  = Join-Path $DataPath "GitRepositoriesPermissions.csv"

$UriOrganization = "https://dev.azure.com/$($Organization)/"
$SecurityNameSpaceIdGitRepositories = "2e9eb7ed-3c0a-47d4-87c1-0ffdd275fd87"

# Create data path if it does not exist.
Initialize-Path -Path $DataPath

# Remove file if it already exists.
if (Test-Path $Filename) {
    Write-Verbose "Removing existing data file: [$($Filename)]"
    Remove-Item $Filename
}

# Login to Azure DevOps using PAT
Write-Verbose "Logging into Azure DevOps."
Write-Output $PAT | az devops login --org $UriOrganization

# Set the organization for all subsequent calls
Write-Verbose "Setting organization."
az devops configure --defaults organization=$UriOrganization

# Get project list
Write-Verbose "Getting project list."
$allProjects = az devops project list --org $UriOrganization --top 500 | ConvertFrom-Json
$allProjects = $allProjects.value

if ($ProjectName.Length -gt 0)
{
    Write-Debug "Filtering by Project: $($ProjectName)"
    $allProjects = $allProjects | Where-Object name -EQ $ProjectName
}

Write-Verbose "Found $($allProjects.Count) projects."
$stepCounter = 0
foreach ($project in $allProjects)
{
    Write-ProgressHelper -Message "$($project.name)" -Steps $allProjects.Count -StepNumber ($stepCounter++)

    Write-Verbose "[*] $($project.name)"
    
    $groups = az devops security group list --org $UriOrganization --project $project.id --only-show-errors | ConvertFrom-Json

    if (Test-Debug)
    {
        $groups | ConvertTo-Json -Depth 20 | Out-File -FilePath (Join-Path $DataPath "$($project.name)-groups-raw.json")
    }

    $groups = $groups.graphGroups
    
    $allrepos = az repos list --org $UriOrganization --project $project.id --only-show-errors | ConvertFrom-Json
    $r = 0

    Foreach ($ar in $allrepos)
    {
        Write-ProgressHelper -Message "$($ar.name)" -Steps $allrepos.Count -StepNumber ($r++)
        Write-Verbose " |- $($ar.name)"

        foreach ($aug in $groups)
        {
            Write-Verbose "   |- $($aug.principalName)"

            $groupPermissions = @()
            $gitToken = "repoV2/$($project.id)/$($ar.id)"
            $gitCommands = az devops security permission show --id $SecurityNameSpaceIdGitRepositories --subject $aug.descriptor --token $gitToken --org $UriOrganization --only-show-errors | ConvertFrom-Json
            $gitPermissions = ($gitCommands[0].acesDictionary | Get-Member -MemberType NoteProperty).Name

            if (Test-Debug)
            {
                Write-Debug "Writing raw response to file."
                $gitCommands | ConvertTo-Json -Depth 20 | Out-File -FilePath (Join-Path $DataPath "$($project.name)-$($ar.name)-$($aug.displayName)-gitcommands-raw.json")
            }

            foreach($gp in $gitCommands.acesDictionary.$gitPermissions.resolvedPermissions)
            {
                $groupPermissions += [PSCustomObject]@{
                    timestamp              = $TimeStamp
                    projectId              = $project.id
                    projectName            = $project.name
                    repoId                 = $ar.id
                    repoName               = $ar.name
                    securityNameSpaceId    = $SecurityNameSpaceIdGitRepositories
                    securityNameSpace      = 'Git Repositories'
                    groupDomain            = $aug.domain
                    groupDisplayName       = $aug.displayName
                    groupAccountName       = $aug.principalName
                    gitCommandName         = $gp.displayName.Replace("'",'')
                    gitCommandInternalName = $gp.name
                    gitCommandPermission   = $gp.effectivePermission
                }
            }

            Write-Debug "Writing to file."
            $groupPermissions | Export-Csv -Path $Filename -NoTypeInformation -Append
        }
    }
}