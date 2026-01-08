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
$Filename  = Join-Path $DataPath "ProjectGroupsAndUsers.csv"

$UriOrganization = "https://dev.azure.com/$($Organization)/"

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
foreach($project in $allProjects)
{
    Write-ProgressHelper -Message "$($project.name)" -Steps $allProjects.Count -StepNumber ($stepCounter++)
    Write-Verbose "[*] $($project.name)"
    
    $groups = az devops security group list --org $UriOrganization --project $project.id --only-show-errors | ConvertFrom-Json
    $groups = $groups.graphGroups

    foreach ($g in $groups)
    {
        Write-Verbose " |- $($g.principalName)"

        $members = az devops security group membership list --id $g.descriptor --relationship members | ConvertFrom-Json
        [array]$mArray = ($members | Get-Member -MemberType NoteProperty).Name

        $ProjectsGroupsAndUsers = @()

        foreach($m in $mArray)
        {
            Write-Verbose "   |- [$($members.$m.subjectKind)] - $($members.$m.displayName)"

            $ProjectsGroupsAndUsers += [ordered]@{
                timeStamp     = $TimeStamp
                projectName   = $project.name
                groupName     = $g.principalName
                principalName = $members.$m.principalName
                displayName   = $members.$m.displayName
                origin        = $members.$m.origin
                type          = $members.$m.subjectKind
            }
        }

        Write-Verbose "    Writing to file."
        $ProjectsGroupsAndUsers | Export-Csv -Path $Filename -NoTypeInformation -Append
    }
}
