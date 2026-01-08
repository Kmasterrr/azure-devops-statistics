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
$Filename  = Join-Path $DataPath "ProjectLevelPermissions.csv"

$UriOrganization = "https://dev.azure.com/$($Organization)/"

$SecurityNameSpaceIds = @(
    [pscustomobject]@{SecurityNameSpace='Project';SecurityIdSpace='52d39943-cb85-4d7f-8fa8-c6baac873819'}
    [pscustomobject]@{SecurityNameSpace='Tagging';SecurityIdSpace='bb50f182-8e5e-40b8-bc21-e8752a1e7ae2'}
    [pscustomobject]@{SecurityNameSpace='AnalyticsViews';SecurityIdSpace='d34d3680-dfe5-4cc6-a949-7d9c68f73cba'}
    [pscustomobject]@{SecurityNameSpace='Analytics';SecurityIdSpace='58450c49-b02d-465a-ab12-59ae512d6531'}
)

$Commands = @(
    [pscustomobject]@{CommandType='General';CommandName='View project-level information'}
    [pscustomobject]@{CommandType='General';CommandName='Edit project-level information'}
    [pscustomobject]@{CommandType='General';CommandName='Delete team project'}
    [pscustomobject]@{CommandType='General';CommandName='Rename team project'}
    [pscustomobject]@{CommandType='General';CommandName='Manage project properties'}
    [pscustomobject]@{CommandType='General';CommandName='Suppress notifications for work item updates'}
    [pscustomobject]@{CommandType='General';CommandName='Update project visibility'}
    [pscustomobject]@{CommandType='Test Plans';CommandName='Create test runs'}
    [pscustomobject]@{CommandType='Test Plans';CommandName='Delete test runs'}
    [pscustomobject]@{CommandType='Test Plans';CommandName='View test runs'}
    [pscustomobject]@{CommandType='Test Plans';CommandName='Manage test environments'}
    [pscustomobject]@{CommandType='Test Plans';CommandName='Manage test configurations'}
    [pscustomobject]@{CommandType='Boards';CommandName='Delete and restore work items'}
    [pscustomobject]@{CommandType='Boards';CommandName='Move work items out of this project'}
    [pscustomobject]@{CommandType='Boards';CommandName='Permanently delete work items'}
    [pscustomobject]@{CommandType='Boards';CommandName='Bypass rules on work item updates'}
    [pscustomobject]@{CommandType='Boards';CommandName='Change process of team project.'}
    [pscustomobject]@{CommandType='Boards';CommandName='Create tag definition'}
    [pscustomobject]@{CommandType='Analytics';CommandName='Delete shared Analytics views'}
    [pscustomobject]@{CommandType='Analytics';CommandName='Edit shared Analytics views'}
    [pscustomobject]@{CommandType='Analytics';CommandName='View analytics'}
)

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

# Get list of users
Write-Verbose "Getting list of users."
$allUsers = az devops user list --org $UriOrganization --top 500 | ConvertFrom-Json
Write-Verbose "Found $($allUsers.members.Count) users."

Write-Verbose "Getting user group memberships."
$GroupMemberships = @()
$stepCounter = 0
foreach($u in $allUsers.members)
{
    Write-ProgressHelper -Message "Getting group memberships." -Steps $allUsers.members.Count -StepNumber ($stepCounter++)
    Write-Verbose "$($u.user.displayName) "

    $membershipList = az devops security group membership list --id $u.user.principalName --org $UriOrganization --relationship memberof | ConvertFrom-Json
    [array]$devopsGroups = ($membershipList | Get-Member -MemberType NoteProperty).Name

    foreach($g in $devopsGroups)
    {
        $GroupMemberships += @{
            domain            = $membershipList.$g.domain
            descriptor        = $membershipList.$g.descriptor
            groupDisplayName  = $membershipList.$g.displayName
            groupAccountName  = $membershipList.$g.principalName
            userPrincipalName = $u.user.principalName
            userDisplayName   = $u.user.displayName
        }
    }
}

# Get project list
Write-Verbose "Getting list of projects."
$allProjects = az devops project list --org $UriOrganization --top 500 | ConvertFrom-Json
$allProjects = $allProjects.value
Write-Verbose "Found $($allProjects.Count) projects."

if ($ProjectName.Length -gt 0)
{
    Write-Verbose "Filtering selected project."
    $allProjects = $allProjects | Where-Object name -EQ $ProjectName
}

Write-Verbose "Getting Project level permissions."
$stepCounter = 0
foreach ($project in $allProjects)
{
    Write-ProgressHelper -Message "Getting Project level permissions." -Steps $allProjects.Count -StepNumber ($stepCounter++)
    Write-Verbose "[*] $($project.name)"

    $Domain = "vstfs:///Classification/TeamProject/$($project.id)"
    $projectGroups = $GroupMemberships | Where-Object domain -EQ $Domain

    foreach($g in $projectGroups)
    {
        Write-Verbose " |- $($g.groupDisplayName) - $($g.userDisplayName)"

        foreach ($snsi in $SecurityNameSpaceIds)
        {
            switch ( $snsi.SecurityNameSpace )
            {
                'Project' { $Token = "`$PROJECT:vstfs:///Classification/TeamProject/$($project.id)" }
                'Tagging' { $Token = "/$($project.id)" }
                'AnalyticsViews' { $Token = "`$/Shared/$($project.id)" }
                'Analytics' { $Token = "`$/$($project.id)" }
            }

            $projectCommands = az devops security permission show --id $snsi.SecurityIdSpace --subject $g.descriptor --token $Token --org $UriOrganization | ConvertFrom-Json
            $projectPermissions = ($projectCommands[0].acesDictionary | Get-Member -MemberType NoteProperty).Name
            $ProjectLevelPermissions = @()

            foreach($pp in $projectCommands.acesDictionary.$projectPermissions.resolvedPermissions)
            {
                $validCommand =  $Commands | Where-Object CommandName -EQ $pp.displayName
                if ($validCommand)
                {
                    $ProjectLevelPermissions += [ordered]@{
                        timestamp                       = $timeStamp
                        teamProjectName                 = $project.name
                        securityNameSpace               = $snsi.SecurityNameSpace
                        userPrincipalName               = $g.userPrincipalName
                        userDisplayName                 = $g.userDisplayName
                        groupDisplayName                = $g.groupDisplayName
                        groupAccountName                = $g.groupAccountName
                        projectLevelType                = $validCommand.CommandType
                        projectLevelCommandName         = $pp.displayName.Replace("'",'')
                        projectLevelCommandInternalName = $pp.name
                        projectLevelCommandPermission   = $pp.effectivePermission
                    }
                }
            }

            Write-Debug "    [$($snsi.SecurityNameSpace)] Writing to file."
            $ProjectLevelPermissions | Export-Csv -Path $Filename -NoTypeInformation -Append
        }
    }
}