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
$FileDate  = Get-Date -Format "yyyy-MM-dd"
$DataPath  = Join-Path $DataPath "$($Organization)" "$($FileDate)"
$Filename  = Join-Path $DataPath "GroupMemberships.csv"

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

# Get list of groups
Write-Verbose "Getting list of groups."
$allGroups = Get-AzDevOpsGroupList -Session $AzSession
Write-Verbose "Found $($allGroups.Count) groups."

$stepCounter = 0
foreach($group in $allGroups)
{
    Write-ProgressHelper -Message "Group Memberships" -Steps $allGroups.Count -StepNumber ($stepCounter++)

    Write-Verbose "Getting membership list for $($group.displayName)"
    $allMembers = Get-AzDevOpsGroupMembershipList `
        -Session $AzSession `
        -SubjectDescriptor $group.descriptor
    Write-Verbose "Found $($allMembers.Count) members."

    if ($allMembers.Count -gt 0)
    {
        $Memberships = @()

        foreach($member in $allMembers)
        {
            $Memberships += [PSCustomObject]@{
                groupId  = $member.containerDescriptor
                memberId = $member.memberDescriptor
            }
        }
    
        Write-Verbose "Writing to file. [$($Filename)]"
        $Memberships | Export-Csv -Path $Filename -NoTypeInformation -Append
    }
}

Write-Verbose "Removing AzDevOps session."
Remove-AzDevOpsSession $AzSession.Id