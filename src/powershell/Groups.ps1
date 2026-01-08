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
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm K"
$FileDate  = Get-Date -Format "yyyy-MM-dd"
$DataPath  = Join-Path $DataPath "$($Organization)" "$($FileDate)"
$Filename  = Join-Path $DataPath "Groups.csv"

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

$Groups = @()
$stepCounter = 0
foreach($group in $allGroups)
{
    Write-ProgressHelper -Message "Groups" -Steps $allGroups.Count -StepNumber ($stepCounter++)

    $Groups += [PSCustomObject]@{
        id            = $group.descriptor
        principalName = $group.principalName
        displayName   = $group.displayName
        description   = $group.description.replace("`n", "").replace("`r", "")
        origin        = $group.origin
        originId      = $group.originId
        domain        = $group.domain
        subjectKind   = $group.subjectKind
        timeStamp     = $Timestamp
    }
}

Write-Verbose "Writing to file. [$($Filename)]"
$Groups | Export-Csv -Path $Filename -NoTypeInformation -Append

Write-Verbose "Removing AzDevOps session."
Remove-AzDevOpsSession $AzSession.Id