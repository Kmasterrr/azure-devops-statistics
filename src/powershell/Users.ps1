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
$Filename  = Join-Path $DataPath "Users.csv"

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

# Get list of users
Write-Verbose "Getting list of users."
$allUsers = Get-AzDevOpsUserEntitlementList -Session $AzSession
Write-Verbose "Found $($allUsers.Count) users."

$Users = @()
$stepCounter = 0
foreach($user in $allUsers)
{
    Write-ProgressHelper -Message "Users" -Steps $allUsers.Count -StepNumber ($stepCounter++)

    $Users += [PSCustomObject]@{
        id               = $user.id
        descriptor       = $user.user.descriptor
        principalName    = $user.user.principalName
        displayName      = $user.user.displayName
        email            = $user.user.mailAddress
        origin           = $user.user.origin
        originId         = $user.user.originId
        kind             = $user.user.subjectKind
        type             = $user.user.metaType
        domain           = $user.user.domain
        status           = $user.accessLevel.status
        license          = $user.accessLevel.licenseDisplayName
        licenseType      = $user.accessLevel.accountLicenseType
        source           = $user.accessLevel.assignmentSource
        dateCreated      = $user.dateCreated
        lastAccessedDate = $user.lastAccessedDate
        timeStamp        = $Timestamp
    }
}

Write-Verbose "Writing to file. [$($Filename)]"
$Users | Export-Csv -Path $Filename -NoTypeInformation -Append

Write-Verbose "Removing AzDevOps session."
Remove-AzDevOpsSession $AzSession.Id