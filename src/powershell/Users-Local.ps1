[CmdletBinding()]
Param
(
    [string]$OutputPath = ".\data",
    [string]$Format = "Both" # Options: JSON, CSV, Both
)

<###[Environment Variables]#####################>
$Organization = $env:ADOS_ORGANIZATION
$PAT          = $env:ADOS_PAT

if (-not $Organization -or -not $PAT) {
    Write-Error "Environment variables ADOS_ORGANIZATION and ADOS_PAT must be set. Please run Set-Environment-Local.ps1 first."
    return
}

<###[Set Paths]#################################>
$ModulePath = Join-Path $PSScriptRoot "modules"

<###[Load Modules]##############################>
Import-Module (Join-Path $ModulePath "ADOS") -Force
Import-Module (Join-Path $ModulePath "AzDevOps") -Force

<###[Script Variables]##########################>
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm K"

# Ensure output path exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$BaseFilename = Join-Path $OutputPath "Users"

Write-Host "Collecting Users data..." -ForegroundColor Yellow

try {
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
    Write-Host "Found $($allUsers.Count) users." -ForegroundColor Green

    $Users = @()
    $stepCounter = 0
    foreach($user in $allUsers) {
        Write-ProgressHelper -Message "Processing Users" -Steps $allUsers.Count -StepNumber ($stepCounter++)

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

    # Export data based on format preference
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $CsvPath = "$BaseFilename.csv"
        Write-Verbose "Writing to CSV file: $CsvPath"
        $Users | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "Users data saved to: $CsvPath" -ForegroundColor Green
    }

    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonPath = "$BaseFilename.json"
        Write-Verbose "Writing to JSON file: $JsonPath"
        
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $Timestamp
                recordCount = $Users.Count
                dataType = "Users"
            }
            data = $Users
        }
        
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $JsonPath -Encoding UTF8
        Write-Host "Users data saved to: $JsonPath" -ForegroundColor Green
    }

    Write-Host "Users collection completed. Total users: $($Users.Count)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error collecting users data: $($_.Exception.Message)"
}
finally {
    # Clean up session
    if ($AzSession) {
        Write-Verbose "Removing AzDevOps session."
        Remove-AzDevOpsSession $AzSession.Id
    }
}