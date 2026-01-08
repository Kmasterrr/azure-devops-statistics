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

$BaseFilename = Join-Path $OutputPath "Groups"

Write-Host "Collecting Groups data..." -ForegroundColor Yellow

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

    # Get list of groups
    Write-Verbose "Getting list of groups."
    $allGroups = Get-AzDevOpsGroupList -Session $AzSession
    Write-Host "Found $($allGroups.Count) groups." -ForegroundColor Green

    $Groups = @()
    $stepCounter = 0
    foreach($group in $allGroups) {
        Write-ProgressHelper -Message "Processing Groups" -Steps $allGroups.Count -StepNumber ($stepCounter++)

        $Groups += [PSCustomObject]@{
            id            = $group.descriptor
            principalName = $group.principalName
            displayName   = $group.displayName
            description   = if ($group.description) { $group.description.replace("`n", "").replace("`r", "") } else { "" }
            origin        = $group.origin
            originId      = $group.originId
            domain        = $group.domain
            subjectKind   = $group.subjectKind
            timeStamp     = $Timestamp
        }
    }

    # Export data based on format preference
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $CsvPath = "$BaseFilename.csv"
        Write-Verbose "Writing to CSV file: $CsvPath"
        $Groups | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "Groups data saved to: $CsvPath" -ForegroundColor Green
    }

    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonPath = "$BaseFilename.json"
        Write-Verbose "Writing to JSON file: $JsonPath"
        
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $Timestamp
                recordCount = $Groups.Count
                dataType = "Groups"
            }
            data = $Groups
        }
        
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $JsonPath -Encoding UTF8
        Write-Host "Groups data saved to: $JsonPath" -ForegroundColor Green
    }

    Write-Host "Groups collection completed. Total groups: $($Groups.Count)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error collecting groups data: $($_.Exception.Message)"
}
finally {
    # Clean up session
    if ($AzSession) {
        Write-Verbose "Removing AzDevOps session."
        Remove-AzDevOpsSession $AzSession.Id
    }
}