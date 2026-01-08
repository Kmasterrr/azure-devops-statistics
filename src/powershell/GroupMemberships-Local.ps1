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

$BaseFilename = Join-Path $OutputPath "GroupMemberships"

Write-Host "Collecting Group Memberships data..." -ForegroundColor Yellow

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

    # Get list of groups first
    Write-Verbose "Getting list of groups."
    $allGroups = Get-AzDevOpsGroupList -Session $AzSession
    Write-Host "Found $($allGroups.Count) groups." -ForegroundColor Green

    $allMemberships = @()
    $stepCounter = 0

    foreach($group in $allGroups) {
        Write-ProgressHelper -Message "Processing Group Memberships" -Steps $allGroups.Count -StepNumber ($stepCounter++)

        try {
            Write-Verbose "Getting memberships for group: $($group.displayName)"
            $memberships = Get-AzDevOpsGroupMembershipList -Session $AzSession -GroupDescriptor $group.descriptor

            foreach($membership in $memberships) {
                $allMemberships += [PSCustomObject]@{
                    groupId           = $group.descriptor
                    groupName         = $group.displayName
                    groupPrincipalName = $group.principalName
                    memberId          = $membership.descriptor
                    memberDisplayName = $membership.displayName
                    memberPrincipalName = $membership.principalName
                    memberType        = $membership.subjectKind
                    timeStamp         = $Timestamp
                }
            }
        }
        catch {
            Write-Warning "Error processing memberships for group '$($group.displayName)': $($_.Exception.Message)"
        }
    }

    # Export data based on format preference
    if ($Format -eq "CSV" -or $Format -eq "Both") {
        $CsvPath = "$BaseFilename.csv"
        Write-Verbose "Writing to CSV file: $CsvPath"
        $allMemberships | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "Group Memberships data saved to: $CsvPath" -ForegroundColor Green
    }

    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $JsonPath = "$BaseFilename.json"
        Write-Verbose "Writing to JSON file: $JsonPath"
        
        $JsonData = @{
            metadata = @{
                organization = $Organization
                collectionDate = $Timestamp
                recordCount = $allMemberships.Count
                groupsProcessed = $allGroups.Count
                dataType = "GroupMemberships"
            }
            data = $allMemberships
        }
        
        $JsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $JsonPath -Encoding UTF8
        Write-Host "Group Memberships data saved to: $JsonPath" -ForegroundColor Green
    }

    Write-Host "Group Memberships collection completed. Total memberships: $($allMemberships.Count)" -ForegroundColor Cyan

}
catch {
    Write-Error "Error collecting group memberships data: $($_.Exception.Message)"
}
finally {
    # Clean up session
    if ($AzSession) {
        Write-Verbose "Removing AzDevOps session."
        Remove-AzDevOpsSession $AzSession.Id
    }
}