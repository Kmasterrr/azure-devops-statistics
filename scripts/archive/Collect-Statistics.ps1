[CmdletBinding()]
param 
(
)

<###[Environment Variables]#####################>
$Organization      = $env:ADOS_ORGANIZATION
$StorageAccount    = $env:AZURE_STORAGE_ACCOUNT
$StorageAccountKey = $env:AZURE_STORAGE_ACCOUNT_KEY
$ContainerName     = $env:AZURE_STORAGE_CONTAINER

<###[Set Paths]#################################>
$ScriptPath = Join-Path $PSScriptRoot "\src\powershell"
$ModulePath = Join-Path $PSScriptRoot "\src\powershell\modules"
$DataPath   = Join-Path $PSScriptRoot "\data"

<###[Load Modules]##############################>
Import-Module (Join-Path $ModulePath "ADOS") -Force
Import-Module (Join-Path $ModulePath "AzDevOps") -Force
Import-Module (Join-Path $ModulePath "AzStorage") -Force

<###[Script Variables]##########################>
$FileDate     = Get-Date -Format "yyyy-MM-dd"
$DataPath     = Join-Path $DataPath "$($Organization)" "$($FileDate)"

$ScriptFiles = @(
    "Users.ps1", 
    "Groups.ps1",
    "GroupMemberships.ps1",
    "GitPullRequests.ps1",
    "GitCommits.ps1",
    "ProjectStatistics.ps1",
    "OrganizationStatistics.ps1")

Write-Verbose "Collecting Azure DevOps Statistics."

$x = 0
foreach ($file in $ScriptFiles)
{
    # $i = (($i -eq $null) ? 1 : $i++)
    Write-Debug "Step Value: $x"
    Write-ProgressHelper -Message "Get Azure DevOps Stats" -ProcessId 1 -Steps $ScriptFiles.Count -StepNumber ($x++)

    $FilePath = Join-Path $ScriptPath $file
    Write-Debug "Executing script: [$($FilePath)]"
    . $($FilePath)
}

Write-Verbose "Copying files to Azure Blob Container."
Copy-ToAzureBlob -uploadRootDirectory $DataPath -storageAccountName $StorageAccount -storageAccountKey $StorageAccountKey -containerName $ContainerName -blobTier 'Hot'
