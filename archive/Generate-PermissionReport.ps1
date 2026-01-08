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
    "ProjectGroupsAndUsers.ps1", 
    "GitRepositoriesPermissions.ps1",
    "ProjectLevelPermissions.ps1")

Write-Verbose "Generating Azure DevOps Git Permissions report..."

$x = 0
foreach ($file in $ScriptFiles)
{
    Write-ProgressHelper -Message "Git Permissions - $($file)" -ProcessId 1 -Steps $ScriptFiles.Count -StepNumber ($x++)
    Write-Verbose "Executing scripts: [$($x) of $($ScriptFiles.Count)]"

    $FilePath = Join-Path $ScriptPath $file
    Write-Debug "Executing: [$($FilePath)]"
    Invoke-Expression "$FilePath"
}

Write-Verbose "Copying files to Azure Blob Container."
Copy-ToAzureBlob -uploadRootDirectory $DataPath -storageAccountName $StorageAccount -storageAccountKey $StorageAccountKey -containerName $ContainerName -blobTier 'Hot'
