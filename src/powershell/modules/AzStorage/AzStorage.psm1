# Module Variables
$Script:PSModuleRoot = $PSScriptRoot
$Script:ModuleName = "AzDevOps"
$Script:ModuleDataRoot = (Join-Path -Path ([Environment]::GetFolderPath('ApplicationData')) -ChildPath $Script:ModuleName)
$Script:ModuleDataPath = (Join-Path -Path $Script:ModuleDataRoot -ChildPath "ModuleData.json")

$Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )

Foreach($import in @($Public + $Private))
{
    Try
    {
        Write-Verbose -Message "Dot sourcing [$($import.BaseName)]..."
        . $import.FullName
    }
    Catch
    {
        Write-Error -Message "Failed to import function $($import.FullName): $_"
    }
}

Export-ModuleMember -Function $Public.BaseName