Function Remove-AzDevOpsSession
{
    <#
    .SYNOPSIS
    Removes an Azure DevOps session.

    .DESCRIPTION
    Removes an Azure DevOps session.
    If the session is saved, it will be removed from the saved sessions as well.

    .PARAMETER Id
    Session id.

    .PARAMETER Path
    The path where session data will be stored, defaults to $Script:ModuleDataPath.

    .LINK
    Save-AzDevOpsSession
    Remove-AzDevOpsSession

    .INPUTS
    PSObject. Get-AzDevOpsSession

    .OUTPUTS
    None. Does not supply output.

    .EXAMPLE
    Deletes session with the id of '2'.
    Remove-AzDevOpsSession -Id 2

    .EXAMPLE
    Deletes all sessions in memory and stored on disk.
    Remove-AzDevOpsSession
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,
            ValueFromPipelineByPropertyName)]
        [int]
        $Id,
       
        [Parameter()]
        [string]
        $Path = $Script:ModuleDataPath
    )
    Process
    {
        $sessions = Get-AzDevOpsSession -Id $Id
        Foreach ($session in $sessions)
        {
            If ($session.Saved -eq $true)
            {
                $newData = @{SessionData = @() }
                $data = Get-Content -Path $Path -Raw | ConvertFrom-Json
                Foreach ($_data in $data.SessionData)
                {
                    If ($_data.Id -eq $session.Id)
                    {
                        Continue
                    }       
                    else
                    {
                        $newData.SessionData += $_data
                    }
                }
                $newData | ConvertTo-Json -Depth 5 | Out-File -FilePath $Path
            }
            [array] $Global:_AzDevOpsSessions = $Global:_AzDevOpsSessions | Where-Object { $PSItem.Id -ne $session.Id }
        }
    }
}