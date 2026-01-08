function Set-AzDevOpsAuthenticationType
{ 
    <#
    .SYNOPSIS
    Sets the authentication type used by Invoke-AzDevOpsRestMethod.

    .DESCRIPTION
    Sets the authentication type used by Invoke-AzDevOpsRestMethod.
    Default authentication will use the pesonal access token that is stored in session data, unless a credential is provided.

    .PARAMETER InputObject
    The splat parameters used by Invoke-AzDevOpsRestMethod.

    .PARAMETER Credential
    Specifies a user account that has permission to send the request.

    .OUTPUTS
    PSObject, The modifed inputobject.

    .EXAMPLE
    Set-AzDevOpsAuthenticationType -InputObject $inputObject

    .EXAMPLE
    Sets the AP authentication to the credential provided for the input object.
    Set-AzDevOpsAuthenticationType -InputObject $inputObject -Credential $pscredential

    .EXAMPLE
    Sets the AP authentication to the personal access token provided for the input object.
    Set-AzDevOpsAuthenticationType -InputObject $inputObject -PersonalAccessToken $mySecureToken

    .LINK
    https://docs.microsoft.com/en-us/azure/devops/integrate/get-started/authentication/authentication-guidance?view=vsts
    #>
    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory)]
        [PSObject]
        $InputObject,

        [Parameter()]
        [Security.SecureString]
        $PersonalAccessToken,

        [Parameter()]
        [pscredential]
        $Credential
    )

    begin
    {
    }

    process
    {
        If ($Credential)
        {
            Write-Verbose "[$($MyInvocation.MyCommand.Name)]: Authenticating with the provided credential."
            $InputObject.Credential = $Credential
        }
        elseIf ($PersonalAccessToken)
        {
            Write-Verbose "[$($MyInvocation.MyCommand.Name)]: Authenticating with the stored personal access token."
            $PersonalAccessTokenToken = Unprotect-AzDevOpsSecurePersonalAccessToken -PersonalAccessToken $PersonalAccessToken
            $encodedPersonalAccessToken = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$PersonalAccessTokenToken"))
            $InputObject.Headers = @{Authorization = "Basic $encodedPersonalAccessToken" }
        }
        else
        {
            Write-Verbose "[$($MyInvocation.MyCommand.Name)]: Authenticating with default credentials"
            $InputObject.UseDefaultCredentials = $true
        }
    }

    end
    {
        return $InputObject
    }
}