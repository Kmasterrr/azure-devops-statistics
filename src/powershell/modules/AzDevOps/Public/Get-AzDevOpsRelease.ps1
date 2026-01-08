function Get-AzDevOpsRelease
{
    <#
    .SYNOPSIS
    Returns Azure Pipeline release.

    .DESCRIPTION
    Returns Azure Pipeline release by release id.
    The id can be retrieved by using Get-AzDevOpsReleaseList.

    .PARAMETER Instance
    The Team Services account or TFS server.
    
    .PARAMETER Collection
    For Azure DevOps the value for collection should be the name of your orginization. 
    For both Team Services and TFS The value should be DefaultCollection unless another collection has been created.

    .PARAMETER Project
    Project ID or project name.

    .PARAMETER ApiVersion
    Version of the api to use.

    .PARAMETER PersonalAccessToken
    Personal access token used to authenticate that has been converted to a secure string. 
    It is recomended to uses an Azure Pipelines PS session to pass the personal access token parameter among funcitons, See New-AzDevOpsSession.
    https://docs.microsoft.com/en-us/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate?view=vsts
    
    .PARAMETER Credential
    Specifies a user account that has permission to send the request.

    .PARAMETER Proxy
    Use a proxy server for the request, rather than connecting directly to the Internet resource. Enter the URI of a network proxy server.

    .PARAMETER ProxyCredential
    Specifie a user account that has permission to use the proxy server that is specified by the -Proxy parameter. The default is the current user.

    .PARAMETER Session
    Azure DevOps PS session, created by New-AzDevOpsSession.

    .PARAMETER ReleaseId
    Id of the release.

    .PARAMETER ApprovalFilters
    A filter which would allow fetching approval steps selectively based on whether it is automated, or manual. This would also decide whether we should fetch pre and post approval snapshots. Assumes All by default.

    .PARAMETER PropertyFilters
    A comma-delimited list of extended properties to be retrieved. If set, the returned Release will contain values for the specified property Ids (if they exist). If not set, properties will not be included.

    .PARAMETER Expand
    A property that should be expanded in the release.

    .PARAMETER TopGateRecords
    Number of release gate records to get. Default is 5.

    .INPUTS
    None, does not support pipeline.

    .OUTPUTS
    PSObject, Azure Pipelines release(s)

    .EXAMPLE
    Returns AP release with the release id of 7.
    Get-AzDevOpsRelease -Instance 'https://dev.azure.com' -Collection 'myCollection' -Project 'myFirstProject' -ReleaseId 7

    .LINK
    https://docs.microsoft.com/en-us/rest/api/azure/devops/release/releases/get%20release?view=azure-devops-rest-5.0
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByPersonalAccessToken')]
    Param
    (
        [Parameter(Mandatory,
            ParameterSetName = 'ByPersonalAccessToken')]
        [Parameter(Mandatory,
            ParameterSetName = 'ByCredential')]
        [uri]
        $Instance,

        [Parameter(Mandatory,
            ParameterSetName = 'ByPersonalAccessToken')]
        [Parameter(Mandatory,
            ParameterSetName = 'ByCredential')]
        [string]
        $Collection,

        [Parameter(Mandatory,
            ParameterSetName = 'ByPersonalAccessToken')]
        [Parameter(Mandatory,
            ParameterSetName = 'ByCredential')]
        [string]
        $ApiVersion,

        [Parameter(ParameterSetName = 'ByPersonalAccessToken')]
        [Security.SecureString]
        $PersonalAccessToken,

        [Parameter(ParameterSetName = 'ByCredential')]
        [pscredential]
        $Credential,

        [Parameter(ParameterSetName = 'ByPersonalAccessToken')]
        [Parameter(ParameterSetName = 'ByCredential')]
        [string]
        $Proxy,

        [Parameter(ParameterSetName = 'ByPersonalAccessToken')]
        [Parameter(ParameterSetName = 'ByCredential')]
        [pscredential]
        $ProxyCredential,

        [Parameter(Mandatory,
            ParameterSetName = 'BySession')]
        [object]
        $Session,

        [Parameter(Mandatory)]
        [string]
        $Project,

        [Parameter(Mandatory)]
        [string]
        $ReleaseId,

        [Parameter()]
        [string]
        $ApprovalFilters,

        [Parameter()]
        [string]
        $PropertyFilters,

        [Parameter()]
        [string]
        [ValidateSet('none', 'tasks')]
        $Expand,

        [Parameter()]
        [int]
        $TopGateRecords
    )

    begin
    {
        If ($PSCmdlet.ParameterSetName -eq 'BySession')
        {
            $currentSession = $Session | Get-AzDevOpsSession
            If ($currentSession)
            {
                $Instance = $currentSession.Instance
                $Collection = $currentSession.Collection
                $PersonalAccessToken = $currentSession.PersonalAccessToken
                $Credential = $currentSession.Credential
                $Proxy = $currentSession.Proxy
                $ProxyCredential = $currentSession.ProxyCredential
                If ($currentSession.Version)
                {
                    $ApiVersion = (Get-AzDevOpsApiVersion -Version $currentSession.Version)
                }
                else
                {
                    $ApiVersion = $currentSession.ApiVersion
                }
            }
        }
    }
    
    process
    {
  
        $apiEndpoint = (Get-AzDevOpsApiEndpoint -ApiType 'release-releaseId') -f $ReleaseId
        $queryParameters = Set-AzDevOpsQueryParameters -InputObject $PSBoundParameters
        $setAzDevOpsUriSplat = @{
            Collection         = $Collection
            Instance           = $Instance
            Project            = $Project
            ApiVersion         = $ApiVersion
            ApiEndpoint        = $apiEndpoint
            Query              = $queryParameters
            ApiSubDomainSwitch = 'vsrm'
        }
        [uri] $uri = Set-AzDevOpsUri @setAzDevOpsUriSplat
        $invokeAzDevOpsRestMethodSplat = @{
            Method              = 'GET'
            Uri                 = $uri
            Credential          = $Credential
            PersonalAccessToken = $PersonalAccessToken
            Proxy               = $Proxy
            ProxyCredential     = $ProxyCredential
        }
        $results = Invoke-AzDevOpsRestMethod @invokeAzDevOpsRestMethodSplat 
        If ($results.value)
        {
            return $results.value
        }
        else
        {
            return $results
        }
    }
    
    end
    {
    }
}