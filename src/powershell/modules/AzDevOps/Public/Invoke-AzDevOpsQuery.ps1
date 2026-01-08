function Invoke-AzDevOpsQuery
{
    <#
    .SYNOPSIS
    Invokes an Azure Pipelines wiql query.

    .DESCRIPTION
    Creates a invoke Azure Pipelines query for the project provided.

    .PARAMETER Instance
    The Team Services account or TFS server.

    .PARAMETER Collection
    For Azure DevOps the value for collection should be the name of your orginization. 
    For both Team Services and TFS The value should be DefaultCollection unless another collection has been created.

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

    .PARAMETER Project
    Project ID or project name.

    .PARAMETER Query
    The text of the WIQL query

    .PARAMETER Team
    Team ID or team name

    .PARAMETER Top
    The max number of results to return.

    .PARAMETER TimePrecision
    Whether or not to use time precision.

    .INPUTS
    None, does not support pipeline.

    .OUTPUTS
    PSObject, Azure Pipelines query results.

    .EXAMPLE
    Invokes a wiql.

    $query = "Select [System.Id], [System.Title], [System.State] From WorkItems Where [System.WorkItemType] = 'Feature'"
    Invoke-AzDevOpsQuery -Session $session -Query $query

    .LINK
    https://docs.microsoft.com/en-us/rest/api/azure/devops/wit/wiql/query%20by%20wiql?view=azure-devops-rest-6.0
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
        $Query, 

        [Parameter()]
        [string]
        $Team, 
        
        [Parameter()]
        [int]
        $Top,

        [Parameter()]
        [bool]
        $TimePrecision
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
        $body = @{
            query = $Query
        }
        $apiEndpoint = (Get-AzDevOpsApiEndpoint -ApiType 'wit-wiql') -f $Query
        $queryParameters = Set-AzDevOpsQueryParameters -InputObject $PSBoundParameters
        $setAzDevOpsUriSplat = @{
            Collection  = $Collection
            Instance    = $Instance
            Project     = $Project
            ApiVersion  = $ApiVersion
            ApiEndpoint = $apiEndpoint
            Query       = $queryParameters
        }
        [uri] $uri = Set-AzDevOpsUri @setAzDevOpsUriSplat
        $invokeAzDevOpsRestMethodSplat = @{
            Method              = 'POST'
            Uri                 = $uri
            Credential          = $Credential
            PersonalAccessToken = $PersonalAccessToken
            Body                = $body
            ContentType         = 'application/json'
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