function Test-Debug {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory = $false)]
      [switch]$IgnorePSBoundParameters,
      [Parameter(Mandatory = $false)]
      [switch]$IgnoreDebugPreference,
      [Parameter(Mandatory = $false)]
      [switch]$IgnorePSDebugContext
    )
  
    process {
      ((-not $IgnoreDebugPreference.IsPresent) -and ($DebugPreference -ne "SilentlyContinue")) -or
      ((-not $IgnorePSBoundParameters.IsPresent) -and $PSBoundParameters.Debug.IsPresent) -or
      ((-not $IgnorePSDebugContext.IsPresent) -and ($PSDebugContext))
    }
  }