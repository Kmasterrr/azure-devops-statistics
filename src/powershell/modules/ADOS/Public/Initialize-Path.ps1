function Initialize-Path {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    process {
        if (!(Test-Path $Path))
        {
            # Create path for data files.
            Write-Verbose "Directory not found. [$($Path)]"
            Write-Verbose "Creating directory [$($Path)]."
            New-Item $Path -Type Directory
        }
    }
}