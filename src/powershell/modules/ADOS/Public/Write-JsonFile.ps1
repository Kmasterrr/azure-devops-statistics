function Write-JsonFile {
    [CmdletBinding()]
      param (
          [PSCustomObject]$Data,
          [int]$Depth = 5,
          [string]$Path,
          [string]$Filename
      )
  
    begin { 
      $outputFile = (Join-Path $Path $Filename)
      Write-Debug "Writing JSON file: $($outputFile)" 
    }
  
    process {
      $Data | ConvertTo-Json -Depth $Depth | Out-File -FilePath $outputFile
    }
  }