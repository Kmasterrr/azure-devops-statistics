function Get-BatchRanges {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory)]
      [int]$TotalItems,

      [Parameter(Mandatory)]
      [ValidateRange(1, [int]::MaxValue)]
      [int]$BatchSize,
      
      [Parameter()]
      [switch]$ZeroIndex
    )
  
    begin { $batches = @() }
  
    process {
      $batchCount = [Math]::Ceiling($TotalItems / $BatchSize)
      $idxStart   = if ($ZeroIndex) { 0 } else { 1 }
      $batchCount = if ($ZeroIndex -and $batchCount -gt 0) { ($batchCount - 1) } else { $batchCount }
  
      foreach ($batchIdx in ($idxStart)..($batchCount))
      {
        $rangeLow  = ($batchIdx * $batchSize)
        $rangeHigh = $batchSize
  
        if ($batchIdx -eq ($batchCount))
        {
            $rangeHigh = ($TotalItems - $rangeLow)
        }
  
        $batches += [PSCustomObject]@{
          idx   = $batchIdx
          range = @(($rangeLow), ($rangeHigh))
          count = $TotalItems
        }
      }
    }
  
    end { return $batches }
  }