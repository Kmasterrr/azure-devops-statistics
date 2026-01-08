function Split-Array {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory, ValueFromPipeline)]
      [String[]] $InputObject,
      [ValidateRange(1, [int]::MaxValue)]
      [int] $Size = 10
    )

    begin   { 
      $items = New-Object System.Collections.Generic.List[object] 
    }

    process { 
      $items.AddRange($InputObject) 
    }
    
    end {
      $chunkCount = [Math]::Floor($items.Count / $Size)

      foreach ($chunkNdx in 0..($chunkCount-1)) {
        , $items.GetRange($chunkNdx * $Size, $Size).ToArray()
      }

      if ($chunkCount * $Size -lt $items.Count) {
        , $items.GetRange($chunkCount * $Size, $items.Count - $chunkCount * $Size).ToArray()
      }
    }
}