function Write-ProgressHelper {
	param (
    [int]$ProcessId = 0,
    [int]$Steps,
    [int]$StepNumber,
    [string]$Message = "Progress"
	)

  $PercentageComplete = 0
  $PercentageComplete = ($StepNumber / $Steps) * 100
  $Status = ($StepNumber / $Steps).ToString("P1")
 
  # Make sure percentage is correct
  if ($PercentageComplete -gt 100) { $PercentageComplete = 100 }

  Write-Debug "PID: $ProcessId - Current Step: $StepNumber - Total Steps: $Steps - PercentageComplete: $PercentageComplete"
	Write-Progress -Id $ProcessId -Activity $Message -Status $Status -PercentComplete $PercentageComplete
}