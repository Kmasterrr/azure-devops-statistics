# =============================================================================
# SCORING CONFIGURATION
# =============================================================================
# This file contains the centralized scoring weights used by:
# - Generate-Report-Enhanced.ps1 (HTML report)
# - Generate-PipelineSummary.ps1 (Markdown summary)
#
# Adjust these weights to change how Team Activity scores are calculated.
# Higher values = more importance in the ranking.
# =============================================================================

$Global:ScoringWeights = @{
    PRsMerged       = 5    # Completed, shipped work (highest value)
    CodeReviews     = 4    # Collaboration and knowledge sharing
    PRsCreated      = 3    # Initiative, work in progress
    WorkItems       = 2    # Delivery of planned work
    Commits         = 0.5    # Lowest weight - quantity doesn't equal quality
}

# Score thresholds for color coding
$Global:ScoreThresholds = @{
    High   = 50    # Green - major contributor
    Medium = 20    # Blue - regular contributor
    # Below Medium = Gray - light activity
}

# Number of team members to show in reports
$Global:TeamActivityLimit = 15

# =============================================================================
# Helper function to calculate activity score
# =============================================================================
function Get-ActivityScore {
    param(
        [int]$PRsMerged = 0,
        [int]$PRsCreated = 0,
        [int]$CodeReviews = 0,
        [int]$WorkItems = 0,
        [int]$Commits = 0
    )
    
    return ($PRsMerged * $Global:ScoringWeights.PRsMerged) +
           ($PRsCreated * $Global:ScoringWeights.PRsCreated) +
           ($CodeReviews * $Global:ScoringWeights.CodeReviews) +
           ($WorkItems * $Global:ScoringWeights.WorkItems) +
           ($Commits * $Global:ScoringWeights.Commits)
}

# =============================================================================
# Helper function to get scoring formula description
# =============================================================================
function Get-ScoringFormulaText {
    param(
        [switch]$Markdown,
        [switch]$HTML
    )
    
    $prsMerged = $Global:ScoringWeights.PRsMerged
    $prsCreated = $Global:ScoringWeights.PRsCreated
    $codeReviews = $Global:ScoringWeights.CodeReviews
    $workItems = $Global:ScoringWeights.WorkItems
    $commits = $Global:ScoringWeights.Commits
    
    if ($Markdown) {
        return "PRs Merged (${prsMerged}pts) + PRs Created (${prsCreated}pts) + Reviews (${codeReviews}pts) + Work Items (${workItems}pts) + Commits (${commits}pt)"
    }
    elseif ($HTML) {
        return "PRs Merged (${prsMerged}pts) + PRs Created (${prsCreated}pts) + Code Reviews (${codeReviews}pts) + Work Items (${workItems}pts) + Commits (${commits}pt)"
    }
    else {
        return "PRs Merged (${prsMerged}pts) + PRs Created (${prsCreated}pts) + Code Reviews (${codeReviews}pts) + Work Items (${workItems}pts) + Commits (${commits}pt)"
    }
}

Write-Verbose "Scoring weights loaded: PRsMerged=$($Global:ScoringWeights.PRsMerged), PRsCreated=$($Global:ScoringWeights.PRsCreated), CodeReviews=$($Global:ScoringWeights.CodeReviews), WorkItems=$($Global:ScoringWeights.WorkItems), Commits=$($Global:ScoringWeights.Commits)"
