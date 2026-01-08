# =============================================================================
# Test-TeamsWebhook.ps1
# =============================================================================
# Quick test script to verify Teams/Power Automate webhook is working
# Supports both:
#   - Teams Incoming Webhook (Adaptive Cards)
#   - Power Automate HTTP trigger (simple JSON)
#
# Usage: .\Test-TeamsWebhook.ps1 -WebhookUrl "https://..."
# =============================================================================

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$WebhookUrl = ""
)

# If no webhook provided, check environment variable
if (-not $WebhookUrl) {
    $WebhookUrl = $env:TEAMS_WEBHOOK_URL
}

if (-not $WebhookUrl) {
    Write-Host "Usage: .\Test-TeamsWebhook.ps1 -WebhookUrl 'https://your-webhook-url'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Or set environment variable:" -ForegroundColor Yellow
    Write-Host '  $env:TEAMS_WEBHOOK_URL = "https://your-webhook-url"' -ForegroundColor Cyan
    Write-Host "  .\Test-TeamsWebhook.ps1" -ForegroundColor Cyan
    exit 1
}

Write-Host "Testing webhook..." -ForegroundColor Cyan
Write-Host "Webhook URL: $($WebhookUrl.Substring(0, [Math]::Min(60, $WebhookUrl.Length)))..." -ForegroundColor Gray

# Detect webhook type
$isPowerAutomate = $WebhookUrl -match "powerautomate|flow\.microsoft|logic\.azure"
if ($isPowerAutomate) {
    Write-Host "Detected: Power Automate workflow" -ForegroundColor Magenta
} else {
    Write-Host "Detected: Teams Incoming Webhook" -ForegroundColor Magenta
}

# Load real data if available
$Organization = $env:ADOS_ORGANIZATION
if (-not $Organization) { $Organization = "TestOrganization" }

$BaseDataPath = Join-Path (Join-Path $PSScriptRoot "data") $Organization
$summary = $null
$topContributors = @()

if (Test-Path $BaseDataPath) {
    $LatestDate = Get-ChildItem -Path $BaseDataPath -Directory | Sort-Object Name -Descending | Select-Object -First 1
    if ($LatestDate) {
        $summaryPath = Join-Path $LatestDate.FullName "Collection-Summary.json"
        if (Test-Path $summaryPath) {
            $summary = Get-Content $summaryPath | ConvertFrom-Json
            Write-Host "Loaded real data from: $($LatestDate.Name)" -ForegroundColor Green
        }
        
        # Load scoring config
        $configPath = Join-Path $PSScriptRoot "Config-ScoringWeights.ps1"
        if (Test-Path $configPath) {
            . $configPath
        } else {
            # Default weights if config not found
            $Global:ScoringWeights = @{ PRsMerged = 5; CodeReviews = 4; PRsCreated = 3; WorkItems = 2; Commits = 0.5 }
        }
        
        # Load data files and calculate top contributors
        $usersPath = Join-Path $LatestDate.FullName "Users.json"
        $commitsPath = Join-Path $LatestDate.FullName "GitCommits.json"
        $prsPath = Join-Path $LatestDate.FullName "GitPullRequests.json"
        $workItemsPath = Join-Path $LatestDate.FullName "WorkItems.json"
        
        if ((Test-Path $commitsPath) -or (Test-Path $prsPath)) {
            $commitsData = if (Test-Path $commitsPath) { Get-Content $commitsPath | ConvertFrom-Json } else { @{ data = @() } }
            $prsData = if (Test-Path $prsPath) { Get-Content $prsPath | ConvertFrom-Json } else { @{ data = @() } }
            $workItemsData = if (Test-Path $workItemsPath) { Get-Content $workItemsPath | ConvertFrom-Json } else { @{ data = @() } }
            
            # Extract data arrays (handle nested structure)
            $commitsList = if ($commitsData.data) { $commitsData.data } else { $commitsData }
            $prsList = if ($prsData.data) { $prsData.data } else { $prsData }
            $workItemsList = if ($workItemsData.data) { $workItemsData.data } else { $workItemsData }
            
            # Build team activity by NAME (matching HTML report logic)
            $teamActivity = @{}
            
            # Count commits per person (by author name)
            foreach ($commit in $commitsList) {
                $name = $commit.Author
                if ($name) {
                    if (-not $teamActivity.ContainsKey($name)) {
                        $teamActivity[$name] = @{ Name = $name; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0 }
                    }
                    $teamActivity[$name].Commits++
                }
            }
            
            # Count PRs created, merged, and reviews per person (by display name)
            foreach ($pr in $prsList) {
                $creator = $pr.createdBy
                if ($creator) {
                    if (-not $teamActivity.ContainsKey($creator)) {
                        $teamActivity[$creator] = @{ Name = $creator; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0 }
                    }
                    $teamActivity[$creator].PRsCreated++
                    if ($pr.status -eq "completed") {
                        $teamActivity[$creator].PRsMerged++
                    }
                }
                # Count reviewers
                if ($pr.reviewers) {
                    $reviewerList = $pr.reviewers -split "; "
                    foreach ($reviewer in $reviewerList) {
                        if ($reviewer -and $reviewer.Trim()) {
                            $reviewerName = $reviewer.Trim()
                            if (-not $teamActivity.ContainsKey($reviewerName)) {
                                $teamActivity[$reviewerName] = @{ Name = $reviewerName; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0 }
                            }
                            $teamActivity[$reviewerName].PRsReviewed++
                        }
                    }
                }
            }
            
            # Count work items assigned per person (by assignedTo)
            foreach ($wi in $workItemsList) {
                if ($wi.assignedTo) {
                    $assignee = $wi.assignedTo
                    if (-not $teamActivity.ContainsKey($assignee)) {
                        $teamActivity[$assignee] = @{ Name = $assignee; Commits = 0; PRsCreated = 0; PRsMerged = 0; PRsReviewed = 0; WorkItemsAssigned = 0 }
                    }
                    $teamActivity[$assignee].WorkItemsAssigned++
                }
            }
            
            # Calculate scores
            $userScores = foreach ($person in $teamActivity.Keys) {
                $activity = $teamActivity[$person]
                $score = ($activity.PRsMerged * $Global:ScoringWeights.PRsMerged) +
                         ($activity.PRsCreated * $Global:ScoringWeights.PRsCreated) +
                         ($activity.PRsReviewed * $Global:ScoringWeights.CodeReviews) +
                         ($activity.WorkItemsAssigned * $Global:ScoringWeights.WorkItems) +
                         ($activity.Commits * $Global:ScoringWeights.Commits)
                
                if ($score -gt 0) {
                    [PSCustomObject]@{
                        Name = $activity.Name
                        Score = [math]::Round($score, 1)
                        PRsMerged = $activity.PRsMerged
                        Reviews = $activity.PRsReviewed
                        Commits = $activity.Commits
                    }
                }
            }
            
            $topContributors = $userScores | Sort-Object Score -Descending | Select-Object -First 5
            Write-Host "Calculated top $($topContributors.Count) contributors" -ForegroundColor Green
        }
    }
}

# Prepare data
$orgName = if ($summary) { $summary.Organization } else { $Organization }
$users = if ($summary) { "$($summary.UsersCollected)" } else { "41" }
$repos = if ($summary) { "$($summary.RepositoriesCollected)" } else { "32" }
$commits = if ($summary) { "$($summary.CommitsCollected)" } else { "411" }
$prs = if ($summary) { "$($summary.PullRequestsCollected)" } else { "18" }
$workItems = if ($summary) { "$($summary.WorkItemsCollected)" } else { "76" }
$duration = if ($summary) { "$($summary.Duration)".Split('.')[0] } else { "00:00:19" }

if ($isPowerAutomate) {
    # Power Automate - send simple JSON that the flow can process
    $body = @{
        Organization = $orgName
        Users = $users
        Repositories = $repos
        Commits = $commits
        PullRequests = $prs
        WorkItems = $workItems
        Duration = $duration
        CollectionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        IsTest = $true
    } | ConvertTo-Json -Depth 10
} else {
    # Build top contributors section for the card
    $contributorItems = @()
    if ($topContributors.Count -gt 0) {
        $contributorItems += @{
            type = "TextBlock"
            text = "Top 5 Team Activity"
            weight = "Bolder"
            size = "Medium"
            spacing = "Large"
            separator = $true
        }
        
        # Create a nicely formatted list
        $rank = 1
        foreach ($contributor in $topContributors) {
            $medal = switch ($rank) { 1 { "1." } 2 { "2." } 3 { "3." } default { "$rank." } }
            $contributorItems += @{
                type = "ColumnSet"
                spacing = "Small"
                columns = @(
                    @{ 
                        type = "Column"
                        width = "auto"
                        items = @(@{ type = "TextBlock"; text = $medal; weight = "Bolder"; size = "Default" })
                    }
                    @{ 
                        type = "Column"
                        width = "stretch"
                        items = @(@{ type = "TextBlock"; text = $contributor.Name; weight = "Bolder"; wrap = $true })
                    }
                    @{ 
                        type = "Column"
                        width = "auto"
                        items = @(@{ type = "TextBlock"; text = "$($contributor.Score)"; color = "Good"; weight = "Bolder"; size = "Default" })
                    }
                )
            }
            $rank++
        }
    }
    
    # Teams Incoming Webhook - send Adaptive Card matching pipeline summary format
    $cardBody = @(
        @{
            type = "TextBlock"
            size = "Large"
            weight = "Bolder"
            text = "Azure DevOps Statistics Report"
            color = "Accent"
        }
        @{
            type = "TextBlock"
            text = "TEST - Weekly statistics collection completed successfully"
            wrap = $true
            spacing = "Small"
        }
        @{
            type = "Container"
            style = "emphasis"
            items = @(
                @{
                    type = "ColumnSet"
                    columns = @(
                        @{
                            type = "Column"
                            width = "stretch"
                            items = @(
                                @{ type = "TextBlock"; text = "Organization"; weight = "Bolder"; size = "Small" }
                                @{ type = "TextBlock"; text = $orgName; spacing = "None" }
                            )
                        }
                        @{
                            type = "Column"
                            width = "stretch"
                            items = @(
                                @{ type = "TextBlock"; text = "Duration"; weight = "Bolder"; size = "Small" }
                                @{ type = "TextBlock"; text = $duration; spacing = "None" }
                            )
                        }
                    )
                }
            )
        }
        @{
            type = "ColumnSet"
            separator = $true
            spacing = "Medium"
            columns = @(
                @{
                    type = "Column"
                    width = "stretch"
                    items = @(
                        @{ type = "TextBlock"; text = "Users"; weight = "Bolder"; size = "Small"; color = "Accent" }
                        @{ type = "TextBlock"; text = $users; size = "ExtraLarge"; spacing = "None" }
                    )
                }
                @{
                    type = "Column"
                    width = "stretch"
                    items = @(
                        @{ type = "TextBlock"; text = "Repositories"; weight = "Bolder"; size = "Small"; color = "Accent" }
                        @{ type = "TextBlock"; text = $repos; size = "ExtraLarge"; spacing = "None" }
                    )
                }
                @{
                    type = "Column"
                    width = "stretch"
                    items = @(
                        @{ type = "TextBlock"; text = "Commits"; weight = "Bolder"; size = "Small"; color = "Accent" }
                        @{ type = "TextBlock"; text = $commits; size = "ExtraLarge"; spacing = "None" }
                    )
                }
            )
        }
        @{
            type = "ColumnSet"
            spacing = "Small"
            columns = @(
                @{
                    type = "Column"
                    width = "stretch"
                    items = @(
                        @{ type = "TextBlock"; text = "Pull Requests"; weight = "Bolder"; size = "Small"; color = "Accent" }
                        @{ type = "TextBlock"; text = $prs; size = "ExtraLarge"; spacing = "None" }
                    )
                }
                @{
                    type = "Column"
                    width = "stretch"
                    items = @(
                        @{ type = "TextBlock"; text = "Work Items"; weight = "Bolder"; size = "Small"; color = "Accent" }
                        @{ type = "TextBlock"; text = $workItems; size = "ExtraLarge"; spacing = "None" }
                    )
                }
                @{
                    type = "Column"
                    width = "stretch"
                    items = @()
                }
            )
        }
    )
    
    # Add contributor items if we have them
    $cardBody += $contributorItems
    
    $teamsMessage = @{
        type = "message"
        attachments = @(
            @{
                contentType = "application/vnd.microsoft.card.adaptive"
                contentUrl = $null
                content = @{
                    '$schema' = "http://adaptivecards.io/schemas/adaptive-card.json"
                    type = "AdaptiveCard"
                    version = "1.4"
                    body = $cardBody
                    actions = @(
                        @{
                            type = "Action.OpenUrl"
                            title = "View Pipeline Summary"
                            url = "https://dev.azure.com"
                        }
                        @{
                            type = "Action.OpenUrl"
                            title = "Download HTML Report"
                            url = "https://dev.azure.com"
                        }
                    )
                }
            }
        )
    }
    $body = $teamsMessage | ConvertTo-Json -Depth 20
}

Write-Host ""
Write-Host "Sending test message..." -ForegroundColor Yellow
Write-Host "Payload preview:" -ForegroundColor Gray
Write-Host ($body | ConvertFrom-Json | ConvertTo-Json -Depth 3 | Select-Object -First 20) -ForegroundColor DarkGray

try {
    $response = Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $body -ContentType "application/json"
    Write-Host ""
    Write-Host "SUCCESS! Check your Teams channel for the message." -ForegroundColor Green
    if ($response) {
        Write-Host "Response: $response" -ForegroundColor Gray
    }
}
catch {
    Write-Host ""
    Write-Host "FAILED to send message:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        Write-Host "Status Code: $statusCode" -ForegroundColor Red
        
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            Write-Host "Response: $responseBody" -ForegroundColor Red
        } catch {}
    }
    exit 1
}
