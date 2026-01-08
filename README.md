# Azure DevOps Statistics

A comprehensive tool for collecting and visualizing statistics from your Azure DevOps organization. Generates beautiful HTML reports with charts, team activity metrics, and project insights.

![Azure DevOps](https://img.shields.io/badge/Azure%20DevOps-0078D7?style=flat&logo=azure-devops&logoColor=white)
![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?style=flat&logo=powershell&logoColor=white)

## Features

- **Automated Data Collection** - Gathers statistics from all projects in your organization
- **Pull Request Analytics** - Tracks PRs created, merged, and reviewed
- **Team Activity Scoring** - Fair, multi-metric ranking of team contributions
- **Interactive Charts** - Visual insights using Chart.js
- **Azure Pipeline Integration** - Scheduled reports with artifacts
- **Export Formats** - JSON and CSV for further analysis

---

## Quick Start

### Prerequisites

- PowerShell 5.1 or later
- Azure DevOps Personal Access Token (PAT) with read permissions for:
  - Code (read)
  - Build (read)
  - Work Items (read)
  - Graph (read)
  - Project and Team (read)

### Local Execution

1. Set environment variables:
   ```powershell
   $env:ADOS_ORGANIZATION = "your-organization-name"
   $env:ADOS_PAT = "your-personal-access-token"
   ```

2. Run data collection:
   ```powershell
   .\Collect-Statistics-Fixed.ps1
   ```

3. Generate the report:
   ```powershell
   .\Generate-Report-Enhanced.ps1 -OpenReport
   ```

### Azure Pipeline Execution

The included `azure-pipelines-self.yaml` runs on a weekly schedule (Sundays at 6 AM UTC) and publishes the report as a pipeline artifact.

**Setup:**
1. Create a pipeline variable named `ADOS_PAT` containing your PAT (mark as secret)
2. The `ADOS_ORGANIZATION` is automatically detected from the pipeline context
3. Run the pipeline manually or wait for the scheduled trigger

**Viewing the Report:**
1. Go to the pipeline run summary
2. The report summary is displayed directly on the page
3. Click **"Artifacts"** to download the full HTML report

### Teams Notifications

The pipeline can send a summary notification to Microsoft Teams when the report is generated.

**Setup:**
1. In your Teams channel, add an **Incoming Webhook** connector:
   - Click the `...` menu on the channel → **Connectors**
   - Add **Incoming Webhook** and copy the URL
2. In the Azure DevOps pipeline, add a variable:
   - Name: `TEAMS_WEBHOOK_URL`
   - Value: Your webhook URL (mark as secret)

**Notification includes:**
- Organization name and collection duration
- Summary metrics (users, repos, commits, PRs, work items)
- Top 5 Team Activity with scores
- Quick links to pipeline summary and report artifact

**Test the webhook locally:**
```powershell
.\Test-TeamsWebhook.ps1 -WebhookUrl "https://your-webhook-url"
```

---

## Data Collection

### APIs Used

The collection script queries the following Azure DevOps REST APIs (v7.1-preview):

| Data Type | API Endpoint | Description |
|-----------|--------------|-------------|
| **Projects** | `/_apis/projects` | All projects in the organization |
| **Users** | `/_apis/graph/users` | Organization members with license info |
| **Repositories** | `/{project}/_apis/git/repositories` | Git repositories per project |
| **Commits** | `/{project}/_apis/git/repositories/{repo}/commits` | Commit history (last 500 per repo) |
| **Pull Requests** | `/{project}/_apis/git/repositories/{repo}/pullrequests` | PRs with status, reviewers, dates |
| **Pipelines** | `/{project}/_apis/pipelines` | Build/release pipeline definitions |
| **Builds** | `/{project}/_apis/build/builds` | Build execution history |
| **Work Items** | `/{project}/_apis/wit/wiql` | Work items via WIQL query |

### Collected Fields

#### Git Commits
- `commitId`, `author`, `authorDate`, `message`
- `repository`, `repositoryId`, `project`, `projectId`

#### Pull Requests
- `pullRequestId`, `title`, `status` (active/completed/abandoned)
- `createdBy`, `creationDate`, `closedDate`, `closedBy`
- `sourceBranch`, `targetBranch`, `mergeStatus`
- `reviewers` (semicolon-separated list)

#### Work Items
- `id`, `workItemType`, `title`, `state`
- `assignedTo`, `createdBy`, `createdDate`
- `project`, `areaPath`, `iterationPath`

#### Users
- `displayName`, `mailAddress`, `principalName`
- `license`, `status` (active/inactive)

---

## Team Activity Score

The report includes a **Team Activity** section that ranks contributors using a weighted scoring system designed to measure meaningful contributions, not just commit volume.

### Scoring Formula

```
Activity Score = (PRs Merged × 5) + (Code Reviews × 4) + (PRs Created × 3) + (Work Items × 2) + (Commits × 0.5)
```

### Point Values

| Activity | Points | Rationale |
|----------|--------|-----------|
| **PRs Merged** | 5 | Represents completed, shipped work that has been reviewed and integrated |
| **Code Reviews** | 4 | Collaboration and knowledge sharing; reviewers help maintain code quality |
| **PRs Created** | 3 | Shows initiative and work in progress; creating PRs drives the review process |
| **Work Items Assigned** | 2 | Delivery of planned work; reflects participation in sprint/project planning |
| **Commits** | 0.5 | Lowest weight because commit count doesn't reflect quality or impact |

### Why This Approach?

Traditional metrics like commit count can be misleading:
- Someone could make 100 tiny commits while another person ships a major feature in 5 commits
- Commits don't capture collaboration (code reviews)
- Commits don't reflect work item completion

The weighted score provides a more holistic view of contribution by valuing:
1. **Shipping** (merged PRs) over starting (commits)
2. **Collaboration** (reviews) alongside individual work
3. **Planned work** (work items) to align with team goals

### Score Interpretation
 
| Score Range | Color | Meaning |
|-------------|-------|---------|
| 50+ | 🟢 Green | High activity, major contributor |
| 20-49 | 🔵 Blue | Moderate activity, regular contributor |
| 1-19 | ⚫ Gray | Light activity |

---

## Report Sections

The generated HTML report includes:

### Summary Metrics
- Total projects, users, repositories
- Commit count (with 30-day trend)
- Build count and pipeline count
- Pull request and work item totals

### Charts
1. **Activity by Project** - Bar chart comparing commits, builds, and work items
2. **Build Status Distribution** - Doughnut chart with success rate
3. **Commits by Day of Week** - Activity patterns
4. **License Distribution** - User license breakdown

### Tables
1. **Team Activity** - Ranked contributors with multi-metric breakdown
2. **Project Statistics** - Per-project metrics with build success rates
3. **User Analytics** - License and status distribution
4. **Work Items Analysis** - By type and state
5. **Repository Details** - Size, commits, last activity

---

## Output Files

Data is saved to `data/{organization}/{date}/`:

```
data/
└── YourOrganization/
    └── 2026-01-08/
        ├── Azure-DevOps-Report.html    # Interactive report
        ├── Collection-Summary.json      # Run metadata
        ├── Users.json / Users.csv
        ├── GitRepositories.json / .csv
        ├── GitCommits.json / .csv
        ├── GitPullRequests.json / .csv
        ├── BuildPipelines.json / .csv
        ├── Builds.json / .csv
        └── WorkItems.json / .csv
```

---

## Configuration

### Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `ADOS_ORGANIZATION` | Yes | Azure DevOps organization name |
| `ADOS_PAT` | Yes | Personal Access Token |

### Script Parameters

#### Collect-Statistics-Fixed.ps1
| Parameter | Default | Description |
|-----------|---------|-------------|
| `-OutputFormat` | `Both` | Output format: `JSON`, `CSV`, or `Both` |
| `-OutputDirectory` | `./data` | Base output directory |

#### Generate-Report-Enhanced.ps1
| Parameter | Default | Description |
|-----------|---------|-------------|
| `-DataPath` | Latest | Specific data directory to use |
| `-OutputFormat` | `HTML` | Report format |
| `-OpenReport` | `$false` | Open report in browser after generation |

---

## Azure Pipeline

### Pipeline File

Use `azure-pipelines-self.yaml` for automated scheduled execution:

```yaml
schedules:
  - cron: "0 6 * * 0"  # Every Sunday at 6 AM UTC
    displayName: Weekly Statistics Collection
    branches:
      include:
        - main
    always: true
```

### Required Variables

Create these as pipeline variables (Settings → Variables):

| Variable | Secret | Value |
|----------|--------|-------|
| `ADOS_PAT` | Yes | Your PAT with required scopes |

### Pipeline Output

- **Artifact**: `azure-devops-report` containing the HTML report
- **Summary**: Markdown summary displayed on the pipeline run page

---

## Troubleshooting

### Common Issues

**"Environment variables ADOS_ORGANIZATION and ADOS_PAT must be set"**
```powershell
$env:ADOS_ORGANIZATION = "your-org"
$env:ADOS_PAT = "your-pat"
```

**"TF400813: Resource not available" or 401 errors**
- Your PAT may have expired
- PAT may not have required scopes
- Generate a new PAT with Code, Build, Work Items, Graph, and Project read permissions

**Pipeline only collects from one project**
- The default `System.AccessToken` has limited scope
- Use a user PAT stored as a secret pipeline variable instead

**Garbled characters in report**
- The report uses UTF-8 encoding with BOM
- Ensure your browser supports UTF-8

---

## Advanced: Database Integration

For long-term storage and Power BI reporting, the project includes optional components:

### Architecture

1. PowerShell Scripts → CSV/JSON files
2. Azure Blob Storage container
3. Azure Function (blob trigger)
4. Azure SQL Database
5. Power BI Reports

### Azure Function

A blob-triggered function uploads CSV data to Azure SQL when files are uploaded to storage:

```csharp
[FunctionName("FileProcessor")]
public static void Run(
    [BlobTrigger("devops-stats/{name}", Connection = "AzureStorage")] Stream blob, 
    string name, 
    ILogger log)
{
    // Processes CSV files and loads to Azure SQL via stored procedure
}
```

See [docs/setup.md](docs/setup.md) for detailed database setup instructions.

---

## License

MIT License - See LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Submit a pull request
