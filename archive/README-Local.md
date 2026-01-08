# Azure DevOps Statistics - Local Collection

A simplified version of the Azure DevOps Statistics tool that captures comprehensive information from your Azure DevOps environments and saves it locally without requiring SQL databases or Azure Storage accounts.

## Overview

This tool extracts data from Azure DevOps using the REST API and PowerShell, storing information locally in JSON and CSV formats for easy analysis and reporting.

## Features

- **No Cloud Dependencies**: Works entirely locally - no Azure Storage or SQL Database required
- **Comprehensive Data Collection**: Captures users, groups, projects, work items, git commits, pull requests, builds, and releases
- **Multiple Output Formats**: Saves data in both JSON and CSV formats
- **HTML Reporting**: Generates beautiful HTML reports with statistics and charts
- **Progress Tracking**: Real-time progress indicators during data collection
- **Error Handling**: Robust error handling with detailed logging

## Collected Data Types

1. **Organization Statistics** - High-level metrics about your organization
2. **Project Statistics** - Detailed statistics for each project
3. **Users** - User accounts, licenses, and access levels
4. **Groups** - Security groups and team information  
5. **Group Memberships** - User-to-group relationships
6. **Work Items** - User stories, tasks, bugs, features, and epics
7. **Git Pull Requests** - Code review data across all repositories
8. **Git Commits** - Commit history and author information

## Prerequisites

- PowerShell 5.1 or later
- Azure DevOps Personal Access Token (PAT)
- Access to Azure DevOps organization

## Setup

### 1. Personal Access Token

Create a Personal Access Token in Azure DevOps with the following permissions:
- **Build**: Read
- **Code**: Read  
- **Graph**: Read
- **Project and Team**: Read
- **Release**: Read
- **Work Items**: Read

### 2. Environment Configuration

1. Copy `Set-Environment-Local.ps1` and edit it:
```powershell
# Set your organization name and PAT
$env:ADOS_ORGANIZATION = 'your-organization-name'
$env:ADOS_PAT = 'your-personal-access-token'
```

2. Run the environment script:
```powershell
.\Set-Environment-Local.ps1
```

## Usage

### Basic Data Collection

Run the main collection script:
```powershell
.\Collect-Statistics-Local.ps1
```

### Advanced Options

```powershell
# Specify custom output path
.\Collect-Statistics-Local.ps1 -OutputPath "C:\MyData\DevOpsStats"

# Save only JSON format
.\Collect-Statistics-Local.ps1 -Format "JSON"

# Save only CSV format  
.\Collect-Statistics-Local.ps1 -Format "CSV"

# Save both formats (default)
.\Collect-Statistics-Local.ps1 -Format "Both"
```

### Generate Reports

After data collection, generate an HTML report:
```powershell
# Generate and open HTML report
.\Generate-Report.ps1 -OpenReport

# Generate JSON report
.\Generate-Report.ps1 -OutputFormat "JSON"

# Use specific data path
.\Generate-Report.ps1 -DataPath "C:\MyData\DevOpsStats\org\2024-11-05"
```

## Output Structure

Data is organized in the following structure:
```
data/
├── [organization-name]/
│   └── [date]/
│       ├── Collection-Summary.json
│       ├── OrganizationStatistics.json/csv
│       ├── ProjectStatistics.json/csv
│       ├── Users.json/csv
│       ├── Groups.json/csv
│       ├── GroupMemberships.json/csv
│       ├── WorkItems.json/csv
│       ├── GitPullRequests.json/csv
│       ├── GitCommits.json/csv
│       └── Azure-DevOps-Report.html
```

## File Descriptions

### Collection Scripts
- `Collect-Statistics-Local.ps1` - Main orchestrator script
- `Set-Environment-Local.ps1` - Environment configuration
- `Generate-Report.ps1` - Report generation

### Individual Data Collectors (in src/powershell/)
- `Users-Local.ps1` - Collects user account information
- `Groups-Local.ps1` - Collects security groups and teams
- `GroupMemberships-Local.ps1` - Collects group membership data
- `WorkItems-Local.ps1` - Collects work items across all projects
- `GitPullRequests-Local.ps1` - Collects pull request data
- `GitCommits-Local.ps1` - Collects Git commit history
- `OrganizationStatistics-Local.ps1` - Collects high-level org metrics
- `ProjectStatistics-Local.ps1` - Collects per-project statistics

## JSON Data Format

All JSON files include metadata and structured data:

```json
{
  "metadata": {
    "organization": "your-org",
    "collectionDate": "2024-11-05 14:30 +00:00",
    "recordCount": 150,
    "dataType": "Users"
  },
  "data": [...]
}
```

## Sample Workflows

### Daily Statistics Collection
```powershell
# Set up scheduled task or run manually
.\Set-Environment-Local.ps1
.\Collect-Statistics-Local.ps1
.\Generate-Report.ps1 -OpenReport
```

### Project-Specific Analysis
```powershell
# Collect work items for specific project
.\src\powershell\WorkItems-Local.ps1 -OutputPath ".\analysis" -ProjectName "MyProject"
```

### Data Analysis Examples

#### PowerShell Analysis
```powershell
# Load and analyze data
$users = Get-Content "data\org\2024-11-05\Users.json" | ConvertFrom-Json
$activeUsers = $users.data | Where-Object {$_.status -eq "active"}
Write-Host "Active users: $($activeUsers.Count)"
```

#### Excel Integration
Data can be imported into Excel for advanced analysis and visualization.

## Troubleshooting

### Common Issues

1. **Authentication Errors**
   - Verify PAT has correct permissions
   - Check organization name is correct
   - Ensure PAT hasn't expired

2. **Missing Data**
   - Some projects may have restricted access
   - Check user permissions in Azure DevOps
   - Review error messages in console output

3. **Performance**
   - Large organizations may take significant time
   - Consider running during off-peak hours
   - Monitor API rate limits

### Debug Mode

Enable verbose logging:
```powershell
.\Collect-Statistics-Local.ps1 -Verbose
```

## API Rate Limits

The tool respects Azure DevOps API rate limits and includes built-in retry logic. For large organizations, collection may take several hours.

## Data Privacy

All data is stored locally on your machine. No data is transmitted to external services except the Azure DevOps API calls required for collection.

## Comparison with Original Version

| Feature | Original | Local Version |
|---------|----------|---------------|
| Storage | Azure SQL Database | Local JSON/CSV files |
| Data Processing | Azure Functions | Local PowerShell |
| File Storage | Azure Blob Storage | Local file system |
| Reporting | Power BI | HTML reports |
| Dependencies | Azure services | PowerShell only |
| Cost | Azure resource costs | Free |

## License

This project maintains the same license as the original azure-devops-statistics project.