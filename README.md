# Azure DevOps Statistics (ADOS)

Data extraction scripts and Power BI reports that generate information about your Azure DevOps organization. Using Azure DevOps CLI and Azure DevOps REST API, PowerShell scripts extract data from Azure DevOps, store this information in an Azure SQL Database that is used by Power BI reports.

## Quick Start - Azure DevOps Pipeline

The easiest way to use this tool is through the Azure DevOps Pipeline that collects statistics and publishes an HTML report as an artifact.

### Option 1: Self-Organization Pipeline (Recommended)

Use `azure-pipelines-self.yaml` to collect statistics for the **current organization** where the pipeline runs. This uses the built-in `System.AccessToken` - no PAT required!

1. Create a new pipeline in Azure DevOps pointing to `azure-pipelines-self.yaml`
2. Grant the pipeline access to read organization data (Project Collection Build Service needs permissions)
3. Run the pipeline
4. View the HTML report in the pipeline artifacts

### Option 2: External Organization Pipeline

Use `azure-pipelines-statistics.yaml` to collect statistics for **any organization** (requires PAT):

1. Create a new pipeline in Azure DevOps pointing to `azure-pipelines-statistics.yaml`
2. Add a secret variable named `ADOS_PAT` with a Personal Access Token that has:
   - **Code**: Read
   - **Build**: Read  
   - **Work Items**: Read
   - **Graph**: Read
   - **Project and Team**: Read
   - **Member Entitlement Management**: Read
3. Run the pipeline with the organization name parameter
4. View the HTML report in the pipeline artifacts

### Viewing the HTML Report

After the pipeline runs:
1. Click **"Artifacts"** button in the pipeline run summary
2. Download the `AzureDevOpsStatisticsReport` artifact
3. Open `Azure-DevOps-Report.html` in your browser

The report includes:
- Organization overview statistics
- Project-level metrics (repos, commits, pipelines, builds, work items)
- User license distribution
- Top contributors leaderboard
- Repository details

---

## Project Architecture

1. PowerShell Scripts
2. Azure blob storage container
3. Azure function
4. Azure SQL database
5. PowerBI reports

![architecture](docs/ados-architecture.png)

## Azure Blob Storage

Create a storage account in Azure and create a blob container named '**devops-stats**' within that storage account.

## Azure Function

A blob triggered azure function that is invoked when a file is uploaded in the azure storage container.

```c#
        [FunctionName("FileProcessor")]
        public static void Run([BlobTrigger("devops-stats/{name}", Connection = "AzureStorage")]Stream blob, string name, ILogger log)
        {
            log.LogInformation($"Blob trigger function processed blob: {name}, size: {blob.Length} bytes");

            if (!name.EndsWith(".csv"))
            {
                log.LogInformation($"Blob '{name}' doesn't have the .csv extension. Skipping processing.");
                return;
            }

            log.LogInformation($"Blob '{name}' found. Uploading to Azure SQL");

            string azureSQLConnectionString = Environment.GetEnvironmentVariable("AzureSQLConnStr");

            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection(azureSQLConnectionString);
                conn.Execute("EXEC dbo.BulkLoadFromAzure @sourceFileName", new { @sourceFileName = name }, commandTimeout: 180);
                log.LogInformation($"Blob '{name}' uploaded");
            }
            catch (SqlException se)
            {
                log.LogInformation($"Exception Trapped: {se.Message}");
            }
            finally
            {
                conn?.Close();
            }
        }
```

## Azure SQL database

A connection from Azure SQL Server is set up to the external source (here, the storage account).

Azure BULK INSERT is used in a stored procedure to load the CSV file data to azure sql db table. The filename is sent as an input parameter to the stored procedure.

Data is first loaded into a staging table. If data loading is successful, then an entry is made in a header table and then the data from staging is moved to the main table

## Setup

[Setup](docs/setup.md)

## How to use

1. Make sure to set all the [environment variables](docs/setup.md/#environment-variables).
2. Run PowerShell script.

### Example

```powershell
./Collect-Statistics.ps1
```
