# Setup

The following are the instructions you need to follow in order to setup the solution. Azure CLI 2.0 commands are also provided as a reference.

## Authenticate Azure CLI 2.0

Login with your account

```bash
az login
```

and then select the Subscription you want to use

```bash
az account set --subscription <your-subscription>
```

## Create a Resource Group

All resources of this example will be created in a dedicated Resource Group, named "CSVImportDemo".

```bash
az group create --name CSVImportDemo --location eastus
```

## Create Blob Storage

Create the azure storage account used by Azure Function and to drop in our CSV files to be automatically imported.

```bash
az storage account create --name csvimportdemo --location eastus --resource-group CSVImportDemo --sku Standard_LRS
```

Once this is done get the account key and create a container named 'devops-stats':

```bash
az storage container create --account-name csvimportdemo --account-key <your-account-key> --name devops-stats
```

Generate a Shared Access Signature (SAS) key token and store it for later use. The easiest way to do this is to use [Azure Storage Explorer](https://azure.microsoft.com/en-us/features/storage-explorer/).

Otherwise you can do it via AZ CLI using the `az storage account generate-sas` command.

## Create Azure SQL Server and Database

Create an Azure SQL Server:

```bash
az sql server create --name csvimportdemo --resource-group CSVImportDemo --location eastus --admin-user csvimportdemo --admin-password csvimportdemoPassw0rd!
```

Via the Azure Portal make sure that the firewall is configure to "Allow access to Azure Services".

Also a small Azure SQL Database:

```bash
az sql db create --name CSVImportDemo --resource-group CSVImportDemo --server csvimportdemo
```

## Configure Bulk Load Security

Connect to the created Azure SQL database and execute the script to configure access to blob store for Azure SQL. The script is available here

`src/sql/enable-bulk-insert.sql`

just customize it with your own info before running it.

Please note that when you specific the SAS key token, you have to remove the initial question mark (?) that is automatically added
when you create the SAS key token online.

## Create Database Objects

In the Azure SQL database couple of tables and a stored procedures needs also to be created in order to have the sample working correctly.

Scripts to create the mentioned objects are available in

`src/sql/create-objects.sql`

Just execute it against the Azure SQL database.

## Create and Deploy the function app

The easiest way to install, build and deploy the sample Function App, is to use [Visual Studio Code](https://code.visualstudio.com/). It will automatically detect that the `.csproj` is related to a Function App, will download the Function App runtime and also recommend you to download the Azure Function extension.

Once the project is loaded, add the `AzureSQL` configuration to your `local.settings.json` file:

```json
  "AzureSQL": "<sql-azure-connection-string>"
```

Also make sure that the Function App is correctly monitoring the Azure Storage Account where you plan to drop you CSV files. If you used Visual Studio code, this should have already been set up for you. If not make sure you have the 'AzureStorage' configuration element in your `local.settings.json` file and that it has the connection string for the Azure Blob Storage account you want to use:

```json
  "AzureStorage": "DefaultEndpointsProtocol=https;AccountName=csvimportdemo;AccountKey=[account-key-here];EndpointSuffix=core.windows.net"
```

## Deploy and Run the Function App

You can now run the Function app on your machine, or you can deploy using Visual Studio Code and its Azure Function extension. Or you can use `az functionapp` to deploy the function manually.

## Environment Variables

Setup the environment variables in order to run the PowerShell scripts.

* **ADOS_ORGANIZATION** : Your Azure DevOps organization (ex: "intermexteam")
* **ADOS_PAT** : Personal Access Token to access Azure DevOps APIs
* **AZURE_TENANT_ID** : Azure Service Principal Tenant ID
* **AZURE_APPLICATION_ID** : Azure Service Principal (ClientId) used to connect to Azure resources
* **AZURE_CLIENT_KEY** : Azure Service Principal key
* **AZURE_STORAGE_ACCOUNT** : Azure storage account name
* **AZURE_STORAGE_ACCOUNT_KEY** : Storage account access key
* **AZURE_STORAGE_CONTAINER** : Blob container name to put files into.
