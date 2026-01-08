using System;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage.Blob;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using Dapper;

namespace FileProcessor
{
    public static class FileProcessor
    {
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
    }
}