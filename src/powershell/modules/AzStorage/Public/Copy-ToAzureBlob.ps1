function Copy-ToAzureBlob
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String] $uploadRootDirectory,

        [Parameter(Mandatory = $false)]
        [string] $skipDirectories,

        [Parameter(Mandatory = $true)]
        [string] $storageAccountName,

        [Parameter(Mandatory = $true)]
        [string] $storageAccountKey,

        [Parameter(Mandatory = $true)]
        [string] $containerName,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Hot","Cool","Archive")] 
        [string] $blobTier = "Hot",

        [Parameter(Mandatory = $false)]
        [int] $gracePeriodInSeconds = 0
    )

    begin
    {
        function DirectoryPathIsInSkipDirectories($directory, $skipDirectories)
        {
            foreach ($skipDirectory in $skipDirectories)
            {
                if ($directory -like $skipDirectory)
                {
                    return $true
                }
            }

            return $false
        }

        function UploadDirectory($uploadRootDirectory, $directory, $skipDirectories, $storageContext, $containerName, $filesOlderThan, $blobTier)
        {
            # Return if we need to skip this directory
            if (DirectoryPathIsInSkipDirectories -directory $directory -skipDirectories $skipDirectories)
            {
                return;
            }
        
            # First recurse through subdirectories
            $subDirectories = Get-ChildItem -Path $directory -Directory
            
            foreach ($subDirectory in $subDirectories) 
            {
                UploadDirectory -uploadRootDirectory $uploadRootDirectory -directory $subDirectory -skipDirectories $skipDirectories -storageContext $storageContext -containerName $containerName -filesOlderThan $filesOlderThan -blobTier $blobTier
            }
        
            # Get all files and see if we need to upload them.
            $files = Get-ChildItem -Path $directory -File | Where-Object {$_.Lastwritetime -lt $filesOlderThan}
        
            foreach ($file in $files) 
            {
                Write-Verbose "Uploading $file"
                $relativePath = $file.Directory | Resolve-Path -Relative
                $blobFilename = Join-Path $relativePath $file.Name
                $uploadResult = Set-AzStorageBlobContent -File $file -Container $containerName -Blob $blobFilename -Context $storageContext -Force -StandardBlobTier $blobTier
        
                if($uploadResult)
                {
                    Write-Verbose "Uploaded $file"
                }
                else
                {
                    Write-Error "Something went wrong uploading $file"
                }
            }
        
            # Check if directory is empty. If yes: delete.
            if ((Get-ChildItem $directory | Measure-Object).count -eq 0)
            {
                Remove-Item $directory -Force
                Write-Verbose "Deleted empty directory: $directory"
            }
        }
    }

    process
    {
        try
		{
			$VerbosePreference = 'Continue'

            # Check if the local path exists
            if (!(Test-Path -Path $uploadRootDirectory))
            {
                Write-Error "$uploadRootDirectory not found!"
                return
            }

            $skipDirectories = $skipDirectories.Split(';')
            $cutOffDate = (Get-date).AddSeconds($gracePeriodInSeconds * -1)
            # Set-Location $uploadRootDirectory

			# Check if the Azure modules are loaded
			if ((Get-Module -Name Az.Storage -ListAvailable).Count -le 0)
			{
				# Azure Storage module is not available Exit script
				Write-Error "ERROR: The Azure module is not available, exiting script"
				Write-Warning "Please download the Azure.Storage PowerShell modules from https://www.powershellgallery.com/"
				return
			}
			else
			{
				Write-Verbose "The Azure module is available and loaded..."
			}
			
			# details about Service Principal account. This part will authenticate to Azure.
			$u = $env:AZURE_APPLICATION_ID # application ID
			$key = $env:AZURE_CLIENT_KEY # Key
			$tenantid = $env:AZURE_TENANT_ID
			$pass = ConvertTo-SecureString $key -AsPlainText -Force
			$cred = New-Object -TypeName pscredential -ArgumentList $u, $pass
			Connect-AzAccount -Credential $cred -ServicePrincipal -TenantId $tenantid
			
			# Initiate the Azure Storage Context
			$storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey
			
			# # Check if the defined container already exists
			Write-Verbose "Checking availability of Azure container `"$containerName`""
			$azcontainer = Get-AzStorageContainer -Name $containerName -Context $storageContext -ErrorAction SilentlyContinue
			
            if ($null -eq $azcontainer)
            {
                # Container doesn't exist, create a new one
                Write-Verbose "Container `"$containerName`" does not exist, trying to create container"
                $azcontainer = New-AzStorageContainer -Name $containerName -Context $storageContext -ErrorAction stop
            }

			# Retrieve the files in the given folders
            Write-Host "Copy $uploadRootDirectory"
            UploadDirectory -uploadRootDirectory $uploadRootDirectory -directory $uploadRootDirectory -skipDirectories $skipDirectories -storageContext $storageContext -containerName $containerName -filesOlderThan $cutOffDate -blobTier $blobTier
		}
        # Catch specific types of exceptions thrown by one of those commands
        catch [System.Net.WebException], [System.Exception] 
        {
            $PSCmdlet.ThrowTerminatingError($PSItem)
		}
    }

    end
    {
    }
}
