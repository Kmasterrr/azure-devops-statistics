/* 
-----------------------------------------------------------
  Create staging schema
-----------------------------------------------------------
*/

IF NOT EXISTS (SELECT [name] FROM sys.schemas WHERE [name] = N'staging')
BEGIN
    EXEC('CREATE SCHEMA [staging] AUTHORIZATION [dbo]');
END
GO


/* 
-----------------------------------------------------------
  Create staging tables
-----------------------------------------------------------
*/

DROP TABLE IF EXISTS [staging].[Users];
GO

CREATE TABLE [staging].[Users]
(
  [id] varchar(200) NULL,
  [descriptor] varchar(200) NULL,
  [principalName] varchar(200) NULL,
  [displayName] varchar(200) NULL,
  [email] varchar(200) NULL,
  [origin] varchar(200) NULL,
  [originId] varchar(200) NULL,
  [kind] varchar(200) NULL,
  [type] varchar(200) NULL,
  [domain] varchar(200) NULL,
  [status] varchar(200) NULL,
  [license] varchar(200) NULL,
  [licenseType] varchar(200) NULL,
  [source] varchar(200) NULL,
  [dateCreated] varchar(200) NULL,
  [lastAccessedDate] varchar(200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[Groups];
GO

CREATE TABLE [staging].[Groups]
(
    [id] [varchar](200) NULL,
    [principalName] [varchar](200) NULL,
    [displayName] [varchar](200) NULL,
    [description] [varchar](500) NULL,
    [origin] [varchar](200) NULL,
    [originId] [varchar](200) NULL,
    [domain] [varchar](200) NULL,
    [subjectKind] [varchar](200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[GroupMemberships];
GO

CREATE TABLE [staging].[GroupMemberships]
(
  [groupId] [varchar](200) NULL,
  [memberId] [varchar](200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[GitPullRequests];
GO

CREATE TABLE [staging].[GitPullRequests]
(
  [pullRequestId] [varchar](200) NULL,
  [status] [varchar](200) NULL,
  [title] [varchar](200) NULL,
  [sourceBranch] [varchar](200) NULL,
  [targetBranch] [varchar](200) NULL,
  [projectId] [varchar](200) NULL,
  [projectName] [varchar](200) NULL,
  [repositoryId] [varchar](200) NULL,
  [repositoryName] [varchar](200) NULL,
  [userId] [varchar](200) NULL,
  [user] [varchar](200) NULL,
  [creationDate] [varchar](200) NULL,
  [closedDate] [varchar](200) NULL,
  [timestamp] [varchar](200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[GitCommits];
GO

CREATE TABLE [staging].[GitCommits]
(
  [commitId] [varchar](200) NULL,
  [date] [varchar](200) NULL,
  [author] [varchar](200) NULL,
  [email] [varchar](200) NULL,
  [repositoryId] [varchar](200) NULL,
  [repositoryName] [varchar](200) NULL,
  [defaultBranch] [varchar](200) NULL,
  [projectId] [varchar](200) NULL,
  [projectName] [varchar](200) NULL,
  [comment] [varchar](200) NULL,
  [timestamp] [varchar](200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[ProjectStatistics];
GO

CREATE TABLE [staging].[ProjectStatistics]
(
  [timeStamp] [varchar](200) NULL,
  [projectId] [varchar](200) NULL,
  [projectName] [varchar](200) NULL,
  [workItemsCreated] [varchar](200) NULL,
  [workItemsCompleted] [varchar](200) NULL,
  [commitsPushed] [varchar](200) NULL,
  [pullRequestsCreated] [varchar](200) NULL,
  [pullRequestsCompleted] [varchar](200) NULL,
  [builds] [varchar](200) NULL,
  [releases] [varchar](200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[OrganizationStatistics];
GO

CREATE TABLE [staging].[OrganizationStatistics]
(
  [Organization] [varchar](200) NULL,
  [TimeStamp] [varchar](200) NULL,
  [Projects] [varchar](200) NULL,
  [BuildPipelines] [varchar](200) NULL,
  [Builds] [varchar](200) NULL,
  [BuildsCompleted] [varchar](200) NULL,
  [BuildCompletionPercentage] [varchar](200) NULL,
  [ReleasePipelines] [varchar](200) NULL,
  [Releases] [varchar](200) NULL,
  [ReleasesToProduction] [varchar](200) NULL,
  [ReleasesCompleted] [varchar](200) NULL,
  [ReleaseCompletionPercentage] [varchar](200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[ProjectGroupsAndUsers];
GO

CREATE TABLE [staging].[ProjectGroupsAndUsers]
(
    [timeStamp] [datetime2] NULL,
    [projectName] [varchar](200) NULL,
    [groupName] [varchar](200) NULL,
    [principalName] [varchar](200) NULL,
    [displayName] [varchar](200) NULL,
    [origin] [varchar](200) NULL,
    [type] [varchar](200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[GitRepositoriesPermissions];
GO

CREATE TABLE [staging].[GitRepositoriesPermissions]
(
    [timestamp] [datetime2] NULL,
    [projectId] [varchar](200) NULL,
    [projectName] [varchar](200) NULL,
    [repoId] [varchar](200) NULL,
    [repoName] [varchar](200) NULL,
    [securityNameSpaceId] [varchar](200) NULL,
    [securityNameSpace] [varchar](200) NULL,
    [groupDomain] [varchar](200) NULL,
    [groupDisplayName] [varchar](200) NULL,
    [groupAccountName] [varchar](200) NULL,
    [gitCommandName] [varchar](200) NULL,
    [gitCommandInternalName] [varchar](200) NULL,
    [gitCommandPermission] [varchar](200) NULL
);
GO

---------------------------------------

DROP TABLE IF EXISTS [staging].[ProjectLevelPermissions];
GO

CREATE TABLE [staging].[ProjectLevelPermissions]
(
  [timestamp] [datetime2] NULL,
  [teamProjectName] [varchar](200) NULL,
  [securityNameSpace] [varchar](200) NULL,
  [userPrincipalName] [varchar](200) NULL,
  [userDisplayName] [varchar](200) NULL,
  [groupDisplayName] [varchar](200) NULL,
  [groupAccountName] [varchar](200) NULL,
  [projectLevelType] [varchar](200) NULL,
  [projectLevelCommandName] [varchar](200) NULL,
  [projectLevelCommandInternalName] [varchar](200) NULL,
  [projectLevelCommandPermission] [varchar](200) NULL
);
GO


/* 
-----------------------------------------------------------
  Create data tables
-----------------------------------------------------------
*/

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'Users')
BEGIN
    CREATE TABLE [dbo].[Users]
    (
        [Id] [varchar](50) NOT NULL,
        [PrincipalName] [varchar](200) NOT NULL,
        [DisplayName] [varchar](200) NOT NULL,
        [ProfileImageUrl] [varchar](200) NULL,
        [AccessLevel] [varchar](200) NULL,
        [Source] [varchar](200) NULL,
        [Status] [varchar](200) NULL,
        [Origin] [varchar](200) NULL,
        [OriginId] [varchar](200) NULL,
        [DateCreated] [datetime2] NULL,
        [LastAccessedDate] [datetime2] NULL,
        [Timestamp] [datetime2] NULL
    );
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'Groups')
BEGIN
  CREATE TABLE [dbo].[Groups]
  (
      [Id] [varchar](200) NOT NULL,
      [PrincipalName] [varchar](200) NOT NULL,
      [DisplayName] [varchar](200) NOT NULL,
      [Description] [varchar](500) NULL,
      [Origin] [varchar](200) NOT NULL,
      [OriginId] [varchar](200) NOT NULL,
      [Domain] [varchar](200) NULL,
      [SubjectKind] [varchar](200) NULL
  );
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'GroupMemberships')
BEGIN
  CREATE TABLE [dbo].[GroupMemberships]
  (
      [GroupId] [varchar](200) NOT NULL,
      [MemberId] [varchar](200) NOT NULL
  );
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'GitPullRequests')
BEGIN
  CREATE TABLE [dbo].[GitPullRequests]
  (
      [PullRequestId] [int] NOT NULL,
      [Status] [varchar](200) NULL,
      [Title] [varchar](2000) NULL,
      [SourceBranch] [varchar](200) NOT NULL,
      [TargetBranch] [varchar](200) NOT NULL,
      [ProjectId] [varchar](50) NOT NULL,
      [Project] [varchar](200) NOT NULL,
      [RepositoryId] [varchar](50) NOT NULL,
      [Repository] [varchar](200) NOT NULL,
      [UserId] [varchar](50) NOT NULL,
      [User] [varchar](200) NOT NULL,
      [CreationDate] [datetime2] NOT NULL,
      [ClosedDate] [datetime2] NULL,
      [Timestamp] [datetime2] NOT NULL
  );
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'GitCommits')
BEGIN
  CREATE TABLE [dbo].[GitCommits]
  (
      [CommitId] [varchar](50) NOT NULL,
      [Comment] [varchar](2000) NULL,
      [Branch] [varchar](200) NOT NULL,
      [ProjectId] [varchar](50) NOT NULL,
      [Project] [varchar](200) NOT NULL,
      [RepositoryId] [varchar](50) NOT NULL,
      [Repository] [varchar](200) NOT NULL,
      [Author] [varchar](200) NOT NULL,
      [AuthorEmail] [varchar](200) NOT NULL,
      [CreationDate] [datetime2] NOT NULL,
      [Timestamp] [datetime2] NOT NULL
  );
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'ProjectStatistics')
BEGIN
  CREATE TABLE [dbo].[ProjectStatistics]
  (
      [Timestamp] [datetime2] NULL,
      [TeamProjectName] [varchar](200) NOT NULL,
      [TeamProjectCountWorkItemCreated] [smallint] NOT NULL,
      [TeamProjectCountWorkItemCompleted] [smallint] NOT NULL,
      [TeamProjectCountCommitsPushed] [smallint] NOT NULL,
      [TeamProjectCountPRsCreated] [smallint] NOT NULL,
      [TeamProjectCountPRsCompleted] [smallint] NOT NULL,
      [TeamProjectCountBuilds] [smallint] NOT NULL,
      [TeamProjectCountReleases] [smallint] NOT NULL
  );
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'OrganizationStatistics')
BEGIN
  CREATE TABLE [dbo].[OrganizationStatistics]
  (
      [OrganizationName] [varchar](100) NULL,
      [Timestamp] [datetime2] NULL,
      [Projects] [int] NULL,
      [BuildPipelines] [int] NULL,
      [Builds] [int] NULL,
      [BuildsCompleted] [int] NULL,
      [BuildCompletionPercentage] [varchar](10) NULL,
      [ReleasePipelines] [int] NULL,
      [Releases] [int] NULL,
      [ReleasesToProduction] [int] NULL,
      [ReleasesCompleted] [int] NULL,
      [ReleaseCompletionPercentage] [varchar](10) NULL
  ) ON [PRIMARY];
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'GitRepositoriesPermissions')
BEGIN
  CREATE TABLE [dbo].[GitRepositoriesPermissions]
  (
      [Timestamp] [datetime2] NULL,
      [TeamProjectName] [varchar](100) NULL,
      [RepoName] [varchar](100) NULL,
      [SecurityNameSpace] [varchar](100) NULL,
      [GroupDomain] [varchar](200) NULL,
      [GroupDisplayName] [varchar](200) NULL,
      [GroupAccountName] [varchar](200) NULL,
      [GitCommandName] [varchar](100) NULL,
      [GitCommandInternalName] [varchar](100) NULL,
      [GitCommandPermission] [varchar](50) NULL
  ) ON [PRIMARY];

  CREATE CLUSTERED INDEX [ix_ProjectRepoGitCommandName] 
    ON [dbo].[GitRepositoriesPermissions] 
    (
      [TeamProjectName], 
      [RepoName],
      [GitCommandName]
    );
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'ProjectGroupsAndUsers')
BEGIN
  CREATE TABLE [dbo].[ProjectGroupsAndUsers]
  (
      [Timestamp] [datetime2] NULL,
      [ProjectName] [varchar](200) NOT NULL,
      [GroupName] [varchar](200) NOT NULL,
      [PrincipalName] [varchar](200) NOT NULL,
      [DisplayName] [varchar](200) NOT NULL,
      [Origin] [varchar](200) NOT NULL,
      [Type] [varchar](200) NOT NULL
  );

  CREATE CLUSTERED INDEX [ix_GroupNameDisplayName]
    ON [dbo].[ProjectGroupsAndUsers]
    (
      [GroupName],
      [DisplayName]
    );
END
GO

---------------------------------------

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'ProjectLevelPermissions')
BEGIN
  CREATE TABLE [dbo].[ProjectLevelPermissions]
  (
    [Timestamp] [datetime2] NULL,
    [TeamProjectName] [varchar](100) NULL,
    [SecurityNameSpace] [varchar](100) NULL,
    [UserPrincipalName] [varchar](100) NULL,
    [UserDisplayName] [varchar](100) NULL,
    [GroupDisplayName] [varchar](200) NULL,
    [GroupAccountName] [varchar](200) NULL,
    [ProjectLevelType] [varchar](50) NULL,
    [ProjectLevelCommandName] [varchar](100) NULL,
    [ProjectLevelCommandInternalName] [varchar](100) NULL,
    [ProjectLevelCommandPermission] [varchar](50) NULL
  );

  CREATE CLUSTERED INDEX [ix_ProjectLevelCommandPermission]
    ON [dbo].[ProjectLevelPermissions]
    (
      [ProjectLevelCommandPermission]
    );
END
GO


/* 
-----------------------------------------------------------
  Create process tracking tables
-----------------------------------------------------------
*/

DROP TABLE IF EXISTS [dbo].[BulkImports];
GO

CREATE TABLE [dbo].[BulkImports]
(
  [Id] INT IDENTITY NOT NULL PRIMARY KEY,
  [FileName] NVARCHAR(200) NOT NULL,
  [Status] VARCHAR(50) NULL,
  [ImportDate] DATETIME2 NOT NULL CONSTRAINT [DF_ImportDate] DEFAULT (GETDATE())
);
GO

---------------------------------------

DROP TABLE IF EXISTS [dbo].[BulkImportHistory];
GO

CREATE TABLE [dbo].[BulkImportHistory]
(
  [Id] INT IDENTITY NOT NULL PRIMARY KEY,
  [ImportId] INT NOT NULL REFERENCES [dbo].[BulkImports](Id) ON DELETE CASCADE,
  [TimeStamp] DATETIME2 NOT NULL CONSTRAINT [DF_TimeStamp] DEFAULT (GETDATE()),
  [Type] NVARCHAR(50) NOT NULL,
  [Message] NVARCHAR(MAX) NOT NULL
);
GO


/* 
-----------------------------------------------------------
  Bulk Insert Stored Procedure
-----------------------------------------------------------
*/

CREATE OR ALTER PROCEDURE [dbo].[BulkLoadFromAzure]
(
  @sourceFileName NVARCHAR(100),
  @ForceReprocess BIT = 0
)
AS
BEGIN
  SET NOCOUNT ON;

  -- Error handling variables 
  DECLARE @VAR_ERR_NUM INT; 
  DECLARE @VAR_ERR_LINE INT; 
  DECLARE @VAR_ERR_MSG VARCHAR(1024); 
  DECLARE @VAR_ERR_PROC VARCHAR(1024); 

  DECLARE @fid INT;
	DECLARE @fileName NVARCHAR(MAX);
  DECLARE @FNamePos INT = CHARINDEX('/', REVERSE(@sourceFileName));
  DECLARE @ExtPos INT = CHARINDEX('.', REVERSE(@sourceFileName));
  DECLARE @ExtlDS varchar(50) = 'Azure-Storage';

  SELECT @fileName = SUBSTRING(@sourceFileName, LEN(@sourceFileName) - @FNamePos + 2, @FnamePos - @ExtPos - 1);

  INSERT INTO [dbo].[BulkImports] ([FileName], [Status])
  VALUES (@sourceFileName, 'Queued');

  SET @fid = (SELECT SCOPE_IDENTITY());

  IF (@ForceReprocess = 1)
  BEGIN
    INSERT INTO [dbo].[BulkImportHistory] ([ImportId], [Type], [Message])
    VALUES (@fid, 'Info', 'Queued for reporcessing.');
  END

  INSERT INTO [dbo].[BulkImportHistory] ([ImportId], [Type], [Message])
  VALUES (@fid, 'Info', 'Started bulk import process.');

  -- ** ERROR HANDLING - START TRY ** 
  BEGIN TRY

    UPDATE [dbo].[BulkImports] SET [Status] = 'Processing' WHERE [Id] = @fid;

    -- Truncate staging table
    INSERT INTO [dbo].[BulkImportHistory] ([ImportId], [Type], [Message])
    VALUES (@fid, 'Info', 'Truncating staging table.');

    DECLARE @TruncateStmt NVARCHAR(MAX);
    SET @TruncateStmt = N'
      TRUNCATE TABLE [staging].[' + @fileName + '];
    ';
    EXEC (@TruncateStmt);

    -- Bulk import into staging table
    DECLARE @BulkInsSQL NVARCHAR(MAX);
    SET @BulkInsSQL = N'
      BULK INSERT [staging].[' + @fileName + ']
      FROM ''' + @sourceFileName + '''
      WITH 
      ( 
        DATA_SOURCE = ''' + @ExtlDS + ''',
        FORMAT = ''CSV'',
        FIELDQUOTE = ''"'',
        FIELDTERMINATOR = '','',
        FIRSTROW = 2,
        ROWTERMINATOR = ''0x0a'',
        TABLOCK
      );
      ';

    EXEC (@BulkInsSQL);

    -- Get count of records imported to staging table
    DECLARE @Count INT, @SqlCmd NVARCHAR(MAX);
    SET @SqlCmd = N'SELECT @cnt=COUNT(*) FROM [staging].[' + @fileName + ']';
    EXECUTE sp_executesql @SqlCmd, N'@cnt INT OUTPUT', @cnt = @Count OUTPUT;

    INSERT INTO [dbo].[BulkImportHistory] ([ImportId], [Type], [Message])
    VALUES (@fid, 'Info', 'Bulk Import successful. Records Imported: ' + CAST(@Count as varchar));

    -- Process staging table into data table
    EXEC [dbo].[ProcessStagingTable] @TableName = @fileName, @ImportId = @fid

    UPDATE [dbo].[BulkImports] SET [Status] = 'Completed' WHERE [Id] = @fid;

  END TRY
  -- ** ERROR HANDLING - END TRY ** 

  -- ** Error Handling - Begin Catch ** 
  BEGIN CATCH

    -- Grab variables 
    SELECT 
      @VAR_ERR_NUM  = ERROR_NUMBER(),
      @VAR_ERR_PROC = ERROR_PROCEDURE(),
      @VAR_ERR_LINE = ERROR_LINE(),
      @VAR_ERR_MSG  = ERROR_MESSAGE();

    -- Log Error
    INSERT INTO [dbo].[BulkImportHistory] ([ImportId], [Type], [Message])
    VALUES (@fid, 'Error', @VAR_ERR_MSG);

    UPDATE [dbo].[BulkImports] SET [Status] = 'Error' WHERE [Id] = @fid;

    -- Raise error 
    RAISERROR ('An error occurred within a user transaction. 
                Error Number        : %d 
                Error Message       : %s 
                Affected Procedure  : %s 
                Affected Line Number: %d'
                , 16, 1 
                , @VAR_ERR_NUM, @VAR_ERR_MSG, @VAR_ERR_PROC, @VAR_ERR_LINE);
  END CATCH
  -- ** Error Handling - End Catch **

END
GO

---------------------------------------

CREATE OR ALTER PROCEDURE [dbo].[ProcessStagingTable]
(
  @ImportId INT,
  @TableName NVARCHAR(100)
)
AS
BEGIN
  SET NOCOUNT ON;

  -- Error handling variables 
  DECLARE @VAR_ERR_NUM INT; 
  DECLARE @VAR_ERR_LINE INT; 
  DECLARE @VAR_ERR_MSG VARCHAR(1024); 
  DECLARE @VAR_ERR_PROC VARCHAR(1024); 

  DECLARE @Processed BIT = 0;

  -- ** ERROR HANDLING - START TRY ** 
  BEGIN TRY

    IF (@TableName = 'Users')
    BEGIN
      TRUNCATE TABLE [dbo].[Users];

      INSERT INTO [dbo].[Users]
      (
        [Id],
        [PrincipalName],
        [DisplayName],
        [AccessLevel],
        [Source],
        [Status],
        [Origin],
        [OriginId],
        [DateCreated],
        [LastAccessedDate],
        [Timestamp]
      )
      SELECT 
        [id], 
        [principalName],
        [displayName],
        [license],
        [source],
        [status],
        [origin],
        [originId],
        [dateCreated],
        [lastAccessedDate],
        GETDATE()
      FROM [staging].[Users];

      SET @Processed = 1;
    END

    IF (@TableName = 'Groups')
    BEGIN
      TRUNCATE TABLE [dbo].[Groups];

      INSERT INTO [dbo].[Groups] 
      (
          [Id],
          [PrincipalName],
          [DisplayName],
          [Description],
          [Origin],
          [OriginId],
          [Domain],
          [SubjectKind]
      )
      SELECT
          [id],
          [principalName],
          [displayName],
          [description],
          [origin],
          [originId],
          [domain],
          [subjectKind]
      FROM [staging].[Groups];

      SET @Processed = 1;
    END

    IF (@TableName = 'GroupMemberships')
    BEGIN
      TRUNCATE TABLE [dbo].[GroupMemberships];

      INSERT INTO [dbo].[GroupMemberships]
      (
        [GroupId],
        [MemberId]
      )
      SELECT
        [groupId],
        [memberId]
      FROM [staging].[GroupMemberships];

      SET @Processed = 1;
    END

    IF (@TableName = 'GitPullRequests')
    BEGIN
      TRUNCATE TABLE [dbo].[GitPullRequests];

      INSERT INTO [dbo].[GitPullRequests]
      (
        [PullRequestId],
        [Status],
        [Title],
        [SourceBranch],
        [TargetBranch],
        [ProjectId],
        [Project],
        [RepositoryId],
        [Repository],
        [UserId],
        [User],
        [CreationDate],
        [ClosedDate],
        [Timestamp]
      )
      SELECT
        CAST([pullRequestId] AS INT),
        [status],
        [title],
        [sourceBranch],
        [targetBranch],
        [projectId],
        [projectName],
        [repositoryId],
        [repositoryName],
        [userId],
        [user],
        CAST([creationDate] AS DATETIME2),
        CAST([closedDate] AS DATETIME2),
        CAST([timestamp] AS DATETIME2)
      FROM [staging].[GitPullRequests];

      SET @Processed = 1;
    END

    IF (@TableName = 'GitCommits')
    BEGIN
      TRUNCATE TABLE [dbo].[GitCommits];

      INSERT INTO [dbo].[GitCommits]
      (
        [CommitId],
        [Comment],
        [Branch],
        [ProjectId],
        [Project],
        [RepositoryId],
        [Repository],
        [Author],
        [AuthorEmail],
        [CreationDate],
        [Timestamp]
      )
      SELECT
        [commitId],
        [comment],
        [defaultBranch],
        [projectId],
        [projectName],
        [repositoryId],
        [repositoryName],
        [author],
        [email],
        CAST([date] AS DATETIME2),
        CAST([timestamp] AS DATETIME2)
      FROM [staging].[GitCommits];

      SET @Processed = 1;
    END

    IF (@TableName = 'ProjectStatistics')
    BEGIN
      -- Remove existing records if reporcessing the same file.
      DECLARE @PSDate DATETIME2;
      SELECT @PSDate = CAST([timeStamp] AS DATETIME2) FROM [staging].[ProjectStatistics];

			IF EXISTS (SELECT TOP(1) [TimeStamp] FROM [dbo].[ProjectStatistics] WHERE [TimeStamp] = @PSDate)
			BEGIN
				DELETE FROM [dbo].[ProjectStatistics] WHERE [TimeStamp] = @PSDate;
			END

      INSERT INTO [dbo].[ProjectStatistics]
      (
        [Timestamp],
        [TeamProjectName],
        [TeamProjectCountWorkItemCreated],
        [TeamProjectCountWorkItemCompleted],
        [TeamProjectCountCommitsPushed],
        [TeamProjectCountPRsCreated],
        [TeamProjectCountPRsCompleted],
        [TeamProjectCountBuilds],
        [TeamProjectCountReleases]
      )
      SELECT
        CAST([timeStamp] AS DATETIME2),
        [projectName],
        CAST([workItemsCreated] AS SMALLINT),
        CAST([workItemsCompleted] AS SMALLINT),
        CAST([commitsPushed] AS SMALLINT),
        CAST([pullRequestsCreated] AS SMALLINT),
        CAST([pullRequestsCompleted] AS SMALLINT),
        CAST([builds] AS SMALLINT),
        CAST([releases] AS SMALLINT)
      FROM [staging].[ProjectStatistics]

      SET @Processed = 1;
    END

    IF (@TableName = 'OrganizationStatistics')
    BEGIN
      -- Remove existing records if reporcessing the same file.
      IF EXISTS(SELECT CAST([TimeStamp] AS DATETIME2) FROM [staging].[OrganizationStatistics])
      BEGIN
        DELETE FROM [dbo].[OrganizationStatistics]
        WHERE [TimeStamp] IN (SELECT CAST([TimeStamp] AS DATETIME2) FROM [staging].[OrganizationStatistics]);
      END

      INSERT INTO [dbo].[OrganizationStatistics]
      (
        [OrganizationName],
        [Timestamp],
        [Projects],
        [BuildPipelines],
        [Builds],
        [BuildsCompleted],
        [BuildCompletionPercentage],
        [ReleasePipelines],
        [Releases],
        [ReleasesToProduction],
        [ReleasesCompleted],
        [ReleaseCompletionPercentage]
      )
      SELECT
        [Organization],
        CAST([TimeStamp] AS DATETIME2),
        CAST([Projects] AS INT),
        CAST([BuildPipelines] AS INT),
        CAST([Builds] AS INT),
        CAST([BuildsCompleted] AS INT),
        [BuildCompletionPercentage],
        CAST([ReleasePipelines] AS INT),
        CAST([Releases] AS INT),
        CAST([ReleasesToProduction] AS INT),
        CAST([ReleasesCompleted] AS INT),
        [ReleaseCompletionPercentage]
      FROM [staging].[OrganizationStatistics]

      SET @Processed = 1;
    END

    IF (@TableName = 'ProjectGroupsAndUsers')
    BEGIN
      TRUNCATE TABLE [dbo].[ProjectGroupsAndUsers];

      INSERT INTO [dbo].[ProjectGroupsAndUsers]
      (
        [Timestamp],
        [ProjectName],
        [GroupName],
        [PrincipalName],
        [DisplayName],
        [Origin],
        [Type]
      )
      SELECT
        [timeStamp],
        [projectName],
        [groupName],
        [principalName],
        [displayName],
        [origin],
        [type]
      FROM [staging].[ProjectGroupsAndUsers];

      SET @Processed = 1;
    END

    IF (@TableName = 'GitRepositoriesPermissions')
    BEGIN
      TRUNCATE TABLE [dbo].[GitRepositoriesPermissions];

      INSERT INTO [dbo].[GitRepositoriesPermissions]
      (
        [Timestamp],
        [TeamProjectName],
        [RepoName],
        [SecurityNameSpace],
        [GroupDomain],
        [GroupDisplayName],
        [GroupAccountName],
        [GitCommandName],
        [GitCommandInternalName],
        [GitCommandPermission]
      )
      SELECT
        [timestamp],
        [projectName],
        [repoName],
        [securityNameSpace],
        [groupDomain],
        [groupDisplayName],
        [groupAccountName],
        [gitCommandName],
        [gitCommandInternalName],
        [gitCommandPermission]
      FROM [staging].[GitRepositoriesPermissions];

      SET @Processed = 1;
    END

    IF (@TableName = 'ProjectLevelPermissions')
    BEGIN
      TRUNCATE TABLE [dbo].[ProjectLevelPermissions];

      INSERT INTO [dbo].[ProjectLevelPermissions]
      (
        [Timestamp],
        [TeamProjectName],
        [SecurityNameSpace],
        [UserPrincipalName],
        [UserDisplayName],
        [GroupDisplayName],
        [GroupAccountName],
        [ProjectLevelType],
        [ProjectLevelCommandName],
        [ProjectLevelCommandInternalName],
        [ProjectLevelCommandPermission]
      )
      SELECT
        [timestamp],
        [teamProjectName],
        [securityNameSpace],
        [userPrincipalName],
        [userDisplayName],
        [groupDisplayName],
        [groupAccountName],
        [projectLevelType],
        [projectLevelCommandName],
        [projectLevelCommandInternalName],
        [projectLevelCommandPermission]
      FROM [staging].[ProjectLevelPermissions];

      SET @Processed = 1;
    END

    -- Log process success/fail
    IF (@Processed = 1)
    BEGIN
      INSERT INTO [dbo].[BulkImportHistory] ([ImportId], [Type], [Message])
      VALUES (@ImportId, 'Info', 'Processed staging table: ' + @TableName);
    END
    ELSE
    BEGIN
      INSERT INTO [dbo].[BulkImportHistory] ([ImportId], [Type], [Message])
      VALUES (@ImportId, 'Error', 'Could not process staging table: ' + @TableName);
    END

  END TRY
  -- ** ERROR HANDLING - END TRY ** 

  -- ** Error Handling - Begin Catch ** 
  BEGIN CATCH

    -- Grab variables 
    SELECT 
      @VAR_ERR_NUM  = ERROR_NUMBER(),
      @VAR_ERR_PROC = ERROR_PROCEDURE(),
      @VAR_ERR_LINE = ERROR_LINE(),
      @VAR_ERR_MSG  = ERROR_MESSAGE();

    -- Log Error
    INSERT INTO [dbo].[BulkImportHistory] ([ImportId], [Type], [Message])
    VALUES (@ImportId, 'Error', @VAR_ERR_MSG);

    UPDATE [dbo].[BulkImports] SET [Status] = 'Error' WHERE [Id] = @ImportId;

    -- Raise error 
    RAISERROR ('An error occurred within a user transaction. 
                Error Number        : %d 
                Error Message       : %s 
                Affected Procedure  : %s 
                Affected Line Number: %d'
                , 16, 1 
                , @VAR_ERR_NUM, @VAR_ERR_MSG, @VAR_ERR_PROC, @VAR_ERR_LINE);
  END CATCH
  -- ** Error Handling - End Catch **

END
GO
