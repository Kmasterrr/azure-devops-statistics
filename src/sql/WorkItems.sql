CREATE TABLE [dbo].[WorkItems]
(
    [ProjectId] [varchar](50) NOT NULL,
    [Project] [varchar](200) NOT NULL,
    [WorkItemId] [int] NOT NULL,
    [AreaPath] [varchar](200) NULL,
    [IterationPath] [varchar](200) NULL,
    [ValueArea] [varchar](200) NULL,
    [WorkItemType] [varchar](50) NOT NULL,
    [Title] [varchar](200) NULL,
    [Priority] [int] NULL,
    [State] [varchar](50) NULL,
    [Reason] [varchar](50) NULL,
    [StoryPoints] [float] NULL,
    [OriginalEstimate] [float] NULL,
    [CompletedWork] [float] NULL,
    [AssignedTo] [varchar](50) NOT NULL,
    [CreatedBy] [varchar](50) NOT NULL,
    [CreatedDate] [datetime2] NOT NULL,
    [ChangedBy] [varchar](50) NULL,
    [ChangedDate] [datetime2] NULL,
    [ClosedBy] [varchar](50) NULL,
    [ClosedDate] [datetime2] NULL,
    [ResolvedBy] [varchar](50) NULL,
    [ResolvedDate] [datetime2] NULL,
    [StateChangeDate] [datetime2] NULL,
    [Timestamp] [datetime2] NOT NULL
)
GO