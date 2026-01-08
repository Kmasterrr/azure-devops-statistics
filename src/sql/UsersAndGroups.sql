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
) 
GO

CREATE TABLE [dbo].[UserGroups]
(
    [Timestamp] [datetime2] NULL,
    [PrincipalName] [varchar](200) NOT NULL,
    [DisplayName] [varchar](200) NOT NULL,
    [GroupName] [varchar](200) NOT NULL,
    [GroupType] [varchar](200) NOT NULL
) 
GO