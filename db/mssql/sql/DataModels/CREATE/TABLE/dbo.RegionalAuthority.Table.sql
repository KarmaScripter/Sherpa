USE [DataModels]
GO
/****** Object:  Table [dbo].[RegionalAuthority]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[RegionalAuthority](
	[RescissionId] [int] NOT NULL,
	[PrcId] [int] NULL,
	[RPIO] [nvarchar](255) NULL,
	[BudgetLevel] [nvarchar](255) NULL,
	[AhCode] [nvarchar](255) NULL,
	[BFY] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[OrgCode] [nvarchar](255) NULL,
	[AccountCode] [nvarchar](255) NULL,
	[ActivityCode] [nvarchar](255) NULL,
	[BocCode] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[Allocation] [money] NULL,
	[Reduction] [money] NULL,
	[Amount] [money] NULL,
	[FundName] [nvarchar](255) NULL,
	[BocName] [nvarchar](255) NULL,
	[Division] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[ProgramProjectCode] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[NpmCode] [nvarchar](255) NULL,
	[NpmName] [nvarchar](255) NULL,
	[GoalCode] [nvarchar](255) NULL,
	[GoalName] [nvarchar](255) NULL,
	[ObjectiveCode] [nvarchar](255) NULL,
	[ObjectiveName] [nvarchar](255) NULL
) ON [PRIMARY]
GO
