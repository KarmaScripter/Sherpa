USE [DataModels]
GO
/****** Object:  Table [dbo].[UpperReconcilliation]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[UpperReconcilliation](
	[ReconcilliationId] [int] NOT NULL,
	[ExtId] [int] NULL,
	[PrcId] [int] NULL,
	[BFY] [nvarchar](255) NULL,
	[BudgetLevel] [nvarchar](255) NULL,
	[AhCode] [nvarchar](255) NULL,
	[FundName] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[OrgCode] [nvarchar](255) NULL,
	[AccountCode] [nvarchar](255) NULL,
	[BocCode] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[BocName] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[ProgramProjectCode] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[System] [money] NULL,
	[Budget] [money] NULL,
	[Delta] [money] NULL,
	[NET] [nvarchar](255) NULL
) ON [PRIMARY]
GO
