USE [DataModels]
GO
/****** Object:  Table [dbo].[DivisionExecution]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DivisionExecution](
	[DivisionExecutionId] [int] NOT NULL,
	[BudgetLevel] [nvarchar](255) NULL,
	[RPIO] [nvarchar](255) NULL,
	[BFY] [nvarchar](255) NULL,
	[AhCode] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[FundName] [nvarchar](255) NULL,
	[OrgCode] [nvarchar](255) NULL,
	[AccountCode] [nvarchar](255) NULL,
	[ActivityCode] [nvarchar](255) NULL,
	[ProgramProjectCode] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[ProgramAreaCode] [nvarchar](255) NULL,
	[ProgramAreaName] [nvarchar](255) NULL,
	[BocCode] [nvarchar](255) NULL,
	[BocName] [nvarchar](255) NULL,
	[SiteProjectCode] [nvarchar](255) NULL,
	[Amount] [float] NULL,
	[OpenCommitments] [float] NULL,
	[CommitmentRatio] [float] NULL,
	[Obligations] [float] NULL,
	[ObligationRatio] [float] NULL,
	[Used] [float] NULL,
	[Utilization] [float] NULL,
	[Available] [money] NULL,
	[Availability] [float] NULL,
 CONSTRAINT [PK_DivisionExecution] PRIMARY KEY CLUSTERED 
(
	[DivisionExecutionId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
