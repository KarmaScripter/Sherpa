USE [DataModels]
GO
/****** Object:  Table [dbo].[ProgrammaticReservePriorYearUtilization]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ProgrammaticReservePriorYearUtilization](
	[ProgrammaticReservePriorYearUtilizationId] [int] NOT NULL,
	[OrgCode] [nvarchar](255) NULL,
	[BFY] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[FundName] [nvarchar](255) NULL,
	[AccountCode] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[BocCode] [nvarchar](255) NULL,
	[BocName] [nvarchar](255) NULL,
	[CurrentYear] [money] NULL,
	[Reduction] [money] NULL,
	[Budget] [money] NULL,
	[PriorYear] [money] NULL,
	[OpenCommitments] [money] NULL,
	[ULO] [money] NULL,
 CONSTRAINT [PK_ProgrammaticReservePriorYearUtilization] PRIMARY KEY CLUSTERED 
(
	[ProgrammaticReservePriorYearUtilizationId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
