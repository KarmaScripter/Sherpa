USE [DataModels]
GO
/****** Object:  Table [dbo].[PayrollObligations]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PayrollObligations](
	[PayrollObligationId] [int] NOT NULL,
	[RPIO] [nvarchar](255) NULL,
	[AhCode] [nvarchar](255) NULL,
	[BFY] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[FundName] [nvarchar](255) NULL,
	[OrgCode] [nvarchar](255) NULL,
	[AccountCode] [nvarchar](255) NULL,
	[ProgramProjectCode] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[FocCode] [nvarchar](255) NULL,
	[FocName] [nvarchar](255) NULL,
	[WorkCode] [nvarchar](255) NULL,
	[WorkCodeName] [nvarchar](255) NULL,
	[HrOrgCode] [nvarchar](255) NULL,
	[PayPeriod] [nvarchar](255) NULL,
	[Amount] [float] NULL,
	[Hours] [float] NULL,
	[CumulativeBenefits] [float] NULL,
	[AnnualBase] [float] NULL,
	[AnnualHours] [float] NULL,
	[AnnualOvertimePaid] [float] NULL,
	[AnnualOvertimeHours] [float] NULL,
	[AnnualOtherPaid] [float] NULL,
	[AnnualOtherHours] [float] NULL,
 CONSTRAINT [PK_PayrollObligations] PRIMARY KEY CLUSTERED 
(
	[PayrollObligationId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
