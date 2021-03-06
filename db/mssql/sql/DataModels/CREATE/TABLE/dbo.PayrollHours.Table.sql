USE [DataModels]
GO
/****** Object:  Table [dbo].[PayrollHours]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PayrollHours](
	[PayrollHoursId] [int] NOT NULL,
	[RpioCode] [nvarchar](255) NULL,
	[PayPeriod] [nvarchar](255) NULL,
	[StartDate] [datetime] NULL,
	[EndDate] [datetime] NULL,
	[EpaNumber] [nvarchar](255) NULL,
	[FirstName] [nvarchar](255) NULL,
	[LastName] [nvarchar](255) NULL,
	[HrOrgCode] [nvarchar](255) NULL,
	[HrOrgName] [nvarchar](255) NULL,
	[ReportingCode] [nvarchar](255) NULL,
	[ReportingCodeName] [nvarchar](255) NULL,
	[WorkCode] [nvarchar](255) NULL,
	[Hours] [float] NULL,
 CONSTRAINT [PK_PayrollHours] PRIMARY KEY CLUSTERED 
(
	[PayrollHoursId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
