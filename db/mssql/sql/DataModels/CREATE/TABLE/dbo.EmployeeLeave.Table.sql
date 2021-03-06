USE [DataModels]
GO
/****** Object:  Table [dbo].[EmployeeLeave]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EmployeeLeave](
	[EmployeeLeaveId] [int] NOT NULL,
	[RcCode] [nvarchar](255) NULL,
	[LastName] [nvarchar](255) NULL,
	[FirstName] [nvarchar](255) NULL,
	[EpaNumber] [nvarchar](255) NULL,
	[HoursEarnedYearToDate] [float] NULL,
	[CarryoverHours] [float] NULL,
	[HoursAdjustedYearToDate] [float] NULL,
	[HoursBalance] [float] NULL,
	[ProjectedAnnualHours] [float] NULL,
	[ProjectedNextPeriodHours] [float] NULL,
	[HoursTakenYearToDate] [float] NULL
) ON [PRIMARY]
GO
