USE [DataModels]
GO
/****** Object:  Table [dbo].[EmployeeData]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EmployeeData](
	[EmployeeDataId] [int] NOT NULL,
	[RpioCode] [nvarchar](255) NULL,
	[RpioName] [nvarchar](255) NULL,
	[ActionDate] [datetime] NULL,
	[HiringAuthority] [nvarchar](255) NULL,
	[SupervisorId] [nvarchar](255) NULL,
	[JobTitle] [nvarchar](255) NULL,
	[HrOrgCode] [nvarchar](255) NULL,
	[HrOrgName] [nvarchar](255) NULL,
	[EmployeeId] [nvarchar](255) NULL,
	[FirstName] [nvarchar](255) NULL,
	[LastName] [nvarchar](255) NULL,
	[RetirementPlan] [nvarchar](255) NULL,
	[ScheduledRetirementDate] [datetime] NULL,
	[HireDate] [datetime] NULL,
	[Grade] [nvarchar](255) NULL,
	[Step] [nvarchar](255) NULL,
	[GradeEntry] [datetime] NULL,
	[LastIncrease] [datetime] NULL,
	[StepEntry] [datetime] NULL,
	[WigiDue] [datetime] NULL,
	[EmployeeStatus] [nvarchar](255) NULL,
	[HoursEarnedYearToDate] [float] NULL,
	[CarryoverHours] [float] NULL,
	[HoursAdjustedYearToDate] [float] NULL,
	[HoursBalance] [float] NULL,
	[ProjectedAnnualHours] [float] NULL,
	[ProjectedNextPeriodHours] [float] NULL,
	[HoursTakenYearToDate] [float] NULL,
 CONSTRAINT [PK_EmployeeData] PRIMARY KEY CLUSTERED 
(
	[EmployeeDataId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
