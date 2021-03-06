USE [ReferenceModels]
GO
/****** Object:  Table [dbo].[FiscalYears]    Script Date: 7/17/2021 9:04:59 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FiscalYears](
	[FiscalYearId] [int] NOT NULL,
	[BFY] [nvarchar](255) NOT NULL,
	[EFY] [nvarchar](255) NULL,
	[FirstYear] [nvarchar](255) NULL,
	[LastYear] [nvarchar](255) NULL,
	[ExpiringYear] [nvarchar](255) NULL,
	[StartDate] [datetime] NULL,
	[EndDate] [datetime] NULL,
	[Availability] [nvarchar](255) NULL,
	[Columbus] [datetime] NULL,
	[Thanksgiving] [datetime] NULL,
	[Christmas] [datetime] NULL,
	[NewYears] [datetime] NULL,
	[MartinLutherKing] [datetime] NULL,
	[Presidents] [datetime] NULL,
	[Memorial] [datetime] NULL,
	[Veterans] [datetime] NULL,
	[Labor] [datetime] NULL,
	[WorkDays] [float] NULL,
	[WeekDays] [float] NULL,
	[WeekEnds] [float] NULL,
 CONSTRAINT [PK_FiscalYears] PRIMARY KEY CLUSTERED 
(
	[FiscalYearId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
