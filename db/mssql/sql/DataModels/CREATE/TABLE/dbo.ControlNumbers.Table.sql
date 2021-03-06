USE [DataModels]
GO
/****** Object:  Table [dbo].[ControlNumbers]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ControlNumbers](
	[ControlNumberId] [int] NOT NULL,
	[RPIO] [nvarchar](255) NULL,
	[RegionNumber] [float] NULL,
	[BFY] [nvarchar](255) NULL,
	[CalendarYear] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[FundNumber] [nvarchar](255) NULL,
	[AhCode] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[DivisionNumber] [float] NULL,
	[DateIssued] [datetime] NULL,
	[Purpose] [nvarchar](max) NULL,
 CONSTRAINT [PK_ControlNumbers] PRIMARY KEY CLUSTERED 
(
	[ControlNumberId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
