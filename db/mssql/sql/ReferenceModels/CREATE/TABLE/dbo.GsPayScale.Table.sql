USE [ReferenceModels]
GO
/****** Object:  Table [dbo].[GsPayScale]    Script Date: 7/17/2021 9:04:59 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GsPayScale](
	[GsPayScaleId] [int] NOT NULL,
	[LOCNAME] [nvarchar](255) NULL,
	[GRADE] [float] NULL,
	[ANNUAL1] [float] NULL,
	[HOURLY1] [nvarchar](255) NULL,
	[OVERTIME1] [nvarchar](255) NULL,
	[ANNUAL2] [float] NULL,
	[HOURLY2] [nvarchar](255) NULL,
	[OVERTIME2] [nvarchar](255) NULL,
	[ANNUAL3] [float] NULL,
	[HOURLY3] [nvarchar](255) NULL,
	[OVERTIME3] [nvarchar](255) NULL,
	[ANNUAL4] [float] NULL,
	[HOURLY4] [nvarchar](255) NULL,
	[OVERTIME4] [nvarchar](255) NULL,
	[ANNUAL5] [float] NULL,
	[HOURLY5] [nvarchar](255) NULL,
	[OVERTIME5] [nvarchar](255) NULL,
	[ANNUAL6] [float] NULL,
	[HOURLY6] [nvarchar](255) NULL,
	[OVERTIME6] [nvarchar](255) NULL,
	[ANNUAL7] [float] NULL,
	[HOURLY7] [nvarchar](255) NULL,
	[OVERTIME7] [nvarchar](255) NULL,
	[ANNUAL8] [float] NULL,
	[HOURLY8] [nvarchar](255) NULL,
	[OVERTIME8] [nvarchar](255) NULL,
	[ANNUAL9] [float] NULL,
	[HOURLY9] [nvarchar](255) NULL,
	[OVERTIME9] [nvarchar](255) NULL,
	[ANNUAL10] [float] NULL,
	[HOURLY10] [nvarchar](255) NULL,
	[OVERTIME10] [nvarchar](255) NULL,
 CONSTRAINT [PK_GsPayScale] PRIMARY KEY CLUSTERED 
(
	[GsPayScaleId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
