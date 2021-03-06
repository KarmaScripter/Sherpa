USE [ReferenceModels]
GO
/****** Object:  Table [dbo].[Holidays]    Script Date: 7/17/2021 9:04:59 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Holidays](
	[HolidayId] [int] NOT NULL,
	[ColumbusDay] [datetime] NULL,
	[ThanksgivingDay] [datetime] NULL,
	[ChristmasDay] [datetime] NULL,
	[NewYearsDay] [datetime] NULL,
	[MartinLutherKingDay] [datetime] NULL,
	[PresidentsDay] [datetime] NULL,
	[MemorialDay] [datetime] NULL,
	[VeteransDay] [datetime] NULL,
	[LaborDay] [datetime] NULL,
 CONSTRAINT [PK_Holidays] PRIMARY KEY CLUSTERED 
(
	[HolidayId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
