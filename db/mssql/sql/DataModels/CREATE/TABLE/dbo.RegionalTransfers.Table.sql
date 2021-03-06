USE [DataModels]
GO
/****** Object:  Table [dbo].[RegionalTransfers]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[RegionalTransfers](
	[RegionalTransferId] [int] NOT NULL,
	[ReprogrammingNumber] [nvarchar](255) NULL,
	[ProcessedDate] [datetime] NULL,
	[Line] [nvarchar](255) NULL,
	[Subline] [nvarchar](255) NULL,
	[RPIO] [nvarchar](255) NULL,
	[AhCode] [nvarchar](255) NULL,
	[AhName] [nvarchar](255) NULL,
	[BFY] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[FundName] [nvarchar](255) NULL,
	[OrgCode] [nvarchar](255) NULL,
	[OrganizationName] [nvarchar](255) NULL,
	[AccountCode] [nvarchar](255) NULL,
	[ProgramProjectCode] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[ProgramAreaCode] [nvarchar](255) NULL,
	[ProgramAreaName] [nvarchar](255) NULL,
	[FromTo] [nvarchar](255) NULL,
	[BocCode] [nvarchar](255) NULL,
	[BocName] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[Amount] [float] NULL,
	[DocPrefix] [nvarchar](255) NULL,
	[DocType] [nvarchar](255) NULL,
	[Purpose] [nvarchar](max) NULL,
	[ExtendedPurpose] [nvarchar](max) NULL,
	[SPIO] [nvarchar](255) NULL,
	[NpmCode] [nvarchar](255) NULL,
	[ResourceType] [nvarchar](255) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
