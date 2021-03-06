USE [DataModels]
GO
/****** Object:  Table [dbo].[CongressionalTransfers]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CongressionalTransfers](
	[CongressionalTransferId] [int] NOT NULL,
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
	[AccountCode] [nvarchar](255) NULL,
	[ProgramProjectCode] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[ProgramAreaCode] [nvarchar](255) NULL,
	[ProgramAreaName] [nvarchar](255) NULL,
	[OrgCode] [nvarchar](255) NULL,
	[OrganizationName] [nvarchar](255) NULL,
	[BocCode] [nvarchar](255) NULL,
	[BocName] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[Amount] [float] NULL,
	[SPIO] [nvarchar](255) NULL,
	[Purpose] [nvarchar](max) NULL,
	[ExtendedPurpose] [nvarchar](max) NULL,
	[FromTo] [nvarchar](255) NULL,
	[DocType] [nvarchar](255) NULL,
	[DocPrefix] [nvarchar](255) NULL,
	[NpmCode] [nvarchar](255) NULL,
	[ResourceType] [nvarchar](255) NULL,
 CONSTRAINT [PK_CongressionalTransfers] PRIMARY KEY CLUSTERED 
(
	[CongressionalTransferId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
