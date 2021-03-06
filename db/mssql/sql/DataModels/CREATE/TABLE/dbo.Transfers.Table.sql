USE [DataModels]
GO
/****** Object:  Table [dbo].[Transfers]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Transfers](
	[TransferId] [int] NOT NULL,
	[BudgetLevel] [nvarchar](255) NULL,
	[DocPrefix] [nvarchar](255) NULL,
	[DocType] [nvarchar](255) NULL,
	[BFY] [nvarchar](255) NULL,
	[RpioCode] [nvarchar](255) NULL,
	[RpioName] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[FundName] [nvarchar](255) NULL,
	[ReprogrammingNumber] [nvarchar](255) NULL,
	[ControlNumber] [nvarchar](255) NULL,
	[ProcessedDate] [datetime] NULL,
	[Quarter] [nvarchar](255) NULL,
	[Line] [nvarchar](255) NULL,
	[Subline] [nvarchar](255) NULL,
	[AhCode] [nvarchar](255) NULL,
	[AhName] [nvarchar](255) NULL,
	[OrgCode] [nvarchar](255) NULL,
	[OrganizationName] [nvarchar](255) NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[AccountCode] [nvarchar](255) NULL,
	[ProgramAreaCode] [nvarchar](255) NULL,
	[ProgramAreaName] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[ProgramProjectCode] [nvarchar](255) NULL,
	[FromTo] [nvarchar](255) NULL,
	[BocCode] [nvarchar](255) NULL,
	[BocName] [nvarchar](255) NULL,
	[NpmCode] [nvarchar](255) NULL,
	[Amount] [float] NULL,
	[Purpose] [nvarchar](max) NULL,
	[ExtendedPurpose] [nvarchar](max) NULL,
	[ResourceType] [nvarchar](255) NULL,
 CONSTRAINT [PK_Transfers] PRIMARY KEY CLUSTERED 
(
	[TransferId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
