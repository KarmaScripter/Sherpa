USE [ReferenceModels]
GO
/****** Object:  Table [dbo].[Funds]    Script Date: 7/17/2021 9:04:59 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Funds](
	[FundId] [int] NOT NULL,
	[Code] [nvarchar](255) NULL,
	[Name] [nvarchar](255) NULL,
	[Title] [nvarchar](255) NULL,
	[TreasurySymbol] [nvarchar](255) NULL,
 CONSTRAINT [PK_Funds] PRIMARY KEY CLUSTERED 
(
	[FundId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
