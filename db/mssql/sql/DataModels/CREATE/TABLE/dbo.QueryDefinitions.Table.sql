USE [DataModels]
GO
/****** Object:  Table [dbo].[QueryDefinitions]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[QueryDefinitions](
	[QueryDefinitionId] [int] NOT NULL,
	[Name] [nvarchar](255) NULL,
	[Type] [nvarchar](max) NULL,
 CONSTRAINT [PK_QueryDefinitions] PRIMARY KEY CLUSTERED 
(
	[QueryDefinitionId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
