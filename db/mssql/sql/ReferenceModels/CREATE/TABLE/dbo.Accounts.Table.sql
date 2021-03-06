USE [ReferenceModels]
GO
/****** Object:  Table [dbo].[Accounts]    Script Date: 7/17/2021 9:04:58 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Accounts](
	[AccountCodeId] [int] NOT NULL,
	[Code] [nvarchar](255) NULL,
	[ProgramProjectCode] [nvarchar](255) NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[ProgramAreaCode] [nvarchar](255) NULL,
	[ProgramAreaName] [nvarchar](255) NULL,
	[GoalCode] [nvarchar](max) NULL,
	[ObjectiveCode] [nvarchar](max) NULL,
	[NpmCode] [nvarchar](max) NULL,
 CONSTRAINT [PK_Accounts] PRIMARY KEY CLUSTERED 
(
	[AccountCodeId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
