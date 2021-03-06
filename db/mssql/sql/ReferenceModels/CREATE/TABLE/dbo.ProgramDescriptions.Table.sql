USE [ReferenceModels]
GO
/****** Object:  Table [dbo].[ProgramDescriptions]    Script Date: 7/17/2021 9:04:59 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ProgramDescriptions](
	[ProgramProjectId] [int] NOT NULL,
	[ProgramProjectCode] [nvarchar](255) NOT NULL,
	[ProgramProjectName] [nvarchar](255) NULL,
	[ProgramProjectTitle] [nvarchar](255) NULL,
	[Laws] [nvarchar](max) NULL,
	[Narrative] [nvarchar](max) NULL,
	[Definition] [nvarchar](max) NULL,
	[ProgramAreaCode] [nvarchar](255) NULL,
	[ProgramAreaName] [nvarchar](255) NULL,
 CONSTRAINT [PK_ProgramDescriptions] PRIMARY KEY CLUSTERED 
(
	[ProgramProjectId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
