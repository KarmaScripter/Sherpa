USE [ReferenceModels]
GO
/****** Object:  Table [dbo].[Appropriations]    Script Date: 7/17/2021 9:04:59 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Appropriations](
	[AppropriationId] [int] NOT NULL,
	[BFY] [nvarchar](255) NOT NULL,
	[Title] [nvarchar](255) NULL,
	[PublicLaw] [nvarchar](255) NULL,
	[EnactedDate] [datetime] NULL,
 CONSTRAINT [PK_Appropriations] PRIMARY KEY CLUSTERED 
(
	[AppropriationId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
