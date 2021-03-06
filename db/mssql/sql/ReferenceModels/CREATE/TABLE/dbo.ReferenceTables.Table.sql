USE [ReferenceModels]
GO
/****** Object:  Table [dbo].[ReferenceTables]    Script Date: 7/17/2021 9:04:59 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ReferenceTables](
	[ReferenceTableId] [int] NOT NULL,
	[TableName] [nvarchar](255) NULL,
	[Type] [nvarchar](max) NULL,
 CONSTRAINT [PK_ReferenceTables] PRIMARY KEY CLUSTERED 
(
	[ReferenceTableId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
