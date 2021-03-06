USE [DataModels]
GO
/****** Object:  Table [dbo].[ExecutionTables]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ExecutionTables](
	[ExecutionTableId] [int] NOT NULL,
	[TableName] [nvarchar](255) NULL,
	[Type] [nvarchar](max) NULL,
 CONSTRAINT [PK_ExecutionTables] PRIMARY KEY CLUSTERED 
(
	[ExecutionTableId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
