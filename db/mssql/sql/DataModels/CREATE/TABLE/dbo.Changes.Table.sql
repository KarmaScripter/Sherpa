USE [DataModels]
GO
/****** Object:  Table [dbo].[Changes]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Changes](
	[ID] [int] NOT NULL,
	[TableName] [nvarchar](255) NULL,
	[FieldName] [nvarchar](255) NULL,
	[Action] [nvarchar](255) NULL,
	[OldValue] [nvarchar](255) NULL,
	[NewValue] [nvarchar](255) NULL,
	[TimeStamp] [datetime] NULL,
	[Message] [nvarchar](255) NULL,
 CONSTRAINT [PK_Changes] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
