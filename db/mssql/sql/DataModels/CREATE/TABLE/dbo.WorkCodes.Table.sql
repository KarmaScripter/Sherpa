USE [DataModels]
GO
/****** Object:  Table [dbo].[WorkCodes]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[WorkCodes](
	[WorkCodeId] [int] NOT NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[WorkCode] [nvarchar](255) NULL,
	[WorkCodeName] [nvarchar](255) NULL
) ON [PRIMARY]
GO
