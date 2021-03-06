USE [DataModels]
GO
/****** Object:  Table [dbo].[ReportingCodes]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ReportingCodes](
	[ReportingCodeId] [int] NOT NULL,
	[ReportingCode] [nvarchar](255) NULL,
	[ReportingCodeName] [nvarchar](255) NULL,
 CONSTRAINT [PK_ReportingCodes] PRIMARY KEY CLUSTERED 
(
	[ReportingCodeId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
