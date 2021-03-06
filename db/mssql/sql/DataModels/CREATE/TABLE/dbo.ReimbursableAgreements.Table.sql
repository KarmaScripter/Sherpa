USE [DataModels]
GO
/****** Object:  Table [dbo].[ReimbursableAgreements]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ReimbursableAgreements](
	[ReimbursableAgreementId] [int] NOT NULL,
	[RPIO] [nvarchar](255) NULL,
	[BFY] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[AgreementNumber] [nvarchar](255) NULL,
	[StartDate] [datetime] NULL,
	[EndDate] [datetime] NULL,
	[RcCode] [nvarchar](255) NULL,
	[OrgCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[SiteProjectCode] [nvarchar](255) NULL,
	[AccountCode] [nvarchar](255) NULL,
	[VendorCode] [nvarchar](255) NULL,
	[VendorName] [nvarchar](255) NULL,
	[Amount] [money] NULL,
	[OpenCommitments] [money] NULL,
	[Obligations] [money] NULL,
	[ULO] [money] NULL,
	[Available] [money] NULL,
 CONSTRAINT [PK_ReimbursableAgreements] PRIMARY KEY CLUSTERED 
(
	[ReimbursableAgreementId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
