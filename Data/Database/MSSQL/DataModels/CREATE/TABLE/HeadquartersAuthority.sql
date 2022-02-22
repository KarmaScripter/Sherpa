IF NOT EXISTS ( SELECT * 
				FROM INFORMATION_SCHEMA.TABLES 
				WHERE TABLE_NAME = N'HeadquartersAuthority' )
BEGIN
CREATE TABLE [dbo].[HeadquartersAuthority]
(
	[HeadquartersAuthorityId] INT NOT NULL,
	[AllocationsId] INT NOT NULL,
	[StatusOfFundsId] INT NOT NULL,
	[BFY] VARCHAR(80) NULL DEFAULT ('NS'),
	[EFY] VARCHAR(80) NULL DEFAULT ('NS'),
	[RpioCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[RpioName] VARCHAR(80) NULL DEFAULT ('NS'),
	[BudgetLevel] VARCHAR(80) NULL DEFAULT ('NS'),
	[AhCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[AhName] VARCHAR(80) NULL DEFAULT ('NS'),
	[FundCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[FundName] VARCHAR(80) NULL DEFAULT ('NS'),
	[OrgCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[OrgName] VARCHAR(80) NULL DEFAULT ('NS'),
	[AccountCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[RcCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[RcName] VARCHAR(80) NULL DEFAULT ('NS'),
	[BocCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[BocName] VARCHAR(80) NULL DEFAULT ('NS'),
	[Amount] [real] NOT NULL,
	[ProgramProjectCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProgramProjectName] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProgramAreaCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProgramAreaName] VARCHAR(80) NULL DEFAULT ('NS'),
	[NpmCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[NpmName] VARCHAR(80) NULL DEFAULT ('NS')
);
END


