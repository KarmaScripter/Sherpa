IF NOT EXISTS ( SELECT * 
				FROM INFORMATION_SCHEMA.TABLES 
				WHERE TABLE_NAME = N'FullTimeEquivalents' )
BEGIN
CREATE TABLE [dbo].[FullTimeEquivalents]
(
	[FullTimeEquivialentsId] INT NOT NULL,
	[OperatingPlanId] INT NOT NULL,
	[RpioCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[RpioName] VARCHAR(80) NULL DEFAULT ('NS'),
	[BFY] VARCHAR(80) NULL DEFAULT ('NS'),
	[EFY] VARCHAR(80) NULL DEFAULT ('NS'),
	[AhCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[FundCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[OrgCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[AccountCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[RcCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[BocCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[BocName] VARCHAR(80) NULL DEFAULT ('NS'),
	[Amount] FLOAT NOT NULL DEFAULT 0,
	[ITProjectCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProjectCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProjectName] VARCHAR(80) NULL DEFAULT ('NS'),
	[NpmCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProjectTypeName] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProjectTypeCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProgramProjectCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProgramAreaCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[NpmName] VARCHAR(80) NULL DEFAULT ('NS'),
	[AhName] VARCHAR(80) NULL DEFAULT ('NS'),
	[FundName] VARCHAR(80) NULL DEFAULT ('NS'),
	[OrgName] VARCHAR(80) NULL DEFAULT ('NS'),
	[RcName] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProgramProjectName] VARCHAR(80) NULL DEFAULT ('NS'),
	[ActivityCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ActivityName] VARCHAR(80) NULL DEFAULT ('NS'),
	[LocalCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[LocalCodeName] VARCHAR(80) NULL DEFAULT ('NS'),
	[ProgramAreaName] VARCHAR(80) NULL DEFAULT ('NS'),
	[CostAreaCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[CostAreaName] VARCHAR(80) NULL DEFAULT ('NS'),
	[GoalCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[GoalName] VARCHAR(80) NULL DEFAULT ('NS'),
	[ObjectiveCode] VARCHAR(80) NULL DEFAULT ('NS'),
	[ObjectiveName] VARCHAR(80) NOT NULL
);
END



