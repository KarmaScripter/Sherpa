IF NOT EXISTS ( SELECT * 
				FROM INFORMATION_SCHEMA.TABLES 
				WHERE TABLE_NAME = N'FundSymbols' )
BEGIN
CREATE TABLE [dbo].[FundSymbols]
(
	[TreasurySymbolsId] INT NOT NULL,
	[TreasuryAccount] VARCHAR(80) NULL DEFAULT ('NS'),
	[OmbAccount] VARCHAR(80) NULL DEFAULT ('NS')
);
END
