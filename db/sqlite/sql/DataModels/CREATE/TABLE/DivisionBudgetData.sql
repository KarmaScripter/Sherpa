CREATE TABLE DivisionBudgetData
(
	DivisionBudgetDataId INTEGER NOT NULL UNIQUE CONSTRAINT PrimaryKeyDivisionBudgetData PRIMARY KEY,
	BFY TEXT(255) NULL,
	AhCode TEXT(255) NULL,
	RcCode TEXT(255) NULL,
	FundCode TEXT(255) NULL
);

