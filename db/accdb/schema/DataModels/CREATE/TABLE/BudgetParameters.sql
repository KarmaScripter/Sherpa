CREATE TABLE BudgetParameters
(
	BudgetParameterId INTEGER NOT NULL UNIQUE CONSTRAINT PrimaryKeyBudgetParameters PRIMARY KEY,
	BFY TEXT(255) NULL,
	AhCode TEXT(255) NULL,
	RcCode TEXT(255) NULL,
	FundCode TEXT(255) NULL
);
