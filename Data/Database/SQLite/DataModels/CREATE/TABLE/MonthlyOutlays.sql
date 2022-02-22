CREATE TABLE IF NOT EXISTS "MonthlyOutlays" 
(
	"MonthlyOutlaysId"	INTEGER NOT NULL UNIQUE,
	"FiscalYear"	TEXT(255) DEFAULT "NS",
	"LineNumber"	TEXT(255) DEFAULT "NS",
	"LineTitle"	TEXT(255) DEFAULT "NS",
	"TaxationCode"	TEXT(255) DEFAULT "NS",
	"TreasuryAgency"	TEXT(255) DEFAULT "NS",
	"TreasuryAccount"	TEXT(255) DEFAULT "NS",
	"SubAccount"	TEXT(255) DEFAULT "NS",
	"BFY"	TEXT(255) DEFAULT "NS",
	"EFY"	TEXT(255) DEFAULT "NS",
	"OmbAgency"	TEXT(255) DEFAULT "NS",
	"OmbBureau"	TEXT(255) DEFAULT "NS",
	"OmbAccount"	TEXT(255) DEFAULT "NS",
	"AgencySequence"	TEXT(255) DEFAULT "NS",
	"BureauSequence"	TEXT(255) DEFAULT "NS",
	"AccountSequence"	TEXT(255) DEFAULT "NS",
	"AgencyTitle"	TEXT(255) DEFAULT "NS",
	"BureauTitle"	TEXT(255) DEFAULT "NS",
	"OmbAccountTitle"	TEXT(255) DEFAULT "NS",
	"TreasuryAccountTitle"	TEXT(255) DEFAULT "NS",
	"October"	REAL DEFAULT 0.0,
	"November"	REAL DEFAULT 0.0,
	"December"	REAL DEFAULT 0.0,
	"January"	REAL DEFAULT 0.0,
	"Feburary"	REAL DEFAULT 0.0,
	"March"	REAL DEFAULT 0.0,
	"April"	REAL DEFAULT 0.0,
	"May"	REAL DEFAULT 0.0,
	"June"	REAL DEFAULT 0.0,
	"July"	REAL DEFAULT 0.0,
	"August"	REAL DEFAULT 0.0,
	"September"	REAL DEFAULT 0.0,
	PRIMARY KEY("MonthlyOutlaysId" AUTOINCREMENT)
);