CREATE TABLE Documents
(
	DocumentId INTEGER NOT NULL UNIQUE CONSTRAINT PK_Documents PRIMARY KEY AUTOINCREMENT,
	Code TEXT(255) NULL,
	Category TEXT(255) NULL,
	Name TEXT(255) NULL,
	System TEXT(255) NULL
);

