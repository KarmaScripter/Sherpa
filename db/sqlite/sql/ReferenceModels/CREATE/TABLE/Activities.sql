CREATE TABLE Activities
(
	ActivityId INTEGER NOT NULL UNIQUE CONSTRAINT PrimaryKeyActivity PRIMARY KEY AUTOINCREMENT,
	Code TEXT(255) NOT NULL,
	Name TEXT(255) NULL,
	Title TEXT(255) NULL
);

