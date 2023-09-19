CREATE TABLE IF NOT EXISTS MarginalOutlays 
(
    MarginalOutlaysId  INTEGER NOT NULL UNIQUE,
    FiscalYear TEXT(80) NULL DEFAULT NS,
    BFY TEXT(80) NULL DEFAULT NS,
    EFY TEXT(80) NULL DEFAULT NS,
    TreasuryAccount TEXT(80) NULL DEFAULT NS,
    SubAccount TEXT(80) NULL DEFAULT NS,
    BPOA TEXT(80) NULL DEFAULT NS,
    EPOA TEXT(80) NULL DEFAULT NS,
    MainAccount TEXT(80) NULL DEFAULT NS,
    FundCode TEXT(80) NULL DEFAULT NS,
    FundName TEXT(80) NULL DEFAULT NS,
    BudgetAccountCode TEXT(80) NULL DEFAULT NS,
    BudgetAccountName TEXT(100) NULL DEFAULT NS,
    TreasuryAccountName TEXT(100) NULL DEFAULT NS,
    October DOUBLE NULL DEFAULT 0.0,
    November DOUBLE NULL DEFAULT 0.0,
    December DOUBLE NULL DEFAULT 0.0,
    January DOUBLE NULL DEFAULT 0.0,
    Feburary DOUBLE NULL DEFAULT 0.0,
    March DOUBLE NULL DEFAULT 0.0,
    April DOUBLE NULL DEFAULT 0.0,
    May DOUBLE NULL DEFAULT 0.0,
    June DOUBLE NULL DEFAULT 0.0,
    July DOUBLE NULL DEFAULT 0.0,
    August DOUBLE NULL DEFAULT 0.0,
    September DOUBLE NULL DEFAULT 0.0,
    PRIMARY KEY(MarginalOutlaysId AUTOINCREMENT)
);