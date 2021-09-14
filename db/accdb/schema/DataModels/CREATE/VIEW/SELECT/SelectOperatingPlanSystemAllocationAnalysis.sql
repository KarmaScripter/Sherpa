SELECT DISTINCTROW StatusOfFunds.BudgetLevel AS BudgetLevel, StatusOfFunds.BFY AS BFY, StatusOfFunds.AhCode AS AhCode, StatusOfFunds.FundCode AS FundCode, UCase(StatusOfFunds.FundName) AS FundName, StatusOfFunds.OrgCode AS OrgCode, StatusOfFunds.AccountCode AS AccountCode, UCase(StatusOfFunds.ProgramProjectName) AS ProgramProjectName, 
StatusOfFunds.ProgramAreaCode AS ProgramAreaCode, UCase(StatusOfFunds.ProgramAreaName) AS ProgramAreaName, 
StatusOfFunds.BocCode AS BocCode, UCase(StatusOfFunds.BocName) As BocName, StatusOfFunds.Amount AS System, 
OperatingPlans.Amount AS OperatingPlan, 
IIf(StatusOfFunds.Amount - OperatingPlans.Amount < 0, "INCREASE", "DECREASE") AS NET, 
Round(Abs(StatusOfFunds.Amount - OperatingPlans.Amount), 2) AS Delta
FROM StatusOfFunds 
INNER JOIN OperatingPlans 
ON (StatusOfFunds.RcCode = OperatingPlans.RcCode) 
AND (StatusOfFunds.BocCode = OperatingPlans.BocCode) 
AND (StatusOfFunds.AccountCode = OperatingPlans.AccountCode) 
AND (StatusOfFunds.OrgCode = OperatingPlans.OrgCode) 
AND (StatusOfFunds.FundCode = OperatingPlans.FundCode) 
AND (StatusOfFunds.AhCode = OperatingPlans.AhCode) 
AND (StatusOfFunds.BFY = OperatingPlans.BFY)
WHERE StatusOfFunds.Amount <> OperatingPlans.Amount;