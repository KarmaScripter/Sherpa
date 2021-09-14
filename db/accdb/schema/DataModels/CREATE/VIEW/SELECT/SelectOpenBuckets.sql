SELECT DISTINCTROW Allocations.BudgetLevel, Allocations.BFY, Allocations.AhCode, Allocations.FundCode, Allocations.OrgCode, Allocations.AccountCode, Allocations.BocCode, Allocations.RcCode, 
CCur(Allocations.Amount) AS Initial, 
CCur(OperatingPlans.Amount) AS Change, 
IIf(Allocations.Amount-OperatingPlans.Amount<0,"INCREASE","DECREASE") AS NET, 
CCUR(Round(Abs(Allocations.Amount-OperatingPlans.Amount),2)) AS Delta
FROM Allocations 
INNER JOIN OperatingPlans 
ON (Allocations.BFY = OperatingPlans.BFY) 
AND (Allocations.BudgetLevel = OperatingPlans.BudgetLevel) 
AND (Allocations.AhCode = OperatingPlans.AhCode) 
AND (Allocations.FundCode = OperatingPlans.FundCode) 
AND (Allocations.OrgCode = OperatingPlans.OrgCode) 
AND (Allocations.BocCode = OperatingPlans.BocCode) 
AND (Allocations.AccountCode = OperatingPlans.AccountCode) 
AND (Allocations.RcCode = OperatingPlans.RcCode)
WHERE (((Allocations.Amount) <> OperatingPlans.Amount));

