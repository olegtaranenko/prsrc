if exists (select 1 from sysviews where viewname = 'wYAG') then
	drop view wYAG
end if;
 

CREATE VIEW wYAG (
	numOrder, 
	Manag, 
	StatusId, 
	ProblemId, 
	DateRS, 
	outDateTime, 
	workTime, 
	Name, 
	Logo, 
	Product,
	rowLock, 
	Stat, 
	Nevip, 
	DateTimeMO, 
	workTimeMO, 
	StatM, 
	StatO
) as
SELECT 
	[wYAG_plus].*, 
	OrdersInCeh.rowLock, 
	OrdersInCeh.Stat, 
	OrdersInCeh.Nevip, 
	OrdersMO.DateTimeMO, 
	OrdersMO.workTimeMO, 
	OrdersMO.StatM, 
	OrdersMO.StatO
FROM (
	[wYAG_plus] 
	INNER JOIN OrdersInCeh ON [wYAG_plus].numOrder = OrdersInCeh.numOrder) 
	LEFT JOIN OrdersMO ON [wYAG_plus].numOrder = OrdersMO.numOrder;