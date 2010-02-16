if exists (select 1 from sysviews where viewname = 'wCO2') then
	drop view wCO2
end if;
 

CREATE VIEW wCO2 (
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
	[wCO2_plus].*, 
	OrdersInCeh.rowLock, 
	OrdersInCeh.Stat, 
	OrdersInCeh.Nevip, 
	OrdersMO.DateTimeMO, 
	OrdersMO.workTimeMO, 
	OrdersMO.StatM, 
	OrdersMO.StatO
FROM (
	[wCO2_plus] 
	INNER JOIN OrdersInCeh ON [wCO2_plus].numOrder = OrdersInCeh.numOrder) 
	LEFT JOIN OrdersMO ON [wCO2_plus].numOrder = OrdersMO.numOrder;