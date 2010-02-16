if exists (select 1 from sysviews where viewname = 'wSUB') then
	drop view wSUB
end if;
 

CREATE VIEW wSUB (
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
	[wSUB_plus].*, 
	OrdersInCeh.rowLock, 
	OrdersInCeh.Stat, 
	OrdersInCeh.Nevip, 
	OrdersMO.DateTimeMO, 
	OrdersMO.workTimeMO, 
	OrdersMO.StatM, 
	OrdersMO.StatO
FROM (
	[wSUB_plus] 
	INNER JOIN OrdersInCeh ON [wSUB_plus].numOrder = OrdersInCeh.numOrder) 
	LEFT JOIN OrdersMO ON [wSUB_plus].numOrder = OrdersMO.numOrder;