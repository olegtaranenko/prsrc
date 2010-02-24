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
	w.*, 
	c.rowLock, 
	c.Stat, 
	c.Nevip, 
	oe.DateTimeMO, 
	oe.workTimeMO, 
	mo.StatM, 
	mo.StatO
FROM 
	wYAG_plus w
	INNER JOIN OrdersInCeh c  ON w.numOrder = c.numOrder
	LEFT JOIN OrdersMO     mo ON w.numOrder = mo.numOrder
	LEFT JOIN OrdersEquip  oe ON w.numOrder = oe.numOrder
