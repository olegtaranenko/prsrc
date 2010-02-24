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
	w.*, 
	c.rowLock, 
	c.Stat, 
	c.Nevip, 
	oe.DateTimeMO, 
	oe.workTimeMO, 
	mo.StatM, 
	mo.StatO
FROM 
	wCO2_plus w
	INNER JOIN OrdersInCeh c  ON w.numOrder = c.numOrder
	LEFT JOIN OrdersMO     mo ON w.numOrder = mo.numOrder
	LEFT JOIN OrdersEquip  oe ON w.numOrder = oe.numOrder
