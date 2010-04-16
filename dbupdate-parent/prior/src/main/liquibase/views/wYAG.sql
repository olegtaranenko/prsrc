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
	o.numOrder, 
	m.Manag, 
	o.StatusId, 
	o.ProblemId, 
	o.DateRS, 
	oe.outDateTime, 
	oe.workTime, 
	f.Name, 
	o.Logo, 
	o.Product,
	oc.rowLock, 
	oc.Stat, 
	oc.Nevip, 
	oc.DateTimeMO, 
	oc.workTimeMO, 
	oc.StatM, 
	oc.StatO
FROM Orders o
	JOIN OrdersEquip      oe ON o.numOrder = oe.numOrder
	JOIN GuideFirms       f  ON f.FirmId = o.FirmId
	JOIN GuideManag       m  ON m.ManagId = o.ManagId
	JOIN OrdersInCeh oc  ON o.numOrder = oc.numOrder  and oe.cehId = oc.cehId
WHERE oe.CehId = 1
