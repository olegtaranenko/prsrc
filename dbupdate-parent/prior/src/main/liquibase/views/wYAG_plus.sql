if exists (select 1 from sysviews where viewname = 'wYAG_plus') then
	drop view wYAG_plus
end if;
 

CREATE VIEW wYAG_plus (
	numOrder, 
	Manag, 
	StatusId, 
	ProblemId, 
	DateRS, 
	outDateTime, 
	workTime, 
	Name, 
	Logo, 
	Product
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
	o.Product
FROM Orders o
JOIN OrdersEquip oe ON oe.numorder = o.numorder
JOIN GuideFirms f ON f.FirmId = o.FirmId
JOIN GuideManag m ON m.ManagId = o.ManagId
WHERE oe.CehId = 1