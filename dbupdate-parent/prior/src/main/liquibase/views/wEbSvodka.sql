if exists (select 1 from sysviews where viewname = 'wEbSvodka') then
	drop view wEbSvodka
end if;
 

CREATE VIEW wEbSvodka (
	xLogin, 
	numOrder, 
	StatusId, 
	outDateTime, 
	Problem, 
	Logo, 
	Product, 
	ordered, 
	paid, 
	shipped, 
	Name, 
	Manag, 
	DateRS
) as
SELECT 
	f.xLogin, 
	o.numOrder, 
	o.StatusId, 
	oe.outDateTime, 
	p.Problem, 
	o.Logo, 
	o.Product, 
	o.ordered, 
	o.paid, 
	o.shipped, 
	f.Name, 
	m.Manag, 
	o.DateRS
FROM 
	Orders o 
	JOIN GuideStatus  s ON s.StatusId = o.StatusId
	JOIN GuideProblem p ON p.ProblemId = o.ProblemId
	JOIN GuideFirms   f ON f.FirmId = o.FirmId
	JOIN GuideManag   m ON m.ManagId = o.ManagId
	LEFT JOIN vw_OrdersEquipSummary oe on oe.numorder = o.numorder
WHERE o.StatusId not in (6, 7)