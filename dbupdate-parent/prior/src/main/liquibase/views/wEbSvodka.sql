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
	GuideFirms.xLogin, 
	Orders.numOrder, 
	Orders.StatusId, 
	Orders.outDateTime, 
	GuideProblem.Problem, 
	Orders.Logo, 
	Orders.Product, 
	Orders.ordered, 
	Orders.paid, 
	Orders.shipped, 
	GuideFirms.Name, 
	GuideManag.Manag, 
	Orders.DateRS
FROM 
	GuideStatus 
	INNER JOIN (GuideProblem 
		INNER JOIN (GuideFirms 
			INNER JOIN (GuideManag 
				INNER JOIN Orders ON GuideManag.ManagId = Orders.ManagId) 
			ON GuideFirms.FirmId = Orders.FirmId) 
		ON GuideProblem.ProblemId = Orders.ProblemId) 
	ON GuideStatus.StatusId = Orders.StatusId
WHERE Orders.StatusId not in (6, 7)