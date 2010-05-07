if exists (select 1 from sysviews where viewname = 'vw_Reestr') then
	drop view vw_Reestr
end if;
 

CREATE VIEW vw_Reestr (
	numOrder
	,werkId
	,Manag
	,StatusId
	,ProblemId
	,DateRS
	,outDateTime
	,workTime
	,Name
	,Logo
	,Product
--	Stat, 
--	Nevip, 
--	DateTimeMO, 
--	workTimeMO, 
--	StatM, 
--	StatO
) as
SELECT 
	o.numOrder
	,o.werkId
	,m.Manag
	,o.StatusId
	,o.ProblemId
	,o.DateRS
	,oe.outDateTime
	,oe.workTime
	,f.Name
	,o.Logo
	,o.Product
--	oc.Stat, 
--	oc.Nevip, 
--	oc.DateTimeMO, 
--	oc.workTimeMO, 
--	oc.StatM, 
--	oc.StatO
FROM Orders o
	JOIN vw_OrdersEquipSummary oe ON o.numOrder = oe.numOrder
	JOIN GuideFirms            f  ON f.FirmId = o.FirmId
	JOIN GuideManag            m  ON m.ManagId = o.ManagId
	LEFT JOIN vw_OrdersInCehSummary oc ON o.numOrder = oc.numOrder

