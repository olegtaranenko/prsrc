if exists (select 1 from sysviews where viewname = 'vw_Reestr') then
	drop view vw_Reestr
end if;
 

CREATE VIEW vw_Reestr (
	numOrder
	,werkId
	,Manag
	,StatusId
	,StatusEquipId
	,ProblemId
	,DateRS
	,outDateTime
	,workTime
	,Name
	,Logo
	,Product
	,Stat
	,StatEquip
	,Nevip
	,DateTimeMO
	,workTimeMO
	,StatM
	,StatO
	,equip
	,equipId
) as
SELECT 
	o.numOrder
	,o.werkId
	,m.Manag
	,o.StatusId
	,oe.StatusEquipId
	,o.ProblemId
	,o.DateRS
	,dateadd(hour, isnull(o.outtime, 0), oe.outDateTime)
	,oe.workTime
	,f.Name
	,o.Logo
	,o.Product
	,oc.Stat
	,s.Status
	,isnull(oe.Nevip, 1)
	,oc.DateTimeMO
	,oe.workTimeMO
	,oc.StatM
	,oc.StatO
	,e.equipName
	,oe.equipId
FROM Orders o
	JOIN OrdersEquip       oe ON o.numOrder = oe.numOrder
	JOIN GuideEquip         e ON e.equipId  = oe.equipId
	JOIN GuideFirms         f ON f.FirmId = o.FirmId
	JOIN GuideManag         m ON m.ManagId = o.ManagId
	JOIN OrdersInCeh       oc ON o.numOrder = oc.numOrder
	LEFT JOIN GuideStatus   s ON s.statusId = oe.statusEquipId
where outDatetime is not null
