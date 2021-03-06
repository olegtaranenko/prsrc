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
--	,StatEquip
	,Nevip
	,DateTimeMO
	,workTimeMO
	,StatM
	,StatO
	,equip
	,equipId
	,remark
) as
SELECT 
	o.numOrder
	,o.werkId
	,m.Manag
	,oe.StatusEquipId
	,oe.StatusEquipId
	,o.ProblemId
	,o.DateRS
--	,oe.outDateTime
	,dateadd(hour, isnull(o.outtime, 0), convert(datetime, convert(varchar(10), oe.outDateTime,102)))
	,oe.workTime
	,f.Name
	,o.Logo
	,o.Product
	,oe.Stat
--	,s.Status
	,isnull(oe.Nevip, 1)
	,oc.DateTimeMO
	,oe.workTimeMO
	,oc.StatM
	,oe.StatO
	,e.equipName
	,oe.equipId
	,o.remark
FROM Orders o
	JOIN OrdersEquip       oe ON o.numOrder = oe.numOrder
	JOIN GuideEquip         e ON e.equipId  = oe.equipId
	JOIN FirmGuide          f ON f.FirmId = o.FirmId
	JOIN GuideManag         m ON m.ManagId = o.ManagId
	JOIN OrdersInCeh       oc ON o.numOrder = oc.numOrder
	LEFT JOIN GuideStatus   s ON s.statusId = oe.statusEquipId and s.werkId = o.werkid
where oe.outDatetime is not null;