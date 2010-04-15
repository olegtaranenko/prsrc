if exists (select 1 from sysviews where viewname = 'vw_OrdersEquipSummary') then
	drop view vw_OrdersEquipSummary
end if;
 

CREATE VIEW vw_OrdersEquipSummary (
	  numorder
	, worktime
	, outDateTime
	, minOutDateTime
	, maxStatusId
	, minStatusId
	, lastModifiedEquip
) as
select
	  numorder
	, sum(isnull(worktime, 0.0))
	, max(outDateTime)
	, min(outDateTime)
	, max(statusEquipId)
	, min(statusEquipId)
	, max(lastModified)
from
	OrdersEquip
group by 
	numorder
