if exists (select 1 from sysviews where viewname = 'vw_OrdersEquipSummary') then
	drop view vw_OrdersEquipSummary
end if;
 

CREATE VIEW vw_OrdersEquipSummary (
	  numorder
	, worktime
	, worktimeMO
	, outDateTime
	, minOutDateTime
	, maxStatusId
	, minStatusId
	, lastModifiedEquip
	, statO
) as
select
	  numorder
	, sum(isnull(worktime, 0.0))
	, sum(worktimeMO)
	, max(outDateTime)
	, min(outDateTime)
	, max(isnull(statusEquipId, 0))
	, min(isnull(statusEquipId, 0))
	, max(isnull(lastModified, '20000101'))
	, max(statO)
from
	OrdersEquip
group by 
	numorder
