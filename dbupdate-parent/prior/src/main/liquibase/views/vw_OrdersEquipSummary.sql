if exists (select 1 from sysviews where viewname = 'vw_OrdersEquipSummary') then
	drop view vw_OrdersEquipSummary
end if;
 

CREATE VIEW vw_OrdersEquipSummary (
	  numorder
	, worktime
	, outDateTime
) as
select
	  numorder
	, sum(isnull(worktime, 0.0))
	, max(outDateTime)
from
	OrdersEquip
group by 
	numorder
