if exists (select 1 from sysviews where viewname = 'vw_OrdersInCehSummary') then
	drop view vw_OrdersInCehSummary
end if;
 

CREATE VIEW vw_OrdersInCehSummary (
	  numorder
	, urgent
) as
select
	  numorder
	, max(urgent)
from
	OrdersInCeh
group by 
	numorder
