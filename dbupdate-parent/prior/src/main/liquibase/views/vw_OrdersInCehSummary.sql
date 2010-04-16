if exists (select 1 from sysviews where viewname = 'vw_OrdersInCehSummary') then
	drop view vw_OrdersInCehSummary
end if;
 

CREATE VIEW vw_OrdersInCehSummary (
	  numorder
	, urgent
	, worktimeMO
	, dateTimeMO
	, statM
	, statO
) as
select
	  numorder
	, max(urgent)
	, sum(isnull(worktimeMO, 0.0))
	, max(dateTimeMO)
	, max(statM)
	, max(statO)
from
	OrdersInCeh
group by 
	numorder
