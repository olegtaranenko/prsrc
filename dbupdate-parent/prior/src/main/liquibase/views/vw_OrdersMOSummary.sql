if exists (select 1 from sysviews where viewname = 'vw_OrdersMOSummary') then
	drop view vw_OrdersMOSummary
end if;
 

CREATE VIEW vw_OrdersMOSummary (
	  numorder
	, worktimeMO
	, dateTimeMO
	, statM
	, statO
) as
select
	  numorder
	, sum(isnull(worktimeMO, 0.0))
	, max(dateTimeMO)
	, max(statM)
	, max(statO)
from
	OrdersMO
group by 
	numorder
