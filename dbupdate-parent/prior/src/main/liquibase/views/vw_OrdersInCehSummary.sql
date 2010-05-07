if exists (select 1 from sysviews where viewname = 'vw_OrdersInCehSummary') then
	drop view vw_OrdersInCehSummary
end if;
 

CREATE VIEW vw_OrdersInCehSummary (
	  numorder
	, urgent
	, dateTimeMO
	, statM
	, statO
	, werk
) as
select
	  numorder
	, urgent
	, dateTimeMO
	, statM
	, statO
	, gw.werkCode
from
	OrdersInCeh oc
	join GuideWerk gw on oc.werkId = gw.werkId
