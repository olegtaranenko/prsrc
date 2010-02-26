if exists (select 1 from sysviews where viewname = 'vw_OrdersInCehSummary') then
	drop view vw_OrdersInCehSummary
end if;
 

