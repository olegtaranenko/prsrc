if exists (select 1 from sysviews where viewname = 'vw_OrdersMOSummary') then
	drop view vw_OrdersMOSummary
end if;
 

