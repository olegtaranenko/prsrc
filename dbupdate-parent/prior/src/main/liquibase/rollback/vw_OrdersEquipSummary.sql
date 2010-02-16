if exists (select 1 from sysviews where viewname = 'vw_OrdersEquipSummary') then
	drop view vw_OrdersEquipSummary
end if;
 

