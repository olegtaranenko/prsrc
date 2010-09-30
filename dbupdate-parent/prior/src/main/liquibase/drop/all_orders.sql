if exists (select 1 from sysviews where viewname = 'all_orders' and vcreator = 'dba') then
	drop view all_orders;
end if;

