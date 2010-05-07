if exists (select 1 from sysviews where viewname = 'wCO2') then
	drop view wCO2
end if;
 

