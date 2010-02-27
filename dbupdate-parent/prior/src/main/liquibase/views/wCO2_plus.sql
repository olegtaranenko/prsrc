if exists (select 1 from sysviews where viewname = 'wCO2_plus') then
	drop view wCO2_plus
end if;
 

