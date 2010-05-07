if exists (select 1 from sysviews where viewname = 'wSUB') then
	drop view wSUB
end if;
 

