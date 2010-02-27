if exists (select 1 from sysviews where viewname = 'wSUB_plus') then
	drop view wSUB_plus
end if;
 

