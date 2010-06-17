if exists (select 1 from sysviews where viewname = 'isumBranRequ' and vcreator = 'dba') then
	drop view isumBranRequ;
end if;

create view isumBranRequ
as 
select *
from itemBranRequ r
;
