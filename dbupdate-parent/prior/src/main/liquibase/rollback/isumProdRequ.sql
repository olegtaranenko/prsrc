if exists (select 1 from sysviews where viewname = 'isumProdRequ' and vcreator = 'dba') then
	drop view isumProdRequ;
end if;

create view isumProdRequ
as 
select *
from itemProdRequ r
;
