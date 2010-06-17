if exists (select 1 from sysviews where viewname = 'isumSellRequ' and vcreator = 'dba') then
	drop view isumSellRequ;
end if;

create view isumSellRequ
as 
select *
from itemSellRequ r
;
