if exists (select 1 from sysviews where viewname = 'isumSellShip' and vcreator = 'dba') then
	drop view isumSellShip;
end if;


create view isumSellShip (numorder, nomnom, quant, date1, date2)
as 
select numorder, nomnom, sum(quant) as quant, min(date1), max(date1)
from itemSellShip
group by
	numorder, nomnom
;
