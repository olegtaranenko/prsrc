if exists (select 1 from sysviews where viewname = 'itemBranShip' and vcreator = 'dba') then
	drop view itemBranShip;
end if;


create view itemBranShip (numorder, nomnom, quant, date1, scope)
as 
select r.numorder, r.nomnom, r.quant, r.date1, 'p'
from itemProdShip r
	union all
select r.numorder, r.nomnom, r.quant, r.date1, 'b'
from itemSellShip r
;
