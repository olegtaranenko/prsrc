if exists (select 1 from sysviews where viewname = 'isumBranOrde' and vcreator = 'dba') then
	drop view isumBranOrde;
end if;

create view isumBranOrde (numorder, nomnom, quant, statusid, scope)
as
select numorder, nomnom, quant, statusid, 'p'
from isumProdOrde 
	union all
select numorder, nomnom, quant, statusid, 'b'
from itemSellOrde 
;
