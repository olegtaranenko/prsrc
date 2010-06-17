if exists (select 1 from sysviews where viewname = 'isumSellOrde' and vcreator = 'dba') then
	drop view isumSellOrde;
end if;


create view isumSellOrde (numorder, nomnom, quant, cenaEd, statusid)
as
select i.numorder, i.nomnom, i.quant, i.cenaEd, i.statusid
from itemSellOrde i
;

