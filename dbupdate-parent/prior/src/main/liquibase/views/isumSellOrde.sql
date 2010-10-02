if exists (select 1 from sysviews where viewname = 'isumSellOrde' and vcreator = 'dba') then
	drop view isumSellOrde;
end if;


create view isumSellOrde (numorder, nomnom, quant, cenaEd, statusid)
as
select i.numorder, i.nomnom, sum(isnull(i.quant, 0)), max(isnull(i.cenaEd,0)), i.statusid
from itemSellOrde i
group by i.numorder, i.nomnom, i.statusId
;

