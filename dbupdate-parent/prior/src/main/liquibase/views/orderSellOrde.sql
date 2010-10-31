if exists (select 1 from sysviews where viewname = 'orderSellOrde' and vcreator = 'dba') then
	drop view orderSellOrde;
end if;

create view orderSellOrde (numorder, cena, statusid)
as
select i.numorder, sum(i.quant * i.cenaEd ), i.statusid
from isumSellOrde i
join sguidenomenk n on i.nomnom = n.nomnom
group by i.numorder, i.statusid
;
