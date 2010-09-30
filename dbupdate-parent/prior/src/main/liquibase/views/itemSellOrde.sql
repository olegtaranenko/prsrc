if exists (select 1 from sysviews where viewname = 'itemSellOrde' and vcreator = 'dba') then
	drop view itemSellOrde;
end if;


create view itemSellOrde (numorder, nomnom, quant, cenaEd, statusid)
as
select r.numorder, r.nomnom, r.quant as quant, pn.cenaEd * n.perlist, o.statusid
from 
bayorders o 
join baynomenk r on o.numorder = r.numorder
join xpredmetybynomenk pn on pn.numorder = o.numorder and pn.nomnom = r.nomnom
join sguidenomenk n on n.nomnom = r.nomnom
;

