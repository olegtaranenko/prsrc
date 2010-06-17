if exists (select 1 from sysviews where viewname = 'itemSellOrde' and vcreator = 'dba') then
	drop view itemSellOrde;
end if;


create view itemSellOrde (numorder, nomnom, quant, cenaEd, statusid)
as
select r.numdoc, r.nomnom, r.quantity as quant, r.intQuant, o.statusid
from 
sdmcrez r
join bayorders o on o.numorder = r.numdoc
;

