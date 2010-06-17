if exists (select 1 from sysviews where viewname = 'itemSellRequ' and vcreator = 'dba') then
	drop view itemSellRequ;
end if;

create view itemSellRequ (numorder, nomnom, quant, statusid)
as 
select r.numdoc, r.nomnom, r.curquant * n.perlist, o.statusid
from sdmcrez r
join bayorders o on r.numdoc = o.numorder
join sguidenomenk n on r.nomnom = n.nomnom
where r.curquant > 0
;
