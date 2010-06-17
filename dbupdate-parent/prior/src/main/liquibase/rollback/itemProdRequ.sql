if exists (select 1 from sysviews where viewname = 'itemProdRequ' and vcreator = 'dba') then
	drop view itemProdRequ;
end if;

create view itemProdRequ (numorder, nomnom, quant, statusid)
as 
select r.numdoc, r.nomnom, r.curquant, o.statusid
from sdmcrez r
join orders o on r.numdoc = o.numorder
where r.curquant > 0
;
