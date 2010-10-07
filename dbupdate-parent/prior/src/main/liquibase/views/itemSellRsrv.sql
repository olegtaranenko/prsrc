if exists (select 1 from sysviews where viewname = 'itemSellRsrv' and vcreator = 'dba') then
	drop view itemSellRsrv;
end if;

create view itemSellRsrv (numorder, nomnom, quant, quant_rele, date1)
as
select 
r.numdoc, r.nomnom, r.quantity, sum(isnull(d.quant, 0)), o.indate
from sdmcrez r
left join sdmc d on d.numdoc = r.numdoc and d.nomnom = r.nomnom
join bayorders o on o.numorder = r.numdoc 
group by r.numdoc, r.nomnom, r.quantity, o.indate
;
