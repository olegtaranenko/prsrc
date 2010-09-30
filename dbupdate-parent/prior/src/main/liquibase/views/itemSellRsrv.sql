if exists (select 1 from sysviews where viewname = 'itemSellRsrv' and vcreator = 'dba') then
	drop view itemSellRsrv;
end if;

create view itemSellRsrv (numorder, nomnom, quant, quant_rele, date1)
as
select 
r.numorder, r.nomnom, r.quant, sum(isnull(d.quant, 0)), o.indate
from baynomenk r
left join sdmc d on d.numdoc = r.numorder and d.nomnom = r.nomnom
join bayorders o on o.numorder = r.numorder 
group by r.numorder, r.nomnom, r.quant, o.indate
;
