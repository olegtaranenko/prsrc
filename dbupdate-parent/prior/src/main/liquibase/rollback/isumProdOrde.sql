if exists (select 1 from sysviews where viewname = 'isumProdOrde' and vcreator = 'dba') then
	drop view isumProdOrde;
end if;

create view isumProdOrde(numorder, nomnom, quant, statusid)
as 
select i.numorder, i.nomnom, sum(i.quant), o.statusid
from itemProdOrde i
join orders o on o.numorder = i.numorder
group by i.numorder, i.nomnom, o.statusid
;
