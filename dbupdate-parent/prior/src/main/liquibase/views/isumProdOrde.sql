if exists (select 1 from sysviews where viewname = 'isumProdOrde' and vcreator = 'dba') then
	drop view isumProdOrde;
end if;

create view isumProdOrde(numorder, nomnom, quant, statusid, nomnomWeb)
as 
select i.numorder, i.nomnom, sum(i.quant), o.statusid, n.web
from itemProdOrde i
join sGuideNomenk n on n.nomnom = i.nomnom
join orders o on o.numorder = i.numorder and o.werkId = 2
group by i.numorder, i.nomnom, o.statusid, n.web
