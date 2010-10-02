if exists (select 1 from sysviews where viewname = 'itemProdOrde' and vcreator = 'dba') then
	drop view itemProdOrde;
end if;

create view itemProdOrde (numorder, nomnom, quant, cenaEd, prid, prext, quantEd)
as 
select b.* 
from itemBranOrde b
join orders o on o.numorder = b.numorder
where o.werkId = 2
