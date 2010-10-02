if exists (select 1 from sysviews where viewname = 'itemSellOrde' and vcreator = 'dba') then
	drop view itemSellOrde;
end if;


create view itemSellOrde (numorder, nomnom, quant, cenaEd, prid, prext, quantEd, statusid)
as
select b.*, o.statusid
from itemBranOrde b
join orders o on o.numorder = b.numorder
where o.werkId = 1;
