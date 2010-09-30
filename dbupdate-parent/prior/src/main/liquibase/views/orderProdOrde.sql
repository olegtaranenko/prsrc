if exists (select 1 from sysviews where viewname = 'orderProdOrde' and vcreator = 'dba') then
	drop view orderProdOrde;
end if;

create view orderProdOrde (numorder, nomnom, quant, cenaEd, prid, prext, quantEd, perList, edizm, edizmList, nomName)
as 
select ipo.*, n.perList, n.ed_izmer, n.ed_izmer2, n.nomName
from itemProdOrde ipo
join sGuideNomenk n on n.nomnom = ipo.nomnom
;