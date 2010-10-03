if exists (select 1 from sysviews where viewname = 'itemWareOrde' and vcreator = 'dba') then
	drop view itemWareOrde;
end if;

create view itemWareOrde (numorder, nomnom, quant, cenaEd, prid, prext, quantEd)
as 
select numorder, nomnom, round(fn.quantity * io.quant, 5), null, io.prid, io.prext, fn.quantity
from xpredmetybyizdelia io 
join itemWareFixe fn on io.prid = fn.productid
	union all
select io.numorder, v.nomnom, round(p.quantity * io.quant, 5), null, io.prid, io.prext, p.quantity
from xpredmetybyizdelia io 
join xvariantnomenc v on v.numorder = io.numorder and v.prid = io.prid and v.prext = io.prext
join sproducts p on p.productid = io.prid and v.nomnom = p.nomnom
	union all
select po.numorder, po.nomnom, po.quant, po.cenaEd, null, null, n.perList
from xpredmetybynomenk po
join sGuideNomenk n on n.nomnom = po.nomnom
;

