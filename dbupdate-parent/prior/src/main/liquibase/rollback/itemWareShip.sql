if exists (select 1 from sysviews where viewname = 'itemWareShip' and vcreator = 'dba') then
	drop view itemWareShip;
end if;

create view itemWareShip (outdate, prId, prExt, numorder, nomnom, quant)
-- список номенклатуры, входящие в отгруженные готовые изделия.
-- несколько отгрузок порождают несколько строк с одинаковыми заказ-изделие-номенклатура.
-- учитываются изделия как с фиксированной, так и с вариантной номенклатурой.
as 
select outdate, io.prId, io.prExt, numorder, nomnom, round(fn.quantity, 5) as quant
from xpredmetybyizdeliaout io 
join itemWareFixe fn on io.prid = fn.productid
	union all
select outdate, io.prId, io.prExt, io.numorder, v.nomnom, round(p.quantity, 5) as quant
from xpredmetybyizdeliaout io 
join xvariantnomenc v on v.numorder = io.numorder and v.prid = io.prid and v.prext = io.prext
join sproducts p on p.productid = io.prid and v.nomnom = p.nomnom
;
