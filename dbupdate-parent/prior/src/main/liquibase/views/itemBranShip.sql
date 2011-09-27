ALTER VIEW "DBA"."itemBranShip" (numorder, nomnom, quant, date1, scope)
as 
select r.numorder, r.nomnom, r.quant, r.date1, 'p'
from itemProdShip r
	union all
select r.numorder, r.nomnom, r.quant, r.date1, 'b'
from itemSellShip r
