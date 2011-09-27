ALTER VIEW "DBA"."itemBranProc" (numorder, nomnom, quant, scope, statusid)
as 
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemProdProc r
		union all
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemSellProc r
