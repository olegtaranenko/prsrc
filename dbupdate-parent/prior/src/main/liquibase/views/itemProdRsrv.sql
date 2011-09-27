ALTER VIEW "DBA"."itemProdRsrv" (numorder, nomnom, quant, quant_rele)
as
select 
r.numdoc, r.nomnom, r.quantity, sum(isnull(d.quant, 0))
from sdmcrez r
left join sdmc d on d.numdoc = r.numdoc and d.nomnom = r.nomnom
join orders o on o.numorder = r.numdoc
group by r.numdoc, r.nomnom, r.quantity
