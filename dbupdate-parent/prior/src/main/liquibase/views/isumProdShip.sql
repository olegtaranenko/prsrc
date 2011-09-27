ALTER VIEW "DBA"."isumProdShip" (numorder, nomnom, quant, date1, date2)
as 
select numorder, nomnom, sum(quant) as quant, min(date1), max(date1)
from itemProdShip
group by 	
	numorder, nomnom
