ALTER VIEW "DBA"."itemSellShip" (numorder, nomnom, quant, date1)
-- Отгруженная номенклатура единым списком, включая продажи.
-- на каждую отгрузку или на каждое вхождение номенклатуры заказа через изделия - своя строчка
as 
select io.numorder, io.nomnom, io.quant * n.perlist, outdate
from baynomenkout io
join sguidenomenk n on n.nomnom = io.nomnom
