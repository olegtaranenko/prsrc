--=====================================================
--	Ship сокращение от Ship: отгружено по номенклатуре
--=====================================================

if exists (select 1 from sysviews where viewname = 'itemSellShip' and vcreator = 'dba') then
	drop view itemSellShip;
end if;

create view itemSellShip (numorder, nomnom, quant, date1)
-- Отгруженная номенклатура единым списком, включая продажи.
-- на каждую отгрузку или на каждое вхождение номенклатуры заказа через изделия - своя строчка
as 
select io.numorder, io.nomnom, io.quant * n.perlist, outdate
from baynomenkout io
join sguidenomenk n on n.nomnom = io.nomnom
;
