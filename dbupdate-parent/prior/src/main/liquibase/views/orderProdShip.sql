if exists (select 1 from sysviews where viewname = 'orderProdShip' and vcreator = 'dba') then
	drop view orderProdShip;
end if;



create view orderProdShip (
	numorder
	, quant
	, date1
)
-- Отгруженная номенклатура единым списком, включая продажи.
-- на каждую отгрузку или на каждое вхождение номенклатуры заказа через изделия - своя строчка
as 

select 
