if exists (select 1 from sysviews where viewname = 'isumBranShip' and vcreator = 'dba') then
	drop view isumBranShip;
end if;

create view isumBranShip (numorder, nomnom, quant)

as 
-- объединение все вхождения номенклатуры в заказ через изделия или отгрузки в одну строку
-- имеем общее к-во по отгруженной номенклатуре всего заказа,
-- количество для нештучной ном-ры - в производственных единицах (дм)

select numorder, nomnom, sum(quant) as quant
from
	itemBranShip
group by
	numorder, nomnom
;
