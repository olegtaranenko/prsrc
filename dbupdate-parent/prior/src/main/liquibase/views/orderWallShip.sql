if exists (select 1 from sysviews where viewname = 'orderWallShip' and vcreator = 'dba') then
	drop view orderWallShip;
end if;


create view orderWallShip (
	  outdate
	, numorder
	, type
	, cena
	, cost
	, quant
	, name
	, ventureId
	, statusid
	, werkid
	, edIzm
) as
select 
	  outdate
	, numorder
	, sum(distinct(type))
	, sum(isnull(round(quant * cenaEd , 2), 0))
	, sum(isnull(round(quant * costEd, 2), 0))
	, firmname
	, ventureid
from 
	itemWallShip po
group by 
	  outdate
	, numorder
	, firmname
	, ventureid