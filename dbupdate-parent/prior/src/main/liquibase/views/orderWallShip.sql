if exists (select 1 from sysviews where viewname = 'orderWallShip' and vcreator = 'dba') then
	drop view orderWallShip;
end if;


create view orderWallShip (
	  outdate
	, numorder
	, type
	, cenaTotal
	, costTotal
	, name
	, ventureId
	, werkid
) as
select 
	  po.outdate
	, po.numorder
	, sum(distinct(po.type))
	, sum(isnull(round(po.quant * po.cenaEd , 2), 0))
	, sum(isnull(round(po.quant * po.costEd, 2), 0))
	, po.firmname
	, po.ventureid
	, o.werkId
from 
	itemWallShip po
join 	
	orders o
		on o.numorder = po.numorder
group by 
	  po.outdate
	, po.numorder
	, po.firmname
	, po.ventureid
	, o.werkId