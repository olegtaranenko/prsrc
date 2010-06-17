if exists (select 1 from sysviews where viewname = 'orderWareShip' and vcreator = 'dba') then
	drop view orderWareShip;
end if;

create view orderWareShip (outdate, prId, prExt, numorder, costEd)
-- себестоимость готовых изделий входящие в заказы, которые уже отгруженны.
as 
select outdate, prid, prext, numorder, sum(round(n.cost * io.quant / n.perlist, 2))
from itemWareShip io 
join sGuideNomenk n on io.nomnom = n.nomnom
group by outdate, prid, prext, numorder
;
