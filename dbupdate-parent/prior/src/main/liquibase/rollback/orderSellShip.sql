if exists (select 1 from sysviews where viewname = 'orderSellShip' and vcreator = 'dba') then
	drop view orderSellShip;
end if;



create view orderSellShip (numorder, cenaTotal, statusid)
as 
	select o.numorder, sum(r.intQuant * po.quant) as cenaTotal, o.statusid
	from bayorders o
	join sDMCrez r on r.numDoc = o.numorder
	join baynomenkout po on po.numorder = o.numorder and po.nomnom = r.nomnom
	group by o.numorder, o.statusid
;
