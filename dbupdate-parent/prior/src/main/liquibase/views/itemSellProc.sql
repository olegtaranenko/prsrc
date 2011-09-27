ALTER VIEW "DBA"."itemSellProc" (numorder, nomnom, quant, statusid, firmname, ventureid, date1, werk, manag)
as 
select 
	r.numdoc, r.nomnom, r.quant, o.statusid, f.name, o.ventureid, d.xdate, 'Sell-Old', m.manag
from sdmc r
join sdocs d on r.numdoc = d.numdoc and r.numext = d.numext
join bayorders o on o.numorder = r.numdoc
join guidemanag m on m.managid = o.managid
join bayguidefirms f on f.firmid = o.firmid
where statusid < 6
--left join itemProdShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
