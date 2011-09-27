ALTER VIEW "DBA"."itemProdProc" (numorder, nomnom, quant, statusid, firmname, ventureid, date1, werk, manag)
as 
select 
	r.numdoc, r.nomnom, r.quant, o.statusid, f.name, o.ventureid, d.xdate, werkCodeEN, manag
from sdmc r
join sdocs d on r.numdoc = d.numdoc and r.numext = d.numext
join orders o on o.numorder = r.numdoc
join guidefirms f on f.firmid = o.firmid
join guidemanag m on m.managid = o.managid
join guidewerk w on w.werkid = o.werkid
where o.statusid < 6
--left join itemProdShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
