if exists (select 1 from sysviews where viewname = 'itemSellProc' and vcreator = 'dba') then
	drop view itemSellProc;
end if;

create view itemSellProc (numorder, nomnom, quant, statusid, firmname, ventureid, date1, ceh, manag)
as 
select 
	r.numdoc, r.nomnom, r.quant, o.statusid, f.name, o.ventureid, d.xdate, 'Sell', m.manag
from sdmc r
join sdocs d on r.numdoc = d.numdoc and r.numext = d.numext
join bayorders o on o.numorder = r.numdoc
join guidemanag m on m.managid = o.managid
join bayguidefirms f on f.firmid = o.firmid
where statusid < 6
--left join itemProdShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
;
