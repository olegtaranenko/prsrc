--=====================================================
--		Proc	сокращение от Processed:"незавершека"
--=====================================================

if exists (select 1 from sysviews where viewname = 'itemProdProc' and vcreator = 'dba') then
	drop view itemProdProc;
end if;

create view itemProdProc (numorder, nomnom, quant, statusid, firmname, ventureid, date1, werk, manag)
as 
select 
	r.numdoc, r.nomnom, r.quant, o.statusid, f.name, o.ventureid, d.xdate, werkName, manag
from sdmc r
join sdocs d on r.numdoc = d.numdoc and r.numext = d.numext
join orders o on o.numorder = r.numdoc
join guidefirms f on f.firmid = o.firmid
join guidemanag m on m.managid = o.managid
join guidewerk w on w.werkid = o.werkid
where o.statusid < 6
--left join itemProdShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
;
