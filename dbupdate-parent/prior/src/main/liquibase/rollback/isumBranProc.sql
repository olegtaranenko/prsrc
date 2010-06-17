if exists (select 1 from sysviews where viewname = 'isumBranProc' and vcreator = 'dba') then
	drop view isumBranProc;
end if;


create view isumBranProc (numorder, nomnom, quant, date1, date2, scope, statusid, firmname, ventureid, werk, manag)
as
select r.numorder, r.nomnom, r.quant - isnull(s.quant, 0), r.date1, r.date2, 'p', r.statusid, r.firmname, r.ventureid, werk, manag
from isumProdProc r
left join isumProdShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
		union all
select r.numorder, r.nomnom, r.quant - isnull(s.quant, 0), r.date1, r.date2, 'b', r.statusid, r.firmname, r.ventureid, werk, manag
from isumSellProc r
left join isumSellShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
;
