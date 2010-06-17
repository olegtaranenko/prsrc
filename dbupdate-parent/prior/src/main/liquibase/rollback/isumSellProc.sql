if exists (select 1 from sysviews where viewname = 'isumSellProc' and vcreator = 'dba') then
	drop view isumSellProc;
end if;

create view isumSellProc (numorder, nomnom, quant, statusid, firmname, ventureid, date1, date2, werk, manag)
as 
select 
	r.numorder, r.nomnom, sum(r.quant) as quant, r.statusid, r.firmname, r.ventureid, min(r.date1), max(r.date1), werk, manag
from itemSellProc r
group by r.numorder, r.nomnom, r.statusid, r.firmname, r.ventureid, werk, manag
having quant > 0
;
