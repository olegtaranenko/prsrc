if exists (select 1 from sysviews where viewname = 'isumProdProc' and vcreator = 'dba') then
	drop view isumProdProc;
end if;

create view isumProdProc (numorder, nomnom, quant, statusid, firmname, ventureid, date1, date2, ceh, manag)
as 
select 
	r.numorder, r.nomnom, sum(r.quant) as quant, r.statusid, r.firmname, r.ventureid, min(r.date1), max(date1), ceh, manag
from itemProdProc r
group by r.numorder, r.nomnom, r.statusid, r.firmname, r.ventureid, ceh, manag
having quant > 0
;
