ALTER VIEW "DBA"."orderBranProc"(numorder, sm_processed, statusid, date1, date2, firmname, scope, werk, manag)
as
select r.numorder
	, sum(r.quant * n.cost / n.perlist) as quant
	, r.statusid , min(r.date1), max(r.date2), r.firmname, r.scope, werk, manag
from isumBranProc r
join sguidenomenk n on n.nomnom = r.nomnom
group by r.numorder, r.statusid, r.firmname, r.scope, werk, manag
having round(quant, 2) > 0
