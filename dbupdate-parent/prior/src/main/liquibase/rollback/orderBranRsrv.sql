if exists (select 1 from sysviews where viewname = 'orderBranRsrv' and vcreator = 'dba') then
	drop view orderBranRsrv;
end if;

create view orderBranRsrv (numorder, nomnom, quant, date1, manager, client, note, werk, sm_zakazano, sm_paid, scope, status)
as
select 
	r.numorder, r.nomnom, (r.quant / n.perlist), r.date1, r.manager, r.client, r.note, r.werk, r.sm_zakazano, r.sm_paid, r.scope, r.status
from isumBranRsrv r
join sguidenomenk n on n.nomnom = r.nomnom
;
