if exists (select 1 from sysviews where viewname = 'itemFlawRsrv' and vcreator = 'dba') then
	drop view itemFlawRsrv;
end if;

create view itemFlawRsrv (numorder, nomnom, quant, date1, note)
as
select d.numdoc, r.nomnom, r.quantity, d.xdate, d.note
from sdocs d
join sdmcrez r on r.numdoc = d.numdoc
where d.numext = 0
;

if exists (select 1 from sysviews where viewname = 'isumFlawRsrv' and vcreator = 'dba') then
	drop view isumFlawRsrv;
end if;

create view isumFlawRsrv (numorder, nomnom, quant, status, date1, manager, client, note, ceh, sm_zakazano, sm_paid)
as
select
r.numorder, r.nomnom, r.quant, null, r.date1, null, 'Списание', null, r.note, null, null
from itemFlawRsrv r
;
