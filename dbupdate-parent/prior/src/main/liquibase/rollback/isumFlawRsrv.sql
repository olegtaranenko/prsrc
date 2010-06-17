if exists (select 1 from sysviews where viewname = 'isumFlawRsrv' and vcreator = 'dba') then
	drop view isumFlawRsrv;
end if;

create view isumFlawRsrv (numorder, nomnom, quant, status, date1, manager, client, note, werk, sm_zakazano, sm_paid)
as
select
r.numorder, r.nomnom, r.quant, null, r.date1, null, 'Списание', null, r.note, null, null
from itemFlawRsrv r
;
