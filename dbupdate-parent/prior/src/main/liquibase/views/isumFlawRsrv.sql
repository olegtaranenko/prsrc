ALTER VIEW "DBA"."isumFlawRsrv" (numorder, nomnom, quant, status, date1, manager, client, note, ceh, sm_zakazano, sm_paid)
as
select
r.numorder, r.nomnom, r.quant, null, r.date1, null, 'Списание', null, r.note, null, null
from itemFlawRsrv r