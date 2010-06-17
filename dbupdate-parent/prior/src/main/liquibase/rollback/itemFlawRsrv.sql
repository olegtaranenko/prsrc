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

