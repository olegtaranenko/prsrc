if exists (select 1 from sysviews where viewname = 'isumBranRsrv' and vcreator = 'dba') then
	drop view isumBranRsrv;
end if;

create view isumBranRsrv (numorder, nomnom, quant, status, date1, manager, client, note, ceh, sm_zakazano, sm_paid, scope)
as
select 
	*, 'p'
from 
	isumProdRsrv
		union all
select 
	*, 'b'
from 
	isumSellRsrv
		union all
select 
	*, 'f'
from 
	isumFlawRsrv
;
