ALTER VIEW "DBA"."isumBranRsrv" (numorder, nomnom, quant, status, date1, manager, client, note, werk, sm_zakazano, sm_paid, scope)
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
