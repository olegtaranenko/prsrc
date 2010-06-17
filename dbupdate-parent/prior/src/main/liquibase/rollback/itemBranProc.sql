if exists (select 1 from sysviews where viewname = 'itemBranProc' and vcreator = 'dba') then
	drop view itemBranProc;
end if;


create view itemBranProc (numorder, nomnom, quant, scope, statusid)
as 
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemProdProc r
		union all
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemSellProc r
;
