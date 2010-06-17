if exists (select 1 from sysviews where viewname = 'itemBranRequ' and vcreator = 'dba') then
	drop view itemBranRequ;
end if;


create view itemBranRequ (numorder, nomnom, quant, scope, statusid)
as 
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemProdRequ r
		union all
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemSellRequ r
;
