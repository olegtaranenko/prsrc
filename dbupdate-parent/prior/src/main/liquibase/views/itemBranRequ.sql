if exists (select 1 from sysviews where viewname = 'itemBranRequ' and vcreator = 'dba') then
	drop view itemBranRequ;
end if;


create view itemBranRequ (
	  numorder
	, nomnom
	, quant
	, scope
	, statusid
	, werkId
)
as 
select 
	  r.numdoc
	, r.nomnom
	, r.curquant * n.perlist
	, case werkId when 1 then 'p' else 'b' end case
	, o.statusid
	, o.werkId
from sDmcRez r
join Orders o on r.numdoc = o.numorder
join sGuideNomenk n on r.nomnom = n.nomnom
where r.curQuant > 0;
