if exists (select 1 from sysviews where viewname = 'isumBranRele' and vcreator = 'dba') then
	drop view isumBranRele;
end if;

create view isumBranRele (numorder, nomnom, quant, scope, date1, date2, statusid)
as 
select r.numorder, r.nomnom, sum(r.quant) as quant
, r.scope
, min(r.date1) as date1, max(r.date1) as date2
, r.statusid
from itemBranRele r
where r.statusid < 6
group by r.numorder, r.nomnom, r.scope, r.statusid
;

