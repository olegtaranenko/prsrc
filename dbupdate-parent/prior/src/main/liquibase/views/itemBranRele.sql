ALTER VIEW "DBA"."itemBranRele" (numorder, nomnom, quant, scope, date1, statusid)
as 
select r.numdoc, r.nomnom, r.quant
, if o.numorder is not null then 'p' else if bo.numorder is not null then 'b' else '?' endif endif 
, d.xdate
, if o.numorder is not null then o.statusid else if bo.numorder is not null then bo.statusid else null endif endif 
from sdmc r
join sdocs d on d.numdoc = r.numdoc and d.numext = r.numext
left join orders o on r.numdoc = o.numorder
left join bayorders bo on r.numdoc = bo.numorder
where r.numext < 254
