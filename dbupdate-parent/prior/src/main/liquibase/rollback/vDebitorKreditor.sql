-- Отчет А: строки дебитор/кредитор

if exists (select 1 from sysviews where viewname = 'vDebitorKreditor' and vcreator = 'dba') then
	drop view vDebitorKreditor;
end if;

create 
    view vDebitorKreditor (numorder, name, k, d, type)
as

SELECT numorder, f.name, (isnull(if isnull(paid, 0) > isnull(shipped, 0) then isnull(paid, 0) - isnull(shipped, 0) endif , 0)) AS k
, (isnull(if isnull(paid, 0) < isnull(shipped, 0) then isnull(shipped, 0) - isnull(paid, 0) endif , 0)) AS d, 't'
--, ordered, paid, shipped, statusid
from orders o
join guidefirms f on o.firmid = f.firmid
where (k != 0 or d != 0)
and statusid < 6
        union all
SELECT o.numorder, f.name, (isnull(if isnull(paid, 0) > isnull(cenaTotal, 0) then isnull(paid, 0) - isnull(cenaTotal, 0) endif , 0)) AS k
, (isnull(if isnull(paid, 0) < isnull(cenaTotal, 0) then isnull(cenaTotal, 0) - isnull(paid, 0) endif , 0)) AS d, 'b'
--, ordered, paid, cenaTotal
from bayorders o
join bayguidefirms f on o.firmid = f.firmid
left join orderSellShip ot on ot.numorder = o.numorder
where (k != 0 or d != 0)
and o.statusid < 6
;
