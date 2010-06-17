--------------------------------- единая вью по всем заказам включая продажи ---------------------

if exists (select 1 from sysviews where viewname = 'all_orders' and vcreator = 'dba') then
	drop view all_orders;
end if;

create view all_orders (numorder, tp, xdate, statusid, firmName) 
-- заказы производства и продаж единым списком
as 
select numorder, 'orders', indate, statusid, f.name
from orders o
join guidefirms f on f.firmid = o.firmid
	union 
select numorder, 'bayorders', indate, statusid, f.name
from bayorders o
join bayguidefirms f on f.firmid = o.firmid
;

