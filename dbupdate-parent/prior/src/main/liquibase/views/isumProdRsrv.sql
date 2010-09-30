if exists (select 1 from sysviews where viewname = 'isumProdRsrv' and vcreator = 'dba') then
	drop view isumProdRsrv;
end if;

create view isumProdRsrv (numorder, nomnom, quant, status, date1, manager, client, note, werk, sm_zakazano, sm_paid)
as
select 
r.numorder, r.nomnom, r.quant - r.quant_rele as x_quant, s.status, o.indate
, m.manag, f.name, o.product, w.werkCodeEN, o.ordered, o.paid
from itemProdRsrv r
join orders o on o.numorder = r.numorder
join guidestatus s on s.statusid = o.statusid and s.werkId = o.werkId
left join guidemanag m on m.managid = o.managid
left join guidefirms f on f.firmid = o.firmid
left join guidewerk w on w.werkid = o.werkid
where abs(round(x_quant, 2)) > 0.01
and o.werkId = 2;
