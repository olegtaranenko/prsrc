if exists (select 1 from sysviews where viewname = 'isumSellRsrv' and vcreator = 'dba') then
	drop view isumSellRsrv;
end if;

create view isumSellRsrv (numorder, nomnom, quant, status, date1, manager, client, note, ceh, sm_zakazano, sm_paid)
as
select
	r.numorder, r.nomnom, r.quant - r.quant_rele as x_quant, s.status, r.date1
	, m.manag, f.name, null, 'Продажа', o.ordered, o.paid
from itemSellRsrv r
join bayorders o on o.numorder = r.numorder
join guidestatus s on s.statusid = o.statusid
left join guidemanag m on m.managid = o.managid
left join bayguidefirms f on f.firmid = o.firmid
where abs(round(x_quant, 2)) > 0.01
;
