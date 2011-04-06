if exists (select 1 from sysviews where viewname = 'itemIzdeOrde' and vcreator = 'dba') then
	drop view itemIzdeOrde;
end if;

create view itemIzdeOrde (numorder, quant, cenaEd, prid)
-- только отдельная номенклатура заказа
as 

select pi.numorder, pi.quant, pi.cenaEd, pi.prId
from xPredmetyByIzdelia pi;