if exists (select 1 from sysviews where viewname = 'itemBareOrde' and vcreator = 'dba') then
	drop view itemBareOrde;
end if;

create view itemBareOrde (numorder, nomnom, quant, cenaEd, prid, prext, quantEd)
-- только отдельная номенклатура заказа
as 

select po.numorder, po.nomnom, po.quant, po.cenaEd, null, null, 1
from xPredmetyByNomenk po;