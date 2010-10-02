if exists (select 1 from sysviews where viewname = 'orderWareOrde' and vcreator = 'dba') then
	drop view orderWareOrde;
end if;

create view orderWareOrde (numorder, nomnom, quant, prid, prext)
as 
select io.numorder, null, io.quant, io.prid, io.prext
from xpredmetybyizdelia io 
	union all
select po.numorder, po.nomnom, po.quant, null, null
from xpredmetybynomenk po
;

