if exists (select 1 from sysviews where viewname = 'isumWareOrde' and vcreator = 'dba') then
	drop view isumWareOrde;
end if;

create view isumWareOrde (numorder, nomnom, quant, quantEd)
as 
select 
	numorder, nomnom, sum(quant) as quant, quantEd
from itemWareOrde
group by numorder, nomnom, quantEd
;


