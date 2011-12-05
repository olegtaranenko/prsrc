insert into sdmcrez (numdoc, nomnom, quantity)
select iwo.numorder, iwo.nomnom, isnull(iwo.quant, 0) - isnull(rez.quantity, 0)
from isumwareorde iwo
join sguidenomenk n on n.nomnom = iwo.nomnom
left join sdmcrez rez on rez.numdoc = iwo.numorder and rez.nomnom = iwo.nomnom
where 
not exists (select 1 from sdmcrez r where r.numdoc = iwo.numorder and r.nomnom = iwo.nomnom and r.quantity = iwo.quant)
and iwo.numorder = ${numorder}
