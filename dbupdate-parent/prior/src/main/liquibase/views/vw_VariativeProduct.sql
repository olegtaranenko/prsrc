if exists (select 1 from sysviews where viewname = 'vw_VariativeProduct' and vcreator = 'dba') then
	drop view vw_VariativeProduct;
end if;

create view vw_VariativeProduct 
as

select distinct gp.PrId, gv.xgroup
from sguideproducts gp
left join sguidevariant gv on gp.prid = gv.productid 
where isnull(gv.c, 1) > 1 and gv.xgroup <> ''
