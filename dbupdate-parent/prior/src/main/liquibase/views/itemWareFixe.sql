if exists (select 1 from sysviews where viewname = 'itemWareFixe' and vcreator = 'dba') then
	drop view itemWareFixe;
end if;


create view itemWareFixe
-- список фиксированной части готовых изделий с вариантной номенклатурой.
as 
select p.* from 
sproducts p 
left join sguidevariant gv on p.productid = gv.productid and p.xgroup = gv.xgroup 
where p.xgroup = '' or isnull(gv.c, 1) = 1
;
