if exists (select 1 from sysviews where viewname = 'itemWareFixe' and vcreator = 'dba') then
	drop view itemWareFixe;
end if;


create view itemWareFixe
-- ������ ������������� ����� ������� ������� � ���������� �������������.
as 
select * from 
sproducts p 
where not exists (
	select 1 from sguidevariant gv 
	where 
			p.productid = gv.productid 
		and p.xgroup = gv.xgroup 
		and not (gv.xgroup = '' 
				or (gv.xgroup != '' and gv.c = 1)
		)
)
;


