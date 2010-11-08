if exists (select 1 from sysviews where viewname = 'wf_izdeliaVmtOnly' and vcreator = 'dba') then
	drop view wf_izdeliaVmtOnly;
end if;

create view wf_izdeliaVmtOnly 
as
select * 
from (
select productid, max(web) as vmtOnly
from
	(	
	select p.productid, case web when 'vmt' then 1 else 0 end as web
		from sguidenomenk n, sproducts p where n.nomnom = p.nomnom 
	group by p.productid, web
	) a
group by productid 
having count(*) = 1
) b
where vmtOnly = 1
