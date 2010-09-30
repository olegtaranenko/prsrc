if exists (select 1 from sysviews where viewname = 'wf_izdeliaForSell' and vcreator = 'dba') then
	drop view wf_izdeliaForSell;
end if;

create view wf_izdeliaForSell 
as
select * from sguideproducts gp
where not exists (
	select 1 from sguidenomenk n, sproducts p 
	where 
		p.productid = gp.prid 
		and n.nomnom = p.nomnom 
		and not (n.perlist = 1 or n.web <> 'mat')
);