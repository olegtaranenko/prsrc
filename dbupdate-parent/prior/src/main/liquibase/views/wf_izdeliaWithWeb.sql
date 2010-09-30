if exists (select 1 from sysviews where viewname = 'wf_izdeliaWithWeb' and vcreator = 'dba') then
	drop view wf_izdeliaWithWeb;
end if;

create view wf_izdeliaWithWeb 
as
select * from sguideproducts gp
where exists (
	select 1 from sguidenomenk n, sproducts p where p.productid = gp.prid and n.nomnom = p.nomnom and n.web = 'mat'
);