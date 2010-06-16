begin

	declare v_count int; 

	set v_count = 0;


	create table #torders (ventureId int, numorder int, id_jscet int, primary key(id_jscet, ventureId));

	CREATE 
--	UNIQUE 
	INDEX [torders_jscet] ON #tOrders([id_jscet]);

		insert into #torders (ventureId, id_jscet) 
		select 1, id as r_id
		from jscet_accountn as j with (FASTFIRSTROW)
		union all
		select 2, id as r_id
		from jscet_markmaster as j with (FASTFIRSTROW)
		union all
		select 3,id as r_id
		from jscet_stime as j with (FASTFIRSTROW);

	delete from orders where not exists	(select 1 from #torders t where t.id_jscet = orders.id_jscet);
	delete from bayorders where not exists	(select 1 from #torders t where t.id_jscet = bayorders.id_jscet);

end;

--select * from orders order by numorder desc
--delete from orders where numorder = 10061005
--commit


--select count(*) from orders;
--rollback


/*
create table torders (ventureId int, numorder int, primary key(ventureId, numorder));

DROP STATISTICS torders;
rollback

insert into torders (ventureid, numorder)
select 1, o.numorder
from orders o
where exists (select 1 from jscet_accountn js where js.id = o.id_jscet);

DROP STATISTICS torders;

insert into torders (ventureid, numorder)
select 2, o.numorder
from orders o
where exists (select 1 from jscet_markmaster js where js.id = o.id_jscet);

insert into torders (ventureid, numorder)
select 3, o.numorder
from orders o
where exists (select 1 from jscet_stime js where js.id = o.id_jscet);

update orders set id_jscet = -1 where not exists (select 1 from #torders t where t.numorder = orders.numorder);

delete from orders where id_jscet = -1;


drop table torders;

select count(*) from torders;
--select count(*) from orders;

commit;

*/