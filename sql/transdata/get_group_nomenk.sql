if exists (select 1 from sysprocedure where proc_name = 'get_group_nomenk') then
	drop procedure get_group_nomenk;
end if;


CREATE procedure get_group_nomenk(
		p_klassid integer
)
begin

	declare v_lvl integer;

	create table #tmp_klass(lvl integer, id integer);

	set v_lvl = 0;

	insert into #tmp_klass (lvl, id) select 0, p_klassid;

	branch: loop
		insert into #tmp_klass (lvl, id)
			select v_lvl + 1, k.klassId
			from sguideklass k
			join #tmp_klass t on t.id = k.parentKlassId and t.lvl = v_lvl;

		if @@rowcount = 0 then
			leave branch;
		end if;
		set v_lvl = v_lvl + 1;
	end loop;

	select n.* from sguidenomenk n
	join #tmp_klass t on n.klassid = t.id;

	drop table #tmp_klass;

end;

