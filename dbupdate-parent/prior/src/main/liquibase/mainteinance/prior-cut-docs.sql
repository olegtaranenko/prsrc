begin

	declare v_count int; 

	set v_count = 0;


	create table #tdocs (ventureId int, numdoc int, id_jmat int, primary key(id_jmat, ventureId));

	CREATE 
--	UNIQUE 
	INDEX [tdocs_jmat] ON #tdocs([id_jmat]);

		insert into #tdocs (ventureId, id_jmat) 
		select 1, id as r_id
		from jmat_accountn as j with (FASTFIRSTROW)
		union all
		select 2, id as r_id
		from jmat_markmaster as j with (FASTFIRSTROW)
		union all
		select 3,id as r_id
		from jmat_stime as j with (FASTFIRSTROW);
	message @@rowcount to client;
	delete from sDocs where not exists	(select 1 from #tdocs t where t.id_jmat = sDocs.id_jmat);

end;
