begin
	declare v_count int; 
	declare v_server varchar(31);
	declare v_seria_id_remote int;
	declare v_seria_remote varchar(63);

	set v_server = 'accountn';
	set v_count = 0;
	for c_gi as gi dynamic scroll cursor for
		select p.prId as r_prid, p.prSeriaId as r_prseriaid, p.id_inv as r_product_id, s.id_inv as r_seria_id, p.prname as r_prname, s.serianame as r_serianame
		from sguideproducts p
		join sguideseries s on s.seriaid = p.prseriaid
	do
		set v_seria_id_remote = select_remote(v_server, 'inv', 'belong_id', 'id = ' + convert(varchar(10), r_product_id));
		if v_seria_id_remote <> r_seria_id then
			set v_count = v_count+1;
--			set v_seria_remote = select_remote(v_server, 'inv', 'nm', 'id = ' + convert(varchar(10), v_seria_id_remote));
--			message '''' + r_prname + '''', '; ', r_prid, '; ', r_seria_id, '; ', v_seria_id_remote, ';', r_serianame, ';', v_seria_remote  to client;

			call update_host('inv', 'belong_id', convert(varchar(20), r_seria_id), 'id = ' + convert(varchar(20), r_product_id));
		end if;
	end for;
	message 'total rows: ', v_count to client;
end;

