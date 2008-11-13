-- синхронизировать справочник фирм с бухгалтерскими базами комтеха.

begin 
	declare v_exists int;
	declare v_zakaz_id integer;
	declare v_params long varchar;
	declare v_table varchar(128);

	select id_voc_names into v_zakaz_id from guidefirms where firmid = 0;

	for x as xx dynamic scroll cursor for
	select firmid as r_firmid, id_voc_names as r_id_voc_names, name as r_firmname
	from guidefirms
	do
		for y as yy dynamic scroll cursor for
		select sysname as r_servername from guideventure
		do
			set v_table = 'voc_names_' + r_servername;
			set v_params = 'select count(*) into v_exists from ' + v_table + ' where id = ' + convert(varchar(20), r_id_voc_names);
			execute immediate v_params;
			if v_exists = 0 then
				message 'skipped: server = ', r_servername, ', firmid = ', r_firmid, ' id_voc_names = ', r_id_voc_names to client;
				set v_params =
					 convert(varchar(20), r_id_voc_names)
					+ ', '''''+ substring(r_firmname, 1, 203) + ''''''
					+ ', ' + convert(varchar(20), v_zakaz_id);
			
				call insert_remote(r_servername, 'voc_names', 'id, nm, belong_id', v_params);
			end if;
		end for;
	end for;
end;
