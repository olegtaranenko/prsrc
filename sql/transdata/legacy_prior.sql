if exists (select 1 from sysprocedure where proc_name = 'legacy_size_izd') then
	drop procedure legacy_size_izd;
end if;

-- Из-за ошибки при трансдатации не попали размеры для готовых изделий
-- Эта процедура исправляет эту ошибку
CREATE PROCEDURE legacy_size_izd()
begin
	declare glbId integer;
	declare v_table_name varchar(100);
	declare v_fields varchar(1000);
	declare v_values varchar(1000);
	declare v_belong_root varchar(100);

	message 'legacy_size_izd() started ...' to client;

	set v_table_name = 'size';
	set glbId = get_nextId (v_table_name);

	-- размеры, которые не были трансдатированы
	for aCursor as b dynamic scroll cursor for
		select distinct prsize as sz from sguideproducts g
		where not exists (select 1 from size s where s.name = g.prsize)
	do
		insert into size (id_size, name) values (glbId, sz);
		call insert_host(v_table_name, 'id, nm', convert(varchar(10), glbId) + ', ''''' + sz +'''''');
		set glbId = glbId + 1;
	end for;

	-- изменить единицу измерения в комтехах для изделий,
	-- которая трансдатирована была ошибочно

	for aCursor as a dynamic scroll cursor for
		select prId as r_prid
			, prsize as r_prsize
			, id_size as r_id_size
			, id_inv as r_id_inv
		from sguideproducts g
		join size s on s.name = g.prsize
	do
		call update_host('inv', 'id_size', convert(varchar(20), r_id_size), 'id = '+ convert(varchar(20), r_id_inv));
		call update_host('inv', 'id_size', convert(varchar(20), r_id_size), 'belong_id = '+ convert(varchar(20), r_id_inv));
	end for;
	
	message 'legacy_size_izd() ended successful.' to client;

end;


if exists (select 1 from sysprocedure where proc_name = 'legacy_guides') then
	drop procedure legacy_guides;
end if;

CREATE PROCEDURE legacy_guides()
begin

	declare glbId integer;
	declare v_table_name varchar(100);
	declare v_fields varchar(1000);
	declare v_values varchar(1000);
	declare v_belong_root varchar(100);

	message 'legacy_guides() started ...' to client;
	
	set v_table_name = 'edizm';
	set glbId = get_nextId (v_table_name);

	for aCursor as a dynamic scroll cursor for
		select distinct ed_izmer as ed_izmer from sguidenomenk
			union 
		select distinct ed_izmer2 as ed_izmer from sguidenomenk
	do
		insert into edizm (id_edizm, name) values (glbId, ed_izmer);
		call insert_host(v_table_name, 'id, nm', convert(varchar(10), glbId) + ', ''''' + ed_izmer + '''''');
		set glbId = glbId + 1;
	end for;


	set v_table_name = 'size';
	set glbId = get_nextId (v_table_name);

	for aCursor as b dynamic scroll cursor for
	select distinct size as sz from sguidenomenk
	do
		insert into size (id_size, name) values (glbId, sz);
		call insert_host(v_table_name, 'id, nm', convert(varchar(10), glbId) + ', ''''' + sz +'''''');
		set glbId = glbId + 1;
	end for;
	message 'legacy_guides() ended successful.' to client;
end;



if exists (select 1 from sysprocedure where proc_name = 'legacy_sklad') then
	drop procedure legacy_sklad;
end if;

CREATE PROCEDURE legacy_sklad()
begin
	
	declare v_sklad_id integer;
	declare v_folder_id integer;
	
	message 'legacy_sklad() started ...' to client;
	-- Получить общий id склада
	set v_sklad_id = get_nextid ('voc_names');
	
	for aCursor as a dynamic scroll cursor for
	    select 
			gs.sourceName as r_name
			, id_voc_names 
	    from sguidesource gs 
		where sourceid < -1000
		and id_voc_names is null

	FOR UPDATE 
	DO

		select id into v_folder_id from voc_names_stime where nm = 'Склады';
	
		
		call insert_host('voc_names', 'id, nm, belong_id', 
			convert(varchar(20), v_sklad_id)
			+ ', '''''+ r_name + ''''''
			+ ', ' + convert(varchar(20), v_folder_id)
		);
	
		UPDATE sguidesource set id_voc_names = v_sklad_id WHERE CURRENT OF a ;
		set v_sklad_id = v_sklad_id + 1;
	end for;
	message 'legacy_sklad() ended successful.' to client;
end;



if exists (select 1 from sysprocedure where proc_name = 'legacy_zatr') then
	drop procedure legacy_zatr;
end if;

CREATE PROCEDURE legacy_zatr()
begin
	
	declare v_zatr_id integer;
	declare v_folder_id integer;
	
	message 'legacy_zatr() started ...' to client;
	-- Получить общий id склада
	set v_zatr_id = get_nextid ('voc_names');
	
	for aCursor as a dynamic scroll cursor for
	    select 
			gs.sourceName as r_name
			, id_voc_names 
	    from sguidesource gs 
        where sourceid between -1000 and -1
		and id_voc_names is null

	FOR UPDATE 
	DO

		select id into v_folder_id from voc_names_stime where nm = 'Объекты затрат';
	
		
		call insert_host('voc_names', 'id, nm, belong_id', 
			convert(varchar(20), v_zatr_id)
			+ ', '''''+ r_name + ''''''
			+ ', ' + convert(varchar(20), v_folder_id)
		);
	
		UPDATE sguidesource set id_voc_names = v_zatr_id WHERE CURRENT OF a ;
		set v_zatr_id = v_zatr_id + 1;
	end for;
	message 'legacy_zatr() ended successful.' to client;
end;


if exists (select 1 from sysprocedure where proc_name = 'legacy_firms') then
	drop procedure legacy_firms;
end if;

CREATE PROCEDURE legacy_firms()
begin
	
	declare v_firms_id integer;
	declare v_folder_id integer;
	declare v_postav_id integer;
	declare v_zakaz_id integer;
	declare v_params varchar(1000);
	declare postav_name varchar(32);
	declare zakaz_name varchar(32);
	declare v_gemacht datetime;
	declare v_phone varchar(37);
	declare v_fio varchar(98);
	declare v_rem varchar(98);
	declare v_email varchar(98);

	message 'legacy_firms() started ...' to client;
	set postav_name = 'Поставщики';
	set zakaz_name = 'Заказчики';
	
	select trans_date into v_gemacht from system;

--	call slave_select_stime(v_gemacht, 'jmat', 'count(*)', 'osn = ''' + v_legacy + '''');
	if v_gemacht is not null then
		message 'Унаследованные фирмы уже загружены' to client;
		return;
	end if;
	
	-- корень
	select id into v_folder_id from voc_names_stime where nm = 'Сторонние организации' and belong_Id = 0;

	-- Получить id папки поставщиков
	select id into v_postav_id from voc_names_stime where nm = postav_name and belong_id = v_folder_id and is_group = 1;
	if v_postav_id is null then
		set v_postav_id = get_nextid ('voc_names');
		call insert_host('voc_names', 'id, nm, belong_id, is_group', 
			convert(varchar(20), v_postav_id)
			+ ', ''''' + postav_name + ''''''
			+ ', ' + convert(varchar(20), v_folder_id)
			+ ', 1'
		);

		UPDATE sguidesource set id_voc_names = v_postav_id WHERE sourceid = 0;
	end if;

	-- Получить id папки заказчиков
	select id into v_zakaz_id from voc_names_stime where nm = zakaz_name and belong_id = v_folder_id and is_group = 1;
	if v_zakaz_id is null then
		set v_zakaz_id = get_nextid ('voc_names');

		call insert_host('voc_names', 'id, nm, belong_id, is_group', 
			convert(varchar(20), v_zakaz_id)
			+ ', '''''+ zakaz_name + ''''''
			+ ', ' + convert(varchar(20), v_folder_id)
			+ ', 1'
		);

	end if;


	-- id первой фирмы
	set v_firms_id = get_nextid ('voc_names');

	for aCursor as a1 dynamic scroll cursor for
		select 
			gs.sourceName as r_name, 
			gs.Phone as p_phone, 
			gs.Email as p_email
		from sguidesource gs
		where sourceid > 0
		and id_voc_names is null
	FOR UPDATE DO
		call insert_host('voc_names', 'id, nm, belong_id, is_group', 
			convert(varchar(20), v_firms_id)
			+ ', '''''+ r_name + ''''''
			+ ', ' + convert(varchar(20), v_postav_id)
			+ ', 0'
		);
		update sguidesource set id_voc_names = v_firms_id where current of a1;
		set v_firms_id = v_firms_id + 1;
	end for;
	

	for aCursor as b1 dynamic scroll cursor for
		select 
			f.firmId as f_id
			,b.firmId as b_id
			,isnull(f.name, b.name) as r_name
			,f.fio as f_fio
			,b.fio as b_fio
			,f.phone as f_phone
			,b.phone as b_phone
			,f.email as f_email
			,b.email as b_email
		from guidefirms f
		full join bayguidefirms b on b.name = f.name 
	DO

		set v_fio = isnull(f_fio, b_fio);
		if b_fio is not null and v_fio != b_fio then
			set v_fio = v_fio + ', ' + b_fio;
		end if;
		set v_email = isnull(f_email, b_email);
		if b_email is not null and v_email != b_email then
			set v_email = v_email + ', ' + b_email;
		end if;

		set v_phone = isnull(f_phone, b_phone);
		if b_phone is not null and v_phone != b_phone then
			set v_phone = v_phone + ',' + b_phone;
		end if;

			set v_params =
				 convert(varchar(20), v_firms_id)
				+ ', '''''+ substring(r_name,1,203) + ''''''
				+ ', ''''' + substring(v_phone,1,37) + ''''''
				+ ', ''''' + substring(v_fio,1,98) + ''''''
				+ ', ''''' + substring(v_email,1,98) + ''''''
				+ ', ' + convert(varchar(20), v_zakaz_id);

		call insert_host('voc_names', 'id, nm, phone, address, address_fact, belong_id', v_params);

		if f_id is not null then
			UPDATE guidefirms set id_voc_names = v_firms_id WHERE firmid = f_id;
		end if;

		if b_id is not null then
			UPDATE bayguidefirms set id_voc_names = v_firms_id WHERE firmid = b_id;
		end if;

		set v_firms_id = v_firms_id + 1;

	end for;

	UPDATE guidefirms set id_voc_names = v_zakaz_id WHERE firmId = 0;
	UPDATE bayguidefirms set id_voc_names = v_zakaz_id WHERE firmId = 0;
	
	message 'legacy_firms() ended successful.' to client;
			
end;

if exists (select 1 from sysprocedure where proc_name = 'legacy_inv') then
	drop function legacy_inv;
end if;

create PROCEDURE 
	legacy_inv()
begin
	declare glbId integer;
	declare v_table_name varchar(100);
	declare v_fields varchar(1000);
	declare v_values varchar(1000);
	declare v_belong_root varchar(100);
	declare v_id_size varchar(32);
	declare v_id_inv integer;
	declare v_id_edizm1 integer;
	declare v_id_currency integer;
	declare v_nm varchar(200);
	declare v_blank char(1);
	
	message 'legacy_inv() started ...' to client;
	

	set v_table_name = 'inv';
	set glbId = get_nextId (v_table_name);
	set v_id_currency = system_currency();

--	set glbId = 5854;
--	raiserror 17000 'v_values = %1!', v_values;

	update sguideklass k set k.id_inv = 4 where klassid = 0; --'Материалы';
	update sguideseries set id_inv = 5 where seriaid = 0; --'Готовые изделия';

	
	set v_fields = 'id, nm, is_group';
	for aCursor as a dynamic scroll cursor for
		select 
			  klassname as p_name
		from sguideklass gk
        where klassid > 0
		and id_inv is null
	FOR UPDATE 
	DO
		set v_values =
			 convert(varchar(20), glbId)
			+ ', '''''+ p_name + ''''''
			+ ', 1'   -- is_group
		;	
    
		call insert_host (v_table_name, v_fields, v_values);
		UPDATE sguideklass set id_inv = glbId WHERE CURRENT OF a ;
		set glbId = glbId + 1;
	end for;
		

---------------- Добавляем группы по номенклатуре -----------	
	for aCursor as a1 dynamic scroll cursor for
		select 
			i.id_Inv  as i_id
			, p.id_inv as p_id
		from sguideklass i
		join sguideklass p on i.parentklassid = p.klassid
		where i.klassid != 0
	DO
		call update_host (v_table_name, 'belong_id', p_id, 'id = ' +convert(char(15), i_id));
	end for;


---------------- Добавляем группы по изделиям -----------	
	for aCursor as b dynamic scroll cursor for
		select 
			  serianame as p_name
		from sguideseries g
        where seriaid > 0
		and id_inv is null
	FOR UPDATE 
	DO
		set v_values =
			 convert(varchar(20), glbId)
			+ ', '''''+ p_name + ''''''
			+ ', 1'   -- is_group
		;	
    
		call insert_host (v_table_name, v_fields, v_values);
		UPDATE sguideseries set id_inv = glbId WHERE CURRENT OF b ;
		set glbId = glbId + 1;
	end for;
		

	
---------------- Добавляем ??? -----------	
	for aCursor as b1 dynamic scroll cursor for
		select 
			i.id_Inv  as i_id
			, p.id_inv as p_id
		from sguideseries i
		join sguideseries p on i.parentseriaid = p.seriaid
		where i.seriaid != 0
	DO
		call update_host (v_table_name, 'belong_id', p_id, 'id = ' +convert(char(15), i_id));
	end for;
	

---------------- Добавляем номенклатуру -----------	
	set v_fields ='id'
	+ ', belong_id'
	+ ', nomen'
	+ ', nm'
	+ ', id_edizm1'
	+ ', id_edizm2'
	+ ', id_size'
	+ ', prc1'
	+ ', id_curr'

	;
	for aCursor as c_nom dynamic scroll cursor for
		select 
			nomnom as v_nomnom
			, nomname
			, cod as r_cod
			, e1.id_edizm as id_edizm1
			, e2.id_edizm as id_edizm2
			, s.id_size as c_id_size
			, n.size as r_size
			, n.cost as prc1
			, p.id_inv as p_id_inv
		from sguidenomenk n
			join sguideklass p on p.klassid = n.klassid
			left join edizm e1 on n.ed_izmer = e1.name
			left join edizm e2 on n.ed_izmer2 = e2.name
			left join size s on s.name = n.size
		where n.id_inv is null
	FOR UPDATE 
	DO
		if c_id_size is not null  then 
			set v_id_size = convert(varchar(32), c_id_size);
		else 
			set v_id_size = 'null' ;
		end if;
	
		set v_nm = '''''' + wf_make_invnm (nomname, r_size, r_cod) + '''''';

		set v_values =
			 convert(varchar(20), glbId)
			+ ', ' + convert(varchar(20), p_id_inv)
			+ ', ''''' + v_nomnom + ''''''
			+ ', ' + v_nm
			+ ', ' + convert(varchar(20), id_edizm1)
			+ ', ' + convert(varchar(20), id_edizm2)
			+ ', ' + convert(varchar(20), v_id_size)
			+ ', ' + convert(varchar(50), prc1)
			+ ', ' + convert(varchar(50), v_id_currency)
		;	
    
		call insert_host (v_table_name, v_fields, v_values);
		UPDATE sguidenomenk set id_inv = glbId WHERE nomnom = v_nomnom;
		set glbId = glbId + 1;
	end for;
	



---------------- Добавляем изделия -----------	
	set v_fields ='id'
	+ ', belong_id'
	+ ', nomen'
	+ ', nm'
	+ ', id_edizm1'
	+ ', id_size'
	+ ', prc1'
	+ ', is_compl'
	+ ', is_group'
	+ ', id_curr'
	;
	
	select id_edizm into v_id_edizm1 from edizm where name = 'шт.';

	for aCursor as c_izd dynamic scroll cursor for
		select 
			  prName as r_nomNom
			, prDescript as nomName
			, n.prsize as r_size
			, n.cena4 as prc1
			, n.prseriaid as p_prseriaid
		from sguideproducts n
--		join sguideseries p on p.seriaid = n.prseriaid
--		join edizm e1 on e1.name = 'шт.'
--		left join size s on s.name = n.prsize
		where n.id_inv is null
	FOR UPDATE 
	DO

	
		select isnull(id_size, 0) into v_id_size from size where name = r_size;
		if v_id_size is null then
			set v_id_size = 0;
		end if;
		select id_inv into v_id_inv from sguideseries where seriaid = p_prseriaid;
		
		set v_nm = '''''' + wf_make_invnm (nomname, r_size, r_nomnom) + '''''';

		set v_values =
			 convert(varchar(20), glbId)
			+ ', ' + convert(varchar(20), v_id_inv)
			+ ', ''''' + r_nomNom + ''''''
			+ ', ' + v_nm
			+ ', ' + convert(varchar(20), v_id_edizm1)
			+ ', ' + convert(varchar(20), v_id_size)
			+ ', ' + convert(varchar(50), prc1)
			+ ', 1' -- is_compl
			+ ', 0' -- is_group
			+ ', ' + convert(varchar(50), v_id_currency)
		;

		call insert_host (v_table_name, v_fields, v_values);
		UPDATE sguideproducts set id_inv = glbId WHERE CURRENT OF c_izd;
		set glbId = glbId + 1;
	end for;
	
	//Грузим комплектацию к простым (неваринтным) изделиям
	set v_table_name = 'compl';
	set glbId = get_nextId (v_table_name);
	set v_fields ='id'
		+ ', id_inv'
		+ ', id_inv_belong'
		+ ', id_edizm'
		+ ', kol'
		;

	-- Грузить комплектацию и для вариантных изделий, как для простых
	for aCursor as c_compl dynamic scroll cursor for
	select 
		gp.prid as c_productid
		, p.nomNom as c_nomNom
		, gp.id_inv as c_belong_id
		, gn.id_inv as c_id_inv
		, e.id_edizm as c_id_edizm
		, p.quantity as c_kol
	from sguideproducts gp
		join sproducts p on p.productid = gp.prid --and p.xgroup = ''
		join sguidenomenk gn on gn.nomNom = p.nomnom
		left join edizm e on e.name = gn.ed_izmer
	where 
--		not exists (select 1 from svariantpower vp where vp.productid = gp.prid) and 
		p.id_compl is null
	do
		set v_values =
			 convert(varchar(20), glbId)
			+ ', ' + convert(varchar(20), c_id_inv)
			+ ', ' + convert(varchar(20), c_belong_id)
			+ ', ' + convert(varchar(20), c_id_edizm)
			+ ', ' + convert(varchar(50), c_kol)
		;	

		call insert_host (v_table_name, v_fields, v_values);
		update sproducts set id_compl = glbId where productid = c_productid and nomnom = c_nomnom;
		set glbId = glbId + 1;
			
	end for;

	message 'legacy_inv() ended successful.' to client;
end;



if exists (select 1 from sysprocedure where proc_name = 'host_legacy_variant') then
	drop procedure host_legacy_variant;
end if;

CREATE PROCEDURE host_legacy_variant()
begin

	declare v_fix integer;
	declare v_vari integer;
	declare f_check integer;

	message 'host_legacy_variant() started ...' to client;
	select count(*) into f_check from sGuideVariant;
	if f_check = 0 then
		insert into sGuideVariant
			select count(*), productid, xgroup
		from sproducts 
			where ascii(xgroup) != 0 -- is not null or xgroup != '' or xgroup != ' '
		group by productid, xgroup having count(*) > 1;

		for c_var as v dynamic scroll cursor for
			select distinct productid as r_productid from sguidevariant 
		do
			set v_fix = 0;
			set v_vari = 0;
	    
				for c_grp as grp dynamic scroll cursor for
				select count(*) as r_count, xgroup as r_xgroup from sproducts where productid = r_productid group by xgroup
			do
				if r_count = 1 or ascii(r_xgroup) = 0 then
					set v_fix = v_fix + r_count;
				else
					set v_vari = v_vari + 1;
				end if;
			end for;
			
			insert into sVariantPower (numgroup, productid, fixgroups) values(
				v_vari, r_productid, v_fix
			);
		end for;
	end if;
	message 'host_legacy_variant() ended successful.' to client;
end;



if exists (select 1 from sysprocedure where proc_name = 'move_old_voc_names') then
	drop procedure move_old_voc_names ;
end if;

create 
PROCEDURE move_old_voc_names (p_belong_id integer) 
begin
	declare v_old_folder_name varchar(30);

	declare v_old_folder_id integer;
	declare v_belong_name varchar(255);


	-- Находим Id папки - фирм-контрагентов
	select nm into v_belong_name from voc_names_stime where id = p_belong_id;
	
	-- имя папки, куда будем переносить унаследованные элементы
	set v_old_folder_name = v_belong_name + ' (old)';
	
	if exists (select 1 from voc_names_stime where belong_Id = p_belong_id and nm = v_old_folder_name) then
		return;
	end if;
	
--	call slave_select_stime ('voc_names', id_folder_name, 'id', 'nm = ''' + folder_name + '''');


	-- Получить общий id, который будет иметь папка для унаследованных элементов.
	set v_old_folder_id = get_nextid ('voc_names');

	-- создать папку "OLD"
	-- чтобы избежать зацикливания сначала папка добавляетя в корень, (belong_Id = 0)
	call insert_host('voc_names', 'id, nm, belong_id, is_group', 
		convert(varchar(20), v_old_folder_id)
		+ ', '''''+ v_old_folder_name + ''''''
		+ ', 0'
		+ ', 1' 
	);

	-- переносим все унаследованные элементы в папку OLD
	call update_host('voc_names', 'belong_id', convert(varchar(20), v_old_folder_id), 
		'belong_id = ' + convert(varchar(20), p_belong_id)
	-- + ' and nm != '''''+ old_nomen +'''''' 
	);
	
	-- потом папка OLD перепривязывается к корневой комтеховской папке 
	call update_host('voc_names', 'belong_id', convert(varchar(20), p_belong_id), 
		'id = ' + convert(varchar(20), v_old_folder_id)
	);


end;


if exists (select 1 from sysprocedure where proc_name = 'move_old_inv') then
	drop procedure move_old_inv;
end if;

create 
PROCEDURE move_old_inv (p_belong_id integer) 
begin
	declare v_old_folder_name varchar(30);

	declare v_old_folder_id integer;
	declare v_belong_name varchar(255);

	-- Находим Id папки
	select nm into v_belong_name from inv_stime where id = p_belong_id;
	
	-- имя папки, куда будем переносить унаследованные элементы
	set v_old_folder_name = v_belong_name + ' (old)';
	
	if exists (select 1 from inv_stime where belong_Id = p_belong_id and nm = v_old_folder_name) then
		return;
	end if;
	

	-- Получить общий id, который будет иметь папка для унаследованных элементов.
	set v_old_folder_id = get_nextid ('inv');

	-- создать папку "OLD"
	-- чтобы избежать зацикливания сначала папка добавляетя в корень, (belong_Id = 0)
	call insert_host('inv', 'id, nm, belong_id, is_group', 
		convert(varchar(20), v_old_folder_id)
		+ ', '''''+ v_old_folder_name + ''''''
		+ ', 0'
		+ ', 1' 
	);

	-- переносим все унаследованные элементы в папку OLD
	call update_host('inv', 'belong_id', convert(varchar(20), v_old_folder_id), 
		'belong_id = ' + convert(varchar(20), p_belong_id)
	-- + ' and nm != '''''+ old_nomen +'''''' 
	);
	
	-- потом папка OLD перепривязывается к корневой комтеховской папке 
	call update_host('inv', 'belong_id', convert(varchar(20), p_belong_id), 
		'id = ' + convert(varchar(20), v_old_folder_id)
	);


end;


if exists (select 1 from sysprocedure where proc_name = 'legacy_currency') then
	drop procedure legacy_currency;
end if;

create
	PROCEDURE legacy_currency () 
begin
	declare v_old_folder_name varchar(30);
	declare v_fields varchar(1000);
	declare v_values varchar(2000);
	declare v_where varchar(1000);
	declare v_id_cur_rate integer;
	declare v_id_cur integer;
	declare v_currency_rate float;

	message 'legacy_currency() started ...' to client;

	select id_cur into v_id_cur from system;
	if v_id_cur is not null then
		return;
	end if;


	set v_id_cur = get_nextid ('currency');
	

	set v_fields = 'id, nm, base_1, base_2, base_0, sub_1, sub_2, sub_0, rem, iso_code';
	set v_values = 
		convert(varchar(20), v_id_cur)
		 +', ''''УЕ'''''
		 +', ''''условная единица'''''
		 +', ''''условные единицы'''''
		 +', ''''условных единиц'''''
		 +', ''''цент УЕ'''''
		 +', ''''цента УЕ'''''
		 +', ''''центов УЕ'''''
		 +', ''''Внутренний курс компании'''''
		 +', ''''UE'''''
	;

	call insert_host('currency'
		, v_fields
		, v_values
	);
	update system set id_cur = v_id_cur;

	set v_currency_rate = system_currency_rate();
	set v_id_cur_rate = get_nextid('cur_rate');
	set v_fields = 'id, id_cur, dat, curse, rem';
	set v_values = 
		convert(varchar(20), v_id_cur_rate)
		+', ''''' + 		convert(varchar(20), v_id_cur) + ''''''
		+', ''''' + convert(varchar(20), '2005-03-17', 112) +''''''
		+', ''''' + convert(varchar(20), v_currency_rate) + ''''''
		+', ''''Установлено в Prior'''''
	;
	
--	set v_where = 'id_cur = ' + convert(varchar(20), v_id) + 'and dat ' + convert(varchar(20), now, 112);

	call insert_host('cur_rate', v_fields, v_values);
	update system set id_cur_rate = v_id_cur_rate;

	message 'legacy_currency() ended successful.' to client;
end;



if exists (select 1 from sysprocedure where proc_name = 'legacy_income_order') then
	drop procedure legacy_income_order;
end if;

create
	PROCEDURE legacy_income_order () 
begin

	declare v_id_inventar integer;
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_fields varchar(100);
	declare v_values varchar(500);
	declare v_nu varchar(20);
	declare v_mat_nu integer;
	declare v_quant float;
	declare v_cost float;
	declare v_currency_rate real;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_legacy varchar(100);
	declare v_gemacht datetime;
	declare v_perList float;

	declare sync char(1);

	message 'legacy_income_order() started ...' to client;

	set v_legacy = 'Переход на режим совместного использования Prior/Komtex (сгенерировано)';
	select trans_date into v_gemacht from system;

--	call slave_select_stime(v_gemacht, 'jmat', 'count(*)', 'osn = ''' + v_legacy + '''');
	if v_gemacht is not null then
		message 'Входящие остатки уже загружены' to client;
		return;
	end if;

	create table #saldo(nomnom varchar(20), id integer, debet float, kredit float);

	create table #itogo(nomnom varchar(20), id integer, debet float, kredit float);

	insert into #saldo (nomnom, id, debet)
    select nomnom, destid, sum(quant) from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
    --where sourId = -1001 or destid = -1001
	group by nomnom, destid;
    
	insert into #saldo (nomnom, id, kredit)
    select nomnom, sourid, sum(quant) from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
    --where sourId = -1001 or destid = -1001
	group by nomnom, sourid;
    

	insert into #itogo (nomnom, id, debet, kredit)
    select nomnom, id, sum(isnull(debet,0)), sum(isnull(kredit,0))
	from #saldo 
    group by nomnom, id;

--    begin
--		call call_host('block_table', 'sync, ''prior'', ''jmat''');
--		call call_host('block_table', 'sync, ''prior'', ''mat''');
		call block_remote('stime', get_server_name(), 'jmat');
		call block_remote('stime', get_server_name(), 'mat');

        -- глобальный для загловков накладных
		set v_id_jmat = get_nextid('jmat');

        -- глобальный для загловков накладных
		set v_id_mat = get_nextid('mat');
--		set v_currency_rate = system_currency_rate();
		set v_id_currency = system_currency();
		call slave_currency_rate_stime(v_datev, v_currency_rate);


   	   	for sklad_cur as s dynamic scroll cursor for
			select sourceid as r_sourceid, id_voc_names as r_id_sklad 
			from sguidesource
			where sourceid <= -1001
		do
			call slave_select_stime(v_nu, 'jmat', 'max(nu)', '1=1');
			set v_nu = convert(varchar(20), convert(integer, isnull(v_nu, 0)) + 1);
			select id_voc_names into v_id_inventar from sguidesource where sourceName = 'Инвентаризация';


			call wf_insert_jmat (
				'stime'
				,'1023' --инветаризация
				,v_id_jmat
				,now() --v_jmat_date
				,v_nu
				,v_legacy --v_osn
				,v_id_currency
				,v_datev
				,v_currency_rate
				,v_id_inventar
				,r_id_sklad
			);

        	-- Добавляем предметы к накладной
        	set v_mat_nu = 1;
			for nom_cur as n dynamic scroll cursor for
				select i.nomnom as r_nomnom, n.id_inv as r_nomenklature_id, debet as r_debet, kredit as r_kredit 
				from #itogo i
				join sguidenomenk n on n.nomnom = i.nomnom
	            where id = r_sourceid
			do
				set v_quant = r_debet - r_kredit;

				if v_quant >= 0.01 then

					select cost, perList into v_cost, v_perList from sguidenomenk where nomnom = r_nomnom;

					call wf_insert_mat (
						'stime'
						,v_id_mat
						,v_Id_jmat
						,r_nomenklature_id
						,v_mat_nu
						,v_quant
						,v_cost
						,v_currency_rate
						,v_id_inventar
						,r_id_sklad
						,v_perList
					);

					set v_id_mat = v_id_mat + 1;
					set v_mat_nu = v_mat_nu + 1;
				end if;

			end for;
			set v_id_jmat = v_id_jmat + 1;
		end for;


		call unblock_remote('stime', get_server_name(), 'jmat');
		call unblock_remote('stime', get_server_name(), 'mat');
--		call call_host('unblock_table', 'sync, ''prior'', ''jmat''');
--		call call_host('unblock_table', 'sync, ''prior'', ''mat''');
--	exception 
--		when others then
--			set v_perList = v_perList;
--	end;

	drop table #saldo;
	drop table #itogo;
    
	message 'legacy_income_order() ended successful.' to client;
end;


if exists (select 1 from sysprocedure where proc_name = 'legacy_purpose') then
	drop procedure legacy_purpose;
end if;

create
	PROCEDURE legacy_purpose () 
begin
	declare v_param varchar(512);
	declare i_subSchet integer;

	message 'legacy_purpose() started ...' to client;
		
	for nom_cur as n dynamic scroll cursor for
		select pDescript as r_nm, debit, subdebit, kredit, subkredit
		from yguidePurpose
	do
		set i_subSchet = convert(integer, subdebit);
		if i_subSchet = 0 then
			set subdebit = '';
		else
			set subdebit = convert(varchar(10), i_subSchet);
		end if;

		set i_subSchet = convert(integer, subkredit);
		if i_subSchet = 0 then
			set subkredit = '';
		else
			set subkredit = convert(varchar(10), i_subSchet);
		end if;

		set v_param = 
			'''' + r_nm + ''''
			+ ', '''+debit + ''''
			+ ', '''+subdebit + ''''
			+ ', '''+kredit + ''''
			+ ', '''+subkredit + ''''
		;
		call call_host('legacy_purpose', v_param);
	end for;

	message 'legacy_purpose() ended successful.' to client;
end;