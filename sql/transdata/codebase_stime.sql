--
-- Триггера и функции которые обеспечивают коммуникацию
-- от комтеховской базы st к базе Приора
--	Описание Функциональности:
--		1. При заведении приходных накладных средствами Комтеха
-- добавляются нужные данные в базу приора.

-- Загрузка:
if exists (select 1 from systriggers where trigname = 'wf_income' and tname = 'jmat') then 
	drop trigger jmat.wf_income;
end if;

create TRIGGER "wf_income" before update order 100 on
jmat
referencing old as old_name new as new_name
for each row
when(new_name.id_guide = 1120 or old_name.id_guide = 1120)
begin
	declare v_values varchar(2000);
	declare v_fields varchar(256);
	declare v_numdoc varchar(20);
	declare v_where varchar(1000);
	declare cur_id integer;
	declare no_echo integer;
	
	set no_echo = 0;

  	begin
		select @prr_jmat into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

	if update(id_guide) then
		set v_fields =
			'id_jmat'
		+ ', numExt'
		+ ', xDate'
		+ ', Note'
		;
		set v_values=
			convert(varchar(20),new_name.id)
		+', 255'
		+ ', '''+ convert(varchar(25),new_name.dat) + ''''
		+ ', ''оприходовано через Komtex'''
		;
		call admin.slave_insert_prr('sdocs',v_fields,v_values);
		call admin.slave_select_prr(v_numdoc, 'sdocs', 'numdoc', 'id_jmat = ' + convert(varchar(10), new_name.id));
		set new_name.nu = v_numdoc;
	end if;

	if update(id_s) then
		call admin.slave_select_prr(cur_id,'sguidesource','sourceId','id_voc_names = '+convert(varchar(20),new_name.id_s));
		if cur_Id is not null then
		    set v_fields='sourId';
		    set v_values = convert(varchar(25),cur_id);
				set v_where = 'id_jmat = ' + convert(varchar(20), old_name.id);
			call admin.slave_update_prr('sdocs',v_fields,v_values,v_where);
		end if;
	end if;
	if update(id_d) then
		call admin.slave_select_prr (cur_id, 'sguidesource', 'sourceId', 'id_voc_names = ' + convert(varchar(20), new_name.id_d));
		if cur_Id is not null then
			set v_fields='destId';
			set v_values
			=convert(varchar(25),cur_id);
				set v_where = 'id_jmat = ' + convert(varchar(20), old_name.id);
			call admin.slave_update_prr('sdocs',v_fields,v_values,v_where);
		end if;
	end if
end;


-- При добавлении предметов к накладной в Комехе добавляются 
-- строки и в приоре.
if exists (select 1 from systriggers where trigname = 'wf_income_detail' and tname = 'mat' and event='INSERT') then 
	drop trigger mat.wf_income_detail;
end if;

create TRIGGER "wf_income_detail" before insert order 100 on
mat
referencing new as new_name
for each row
--when(new_name.id_guide = 1120 or old_name.id_guide = 1120)
begin
	declare v_values varchar(2000);
	declare v_fields varchar(256);
	declare v_id_jmat integer;
    declare v_id_mat integer;
	declare v_guide_id integer;
    declare v_nomNom varchar(50);
  	declare v_numDoc varchar(50);
	declare v_numExt integer;
	declare no_echo integer;

	set v_id_mat = new_name.id;
	set v_id_jmat = new_name.id_jmat;
	select id_guide, nu into v_guide_id, v_numDoc from jmat where id = v_id_jmat;

	set no_echo = 0;

  	begin
		select @prr_mat into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if v_guide_Id = 1120 and no_echo = 0 then
		set v_numExt = 255;
		set v_fields =
			  'id_mat'
			+ ', numDoc'
			+ ', numExt'
			+ ', nomNom'
		;

		select nomen into v_nomnom from inv where id = new_name.id_inv;
		set v_values=
			  convert(varchar(20),v_id_mat)
			+ ', ''' + v_numDoc + ''''
			+ ', '+ convert(varchar(25),v_numExt)
			+ ', ''' + v_nomNom + ''''
		;
	
		--message 'new_name.id_jmat', new_name.id_jmat to client;
		call admin.slave_insert_prr('sdmc',v_fields,v_values);
		--set new_name.nu = v_numdoc;
	end if;
end;


-- Редактирование количества прихода а также изменение вызывает 
-- адекватное изменение и в базе Приора
if exists (select 1 from systriggers where trigname = 'wf_income_detail_upd' and tname = 'mat' and event='UPDATE') then 
	drop trigger mat.wf_income_detail_upd;
end if;

create TRIGGER wf_income_detail_upd before update order 100 on
mat
referencing old as old_name new as new_name
for each row
begin
	declare v_nomNom varchar(50);
	declare v_numDoc varchar(20);
	declare no_echo integer;
	
	set no_echo = 0;

  	begin
		select @prr_mat into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

	if update(kol1) then
		call admin.slave_update_prr('sdmc','quant',convert(varchar(20),new_name.kol1),'id_mat = '+convert(varchar(20),old_name.id) )
	end if;
	if update(id_inv) then
		select nomen into v_nomNom from inv where id = new_name.id_inv;
		call admin.slave_update_prr('sdmc','nomnom',''''+v_nomNom + '''','id_mat = '+convert(varchar(20),old_name.id) )
	end if;
	if update(id_jmat) then
		select nu into v_numDoc from jmat where id = new_name.id_jmat;
		call admin.slave_update_prr('sdmc','numDoc',''''+v_numDoc + '''','id_mat = '+convert(varchar(20),old_name.id) )
	end if
end;


-- Каскадное удаление в Приоре из Комтеха 
-- для предментов ...
if exists (select 1 from systriggers where trigname = 'wf_income_detail_del' and tname = 'mat') then 
	drop trigger mat.wf_income_detail_del;
end if;

create TRIGGER "wf_income_detail_del" before delete order 100 on
mat
referencing old as old_name
for each row
begin
	declare no_echo integer;
	set no_echo = 0;

	message 'no_echo = ' + convert(varchar(20), no_echo) to log;

  	begin
		select @prr_mat into no_echo; 
	exception 
		when other then
			message 'Exception! no_echo = ' + convert(varchar(20), no_echo) to log;
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

	message 'before delete from prr :: no_echo = ' + convert(varchar(20), no_echo) to client;
    call admin.slave_delete_prr('sdmc','id_mat = '+convert(varchar(20),old_name.id) );
	message 'after delete from prr :: no_echo = ' + convert(varchar(20), no_echo) to client;
end;


-- .. и для накладной
if exists (select 1 from systriggers where trigname = 'wf_income_del' and tname = 'jmat') then 
	drop trigger jmat.wf_income_del;
end if;

create TRIGGER "wf_income_del" before delete order 100 on
jmat
referencing old as old_name
for each row
begin
	declare no_echo integer;
	set no_echo = 0;

  	begin
		select @prr_jmat into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

    call admin.slave_delete_prr('sdocs','id_jmat = '+convert(varchar(20),old_name.id) );
end;




/*
-- Триггера - определители.
-- Для обнаружения, когда, как и для чего 
-- используются глобальные временные таблицы

if exists (select 1 from systriggers where trigname = 'test_insert' and tname = 'inventory') then 
	drop trigger inventory.test_insert;
end if;

create TRIGGER test_insert before insert order 1 on
inventory
REFERENCING NEW AS new_name
for each row 
begin
	raiserror 17000 'insert into inventory';
end;






if exists (select 1 from systriggers where trigname = 'test_insert' and tname = 'inv_ost') then 
	drop trigger inv_ost.test_insert;
end if;

create TRIGGER test_insert before insert order 1 on
inv_ost
REFERENCING NEW AS new_name
for each row 
begin
	raiserror 17000 'insert into inv_ost';
end;





if exists (select 1 from systriggers where trigname = 'test_insert' and tname = 'inv_ost_fact_scet') then 
	drop trigger inv_ost_fact_scet.test_insert;
end if;

create TRIGGER test_insert before insert order 1 on
inv_ost_fact_scet
REFERENCING NEW AS new_name
for each row 
begin
	raiserror 17000 'insert into inv_ost_fact_scet';
end;




if exists (select 1 from systriggers where trigname = 'test_insert' and tname = 'inv_ost_st') then 
	drop trigger inv_ost_st.test_insert;
end if;

create TRIGGER test_insert before insert order 1 on
inv_ost_st
REFERENCING NEW AS new_name
for each row 
begin
	raiserror 17000 'insert into inv_ost_st';
end;




if exists (select 1 from systriggers where trigname = 'test_insert' and tname = 'inventory_protocol') then 
	drop trigger inventory_protocol.test_insert;
end if;

create TRIGGER test_insert before insert order 1 on
inventory_protocol
REFERENCING NEW AS new_name
for each row 
begin
	raiserror 17000 'insert into inventory_protocol';
end;




if exists (select 1 from systriggers where trigname = 'test_insert' and tname = 'inventory_protocol') then 
	drop trigger inventory_ost.test_insert;
end if;

create TRIGGER test_insert before insert order 1 on
inventory_ost
REFERENCING NEW AS new_name
for each row 
begin
	raiserror 17000 'insert into inventory_ost';
end;





if exists (select 1 from systriggers where trigname = 'test_insert' and tname = 'inventory_error') then 
	drop trigger inventory_error.test_insert;
end if;

create TRIGGER test_insert before insert order 1 on
inventory_error
REFERENCING NEW AS new_name
for each row 
begin
	raiserror 17000 'insert into inventory_error';
end;

*/
