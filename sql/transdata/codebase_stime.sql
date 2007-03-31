

if exists (select '*' from sysprocedure where proc_name like 'wf_calc_cost') then  
	drop procedure wf_calc_cost;
end if;


create procedure wf_calc_cost (
	  out out_ret float
	, p_id_inv integer
) 
begin

	--execute immediate 'create variable @adec_Ost21 decimal(19, 7)';

	--call calc_ost_inv(now(), p_id_inv, -1, -2,  '1' , '2' , '1' , 0 , '0' , '0' , 1 , 1 , '0' , '0' , '0' , 0 );
	set out_ret = calc_summa('mat', -1, now(), p_id_inv, -2, 'summa', 1, 7);
	--message summa_rub, ' ', @adec_Ost21 to client;

--	set v_string_prc = select_remote('stime', 'inv', 'prc1', 'id = ' + convert(varchar(20), p_id_inv));
	--set out_ret = summa_rub / @adec_Ost21;
	--execute immediate 'drop variable @adec_Ost21';

end;


if exists (select '*' from sysprocedure where proc_name like 'wf_cost_date') then  
	drop procedure wf_cost_date;
end if;


create procedure wf_cost_date (
	  out out_ret float
	, p_id_inv integer
	, p_date date
) 
begin

	set out_ret = calc_summa('mat', -1, p_date, p_id_inv, -2, 'summa', 1, 7);

end;




if exists (select 1 from systriggers where trigname = 'wf_analytic_income' and tname = 'jmat') then 
	drop trigger jmat.wf_analytic_income;
end if;

/*
create TRIGGER "wf_analytic_income" before update order 101 on
jmat
referencing old as old_name new as new_name
for each row
when
	(update(id_code))
begin
	declare v_numdoc varchar(20);
	declare v_numext varchar(20);
	declare v_id_analytic_default integer;
	declare v_code_name varchar(30);

		set v_numdoc = admin.select_remote('prior', 'sdocs', 'numDoc', 'id_jmat = ' + convert(varchar(20), old_name.id));
		set v_numext = admin.select_remote('prior', 'sdocs', 'numExt', 'id_jmat = ' + convert(varchar(20), old_name.id));

		-- Приходуем накладную на то или иное предприятие в зависимости от 
		-- кода аналитики
		-- для этого в таблицу sDocsIncome добавляем/удаляем строку со
		-- ссылкой на предприятие

		set v_id_analytic_default = convert(integer, admin.select_remote('prior', 'system', 'id_analytic_default', '1=1'));

		if isnull(v_numdoc, 0) != 0 and isnull(v_numExt, 0) != 0 then
			call admin.delete_remote (
				  'prior'
				, 'sDocsIncome'
				, 'numdoc = ' + v_numdoc + ' and numext = ' + v_numext
			);
		end if;


		if new_name.id_code != v_id_analytic_default and isnull(new_name.id_code, 0) != 0 then
			call admin.insert_remote(
				  'prior'
				, ' sDocsIncome'
				, ' ventureid, numdoc, numext, id_analytic'
				, null
				, ' select v.ventureId'
					+ '		, ' + v_numdoc
					+ '		, ' + v_numext
					+ '		, ' + convert(varchar(20), new_name.id_code)
					+ ' from GuideVenture v'
					+ ' where '
					+ '		v.id_analytic = ' + convert(varchar(20), new_name.id_code)
			);

			select code into v_code_name from analytic where id = new_name.id_code;

			if v_code_name is not null and char_length(v_code_name) > 0 then
				call admin.update_remote(
					'prior'
					, 'sDocs'
					, 'note'
					, ''''''+ v_code_name +''''''
					, 'numdoc = ' + v_numdoc + ' and numext = ' + v_numext
				);
			end if;

		end if;
end;
*/

--
-- Триггера и функции которые обеспечивают коммуникацию
-- от комтеховской базы sTime к базе Приора
--	Описание Функциональности:
--		1. При заведении приходных накладных средствами Комтеха
-- добавляются нужные данные в базу приора.
-- Эти триггера отключены из-за проблем с блокировкой

if exists (select 1 from systriggers where trigname = 'wf_income' and tname = 'jmat') then 
	drop trigger jmat.wf_income;
end if;

/*
create TRIGGER "wf_income" before update order 100 on
jmat
referencing old as old_name new as new_name
for each row
when(new_name.id_guide = 1120 or old_name.id_guide = 1120)
begin
	declare v_values varchar(2000);
	declare v_fields varchar(256);
	declare v_numdoc varchar(20);
	declare v_numext varchar(20);
	declare v_where varchar(1000);
	declare cur_id integer;
	declare no_echo integer;
	declare v_id_analytic_default integer;
	declare v_code_name varchar(30);
	
	set no_echo = 0;

  	begin
  		message '@prior_jmat = ', @prior_jmat to log;
		select @prior_jmat into no_echo; 
	exception 
		when other then
			message 'exception at @prior_jmat' to log;
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

	message 'Server STIME, Trigger jmat.wf_income, no-echo =', no_echo to log;
	call admin.block_remote('prior', @@servername, 'sdocs');

begin
	if update(id_code) then
		
		set v_numdoc = admin.select_remote('prior', 'sdocs', 'numDoc', 'id_jmat = ' + convert(varchar(20), old_name.id));
		set v_numext = admin.select_remote('prior', 'sdocs', 'numExt', 'id_jmat = ' + convert(varchar(20), old_name.id));

		-- Приходуем накладную на то или иное предприятие в зависимости от 
		-- кода аналитики
		-- для этого в таблицу sDocsIncome добавляем/удаляем строку со
		-- ссылкой на предприятие

		set v_id_analytic_default = convert(integer, admin.select_remote('prior', 'system', 'id_analytic_default', '1=1'));

		if isnull(v_numdoc, 0) != 0 and isnull(v_numExt, 0) != 0 then
			call admin.delete_remote (
				  'prior'
				, 'sDocsIncome'
				, 'numdoc = ' + v_numdoc + ' and numext = ' + v_numext
			);
		end if;


		if new_name.id_code != v_id_analytic_default and isnull(new_name.id_code, 0) != 0 then
			call admin.insert_remote(
				  'prior'
				, ' sDocsIncome'
				, ' ventureid, numdoc, numext, id_analytic'
				, null
				, ' select v.ventureId'
					+ '		, ' + v_numdoc
					+ '		, ' + v_numext
					+ '		, ' + convert(varchar(20), new_name.id_code)
					+ ' from GuideVenture v'
					+ ' where '
					+ '		v.id_analytic = ' + convert(varchar(20), new_name.id_code)
			);

			select code into v_code_name from analytic where id = new_name.id_code;

			if v_code_name is not null and char_length(v_code_name) > 0 then
				call admin.update_remote(
					'prior'
					, 'sDocs'
					, 'note'
					, ''''''+ v_code_name +''''''
					, 'numdoc = ' + v_numdoc + ' and numext = ' + v_numext
				);
			end if;

		end if;

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
		call admin.slave_insert_prior('sdocs',v_fields,v_values);
		call admin.slave_select_prior(v_numdoc, 'sdocs', 'numdoc', 'id_jmat = ' + convert(varchar(10), new_name.id));
		set new_name.nu = v_numdoc;
	end if;

	if update(id_s) then
		call admin.slave_select_prior(cur_id,'sguidesource','sourceId','id_voc_names = '+convert(varchar(20),new_name.id_s));
		if cur_Id is not null then
		    set v_fields='sourId';
		    set v_values = convert(varchar(25),cur_id);
			set v_where = 'id_jmat = ' + convert(varchar(20), old_name.id);
			call admin.slave_update_prior('sdocs',v_fields,v_values,v_where);
		end if;
	end if;

	if update(id_d) then
		call admin.slave_select_prior (cur_id, 'sguidesource', 'sourceId', 'id_voc_names = ' + convert(varchar(20), new_name.id_d));
		if cur_Id is not null then
			set v_fields='destId';
			set v_values = convert(varchar(25),cur_id);
			set v_where = 'id_jmat = ' + convert(varchar(20), old_name.id);

			call admin.slave_update_prior('sdocs',v_fields,v_values,v_where);

		end if;
	end if;

exception when others then
	call admin.unblock_remote('prior', @@servername, 'sdocs');
end;

	call admin.unblock_remote('prior', @@servername, 'sdocs');

end;
*/

if exists (select '*' from sysprocedure where proc_name like 'put_sdmc') then
	drop procedure put_sdmc;
end if;

create procedure put_sdmc (
	  in p_id_jmat integer
	, in p_id_mat  integer
	, in p_id_inv  integer
) 
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);

	declare v_guide_id integer;
    declare v_nomNom varchar(50);
  	declare v_numDoc varchar(50);
	declare v_numExt integer;

	declare no_echo integer;
	set no_echo = 0;

  	begin
  		message '@prior_mat = ', @prior_mat to log;
		select @prior_mat into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;




	select id_guide, nu into v_guide_id, v_numDoc from jmat where id = p_id_jmat;


	if v_guide_Id = 1120 then
		set v_numExt = 255;

		set v_nomnom = admin.select_remote('prior', 'sguidenomenk', 'nomnom', 'id_inv = ' + convert(varchar(20), p_id_inv));


		set v_fields =
			  'id_mat'
			+ ', numDoc'
			+ ', numExt'
			+ ', nomNom'
		;
		set v_values=
			  convert(varchar(20),p_id_mat)
			+ ', ' + v_numDoc
			+ ', '+ convert(varchar(25),v_numExt)
			+ ', ''' + v_nomNom + ''''
		;
        
        
		
		call admin.slave_insert_prior('sdmc',v_fields,v_values);

	end if;

end;



-- При добавлении предметов к накладной в Комехе добавляются 
-- строки и в приоре.
if exists (select 1 from systriggers where trigname = 'wf_income_detail' and tname = 'mat' and event='INSERT') then 
	drop trigger mat.wf_income_detail;
end if;

/*
create TRIGGER "wf_income_detail" before insert order 100 on
mat
referencing new as new_name
for each row
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

	set no_echo = 0;

	begin
  		message '@prior_mat = ', @prior_mat to log;
		select @prior_mat into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

	set v_id_mat = new_name.id;
	set v_id_jmat = new_name.id_jmat;

--	raiserror 17001 'v_id_mat = %1!, v_id_jmat = %2!', v_id_mat, v_id_jmat;

    if v_id_jmat is not null and v_id_jmat > 0 then


	   	call admin.put_sdmc (v_id_jmat, v_id_mat, new_name.id_inv);
    
	end if;

end;
*/

-- Редактирование количества прихода а также изменение вызывает 
-- адекватное изменение и в базе Приора
if exists (select 1 from systriggers where trigname = 'wf_income_detail_upd' and tname = 'mat' and event='UPDATE') then 
	drop trigger mat.wf_income_detail_upd;
end if;

/*
create TRIGGER wf_income_detail_upd before update order 100 on
mat
referencing old as old_name new as new_name
for each row
begin
	declare v_nomNom varchar(50);
	declare v_numDoc varchar(20);
	declare no_echo integer;
	declare v_perList varchar(20);
	
	set no_echo = 0;

  	begin
  		message '@prior_mat = ', @prior_mat to log;
		select @prior_mat into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;


	call admin.block_remote('prior', @@servername, 'sdmc');


	if update(kol1) then
--		set v_perList = 1;
		set v_perList = admin.select_remote('prior', 'sguidenomenk', 'perList', 'id_inv = '+ convert(varchar(20), old_name.id_inv));
		if v_perList is not null then
			call admin.update_remote('prior', 'sdmc','quant',convert(varchar(20),round(new_name.kol1*convert(float, v_perList), 2)),'id_mat = '+convert(varchar(20),old_name.id) );
		end if;
	end if;
	if update(id_inv) then
--		select nomen into v_nomNom from inv where id = new_name.id_inv;
		set v_nomnom = admin.select_remote('prior', 'sguidenomenk', 'nomnom', 'id_inv = ' + convert(varchar(20), new_name.id_inv));
		call admin.update_remote('prior', 'sdmc','nomnom',''''+v_nomNom + '''','id_mat = '+convert(varchar(20),old_name.id) );
	end if;
	if update(id_jmat) then
		if old_name.id_jmat = 0 then
			-- в документах приора не было еще этой накладной
		   	call admin.put_sdmc (new_name.id_jmat, old_name.id, old_name.id_inv);
		else
			-- исправляем у предмета накладной номер накладной
			select nu into v_numDoc from jmat where id = new_name.id_jmat;
--			set v_id_jmat = admin.select_remote('prior', 'sdocs', 'id_jmat', 'numdoc = ' + v_numdoc);
			call admin.update_remote('prior', 'sdmc','numDoc',''''+v_numDoc + '''','id_mat = '+convert(varchar(20),old_name.id) );
		end if;

	end if;

	call admin.unblock_remote('prior', @@servername, 'sdmc');

end;
*/

-- Каскадное удаление в Приоре из Комтеха 
-- для предметов ...
if exists (select 1 from systriggers where trigname = 'wf_income_detail_del' and tname = 'mat') then 
	drop trigger mat.wf_income_detail_del;
end if;


/*
create TRIGGER "wf_income_detail_del" before delete order 100 on
mat
referencing old as old_name
for each row
begin
	declare no_echo integer;
	set no_echo = 0;

  	begin
  		message '@prior_mat = ', @prior_mat to log;
		set no_echo = @prior_mat;
	exception 
		when other then
			--message 'Exception! no_echo = ' + convert(varchar(20), no_echo) to log;
			set no_echo = 0;
	end;

	--message 'trigger wf_income_detail_del::no_echo = ' + convert(varchar(20), no_echo) to log;

	if no_echo = 1 then
		return;
	end if;

	call admin.block_remote('prior', @@servername, 'sdmc');
	call admin.slave_delete_prior('sdmc','id_mat = '+convert(varchar(20),old_name.id) );
	call admin.unblock_remote('prior', @@servername, 'sdmc');

end;
*/

-- .. и для накладной
if exists (select 1 from systriggers where trigname = 'wf_income_del' and tname = 'jmat') then 
	drop trigger jmat.wf_income_del;
end if;

/*
create TRIGGER "wf_income_del" before delete order 100 on
jmat
referencing old as old_name
for each row
begin
	declare no_echo integer;
	set no_echo = 0;

  	begin
  		message '@prior_jmat = ', @prior_jmat to log;
		select @prior_jmat into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

	call admin.block_remote('prior', @@servername, 'sdocs');
	call admin.block_remote('prior', @@servername, 'sdmc');

    call admin.slave_delete_prior('sdocs','id_jmat = '+convert(varchar(20),old_name.id) );

	call admin.unblock_remote('prior', @@servername, 'sdmc');
	call admin.unblock_remote('prior', @@servername, 'sdocs');
end;
*/



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



