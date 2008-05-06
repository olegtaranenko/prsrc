

-------------- удалить после вычистки -------------------

if exists (select 1 from sysprocedure where proc_name = 'cast_acc') then
	drop function cast_acc;
end if;

create 
	function cast_acc(
		in sc varchar(26)
		,in base integer default 2
	)
	returns varchar(26)
begin
	if sc is null then
		set sc = '';
	end if;

	set sc = trim(sc);
	return string(repeat('0', base - char_length(sc)), sc);
end;

--------------> удалить после вычистки <-------------------





/**
 get_server_name() => @server_name 
 процедура должна вызыватьс€ один раз из bootstrap_blocking.
*/                 

if exists (select '*' from sysprocedure where proc_name like 'get_server_name') then  
	drop function get_server_name;
end if;

create function get_server_name ()
returns varchar(20) 
begin
	set get_server_name = @@servername;
	if (substring (get_server_name, 1, 3) = 'dev') then
		select namer into get_server_name from guides where id = 0;
	end if;

end;




if exists (select '*' from sysprocedure where proc_name like 'wf_calc_ost_inv') then  
	drop procedure wf_calc_ost_inv;
end if;


create procedure wf_calc_ost_inv (
	  out out_ret float
	, p_id_inv integer
	, p_id_sklad integer default -2
) 
begin

    	for f_ost as ost dynamic scroll cursor for
            call calc_ost_inv(now(), p_id_Inv, -1, p_id_sklad,  '1' , '2' , '1' , 0 , '0' , '0' , 1 , 1 , '0' , '0' , '0' , 0 )
    	do
    		if p_id_sklad = -2 then
    			set out_ret = adec_ost21;
    		else
	    		set out_ret = adec_ost11;
    		end if;
    	end for;
end;


if exists (select '*' from sysprocedure where proc_name like 'wf_order_closed_set') then  
	drop procedure wf_order_closed_set;
end if;


create procedure wf_order_closed_set (
	  in p_id_jscet integer
	, in p_do_close integer
) 
begin
	declare v_id_jmat integer;
	declare v_numorder varchar(20);
	declare v_tp1_close char(1);
	declare v_tp2_close char(1);
	declare v_tp3_close char(1);
	declare v_tp4_close char(1);

	declare v_tp1_open char(1); 
	declare v_tp2_open char(1); 
	declare v_tp3_open char(1); 
	declare v_tp4_open char(1); 
	
	set v_tp1_close = '2';
	set v_tp2_close = '2';
	set v_tp3_close = '2';
	set v_tp4_close = '0';

	set v_tp1_open = '3';
	set v_tp2_open = '2';
	set v_tp3_open = '1';
	set v_tp4_open = '7';

	--set v_id_jmat = select_remote('prior', 'all_orders', 'id_jmat', 'id_jscet = ' + convert(varchar(20), p_id_jscet));
	if isnull(p_id_jscet,0) != 0 then
		if p_do_close = 1 then
			update jmat set 
				  tp1 = v_tp1_close
				, tp2 = v_tp2_close
				, tp3 = v_tp3_close
				, tp4 = v_tp4_close
			where id_jscet = p_id_jscet;
			update jscet set data_lock = 1 where id = p_id_jscet;
	    
		else
			update jmat set 
				tp1 = v_tp1_open
				, tp2 = v_tp2_open
				, tp3 = v_tp3_open
				, tp4 = v_tp4_open
			where id_jscet = p_id_jscet;
			update jscet set data_lock = 0 where id = p_id_jscet;

		end if;
	end if;
end;
        	                            		

if exists (select 1 from systriggers where trigname = 'wf_jscet_close' and tname = 'jscet' and event='update') then 
	drop trigger jscet.wf_jscet_close;
end if;

begin
	declare v_table_name varchar(128);
	declare v_column_name varchar(128);
	declare v_status_close_id integer;
	declare v_trigger_sql varchar(3000);
	

	-- Ќайти пользовательский справочник и колонку в журнале ордеров
	-- получаем что-то типа этого 'GUIDE_803_129574.NM','JSCET__USER_129573'
	select nm, parent_col_name
	into v_table_name, v_column_name
	from browsers where id_guides = 1005 
	and nm like '%guid%' 
	and namer like '%зак%';

	if v_table_name is null then 
		return;
	end if;
	-- очищаем до  'GUIDE_803','USER_129573'
	set v_table_name = 'GUIDE_' + substring(v_table_name, 7, charindex('_', substring(v_table_name, 7))-1);
	set v_column_name =  substring(v_column_name, charindex('__', v_column_name)+2);
	-- 
--	execute immediate 'select id into v_status_close_id from ' + v_table_name + ' where nm = ''да''';
	execute immediate 'select id into v_status_close_id from ' + v_table_name 
		+ ' where substring(lcase(nm), 1, 1) = char(228) and substring(lcase(nm), 2, 1) = char(224) and char_length(nm) = 2';
	--                                              'д'                                        'а'

	if v_status_close_id is not null then
		set v_trigger_sql = 
    
			'\ncreate TRIGGER wf_jscet_close before update order 212 on'
			+ '\njscet'
			+ '\nreferencing new as new_name old as old_name'
			+ '\nfor each row'
			+ '\nwhen (update (' + v_column_name + ') and old_name.' + v_column_name + ' != new_name.' + v_column_name + ')'
			+ '\nbegin'
			+ '\n	if new_name.' + v_column_name + ' = ' + convert(varchar(20), v_status_close_id) + ' then'
			+ '\n		call admin.wf_order_closed_set(old_name.id, 1);'
			+ '\n	elseif old_name.' + v_column_name + ' = ' + convert(varchar(20), v_status_close_id) + ' then'
			+ '\n		call admin.wf_order_closed_set(old_name.id, 0);'
			+ '\n	end if;'
			+ '\nend;'
		;
		execute immediate v_trigger_sql;
	end if;
end;



if exists (select '*' from sysprocedure where proc_name like 'bootstrap_blocking') then  
	drop procedure bootstrap_blocking;
end if;


create procedure bootstrap_blocking (
) 
begin
	call cre_block_var('blocks_inited');

	for v_table as b2 dynamic scroll cursor for
		select 'sdocs' as r_table union select 'sdmc'
	do 
		for v_server_name as a2 dynamic scroll cursor for
			select 
				srvname as r_server 
			from sys.sysservers s 
		do
			
			message 'call slave_cre_block_var_' + r_server + '(''' + make_block_name(admin.get_server_name(), r_table) + ''')' to log;
			execute immediate 'call slave_cre_block_var_' + r_server + '(''' + make_block_name(admin.get_server_name(), r_table) + ''')';
			--call cre_block_var(make_block_name(r_server, r_table));
		end for;
	end for;

end;
	

--****************************************************************
--              ∆урнал хоз€йственных операций
--****************************************************************


if exists (select 1 from systriggers where trigname = 'wf_xoz_insert' and tname = 'xoz' and event='INSERT') then 
	drop trigger xoz.wf_xoz_insert;
end if;


create TRIGGER wf_xoz_insert before insert order 211 on
xoz
referencing new as new_name
for each row
when (new_name.ost != 'Y')
begin

	call admin.wf_submit_xoz(
		  new_name.id_accd
		, new_name.id_accc
		, new_name.id_m_xoz
		, new_name.id_deb
		, new_name.id_jdog
		, new_name.id
		, new_name.dat
		, new_name.sum
		, new_name.sumv
		, new_name.id_curr
		, new_name.rem
		, 1                   -- искать прив€зку к заказу
	)                       
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_submit_xoz') then  
	drop procedure wf_submit_xoz;
end if;

create procedure wf_submit_xoz (
	  in p_id_accd integer
	, in p_id_accc integer
	, in p_id_m_xoz integer
	, in p_id_deb integer
	, in p_id_jdog integer
	, in p_id      integer
	, in p_dat     date
	, in p_sum     float
	, in p_sumv    float
	, in p_id_curr integer
	, in p_rem    varchar(99)
	, in p_bind_zakaz integer
) 
begin
	declare v_debit_sc varchar(26);
	declare v_debit_sub varchar(10);
	declare v_credit_sc varchar(26);
	declare v_credit_sub varchar(10);
	declare f_account_exists integer;
	declare v_nm varchar(98);
	declare v_rem varchar(98);
	declare v_purpose_id integer;
	declare v_detail_id integer;
	declare v_purpose varchar(99);
	declare v_kredDebitor integer;
	declare v_note varchar(50);

	select d.sc, d.sub_sc, d.nm, isnull(d.rem, '')
	into v_debit_sc, v_debit_sub, v_nm, v_rem
	from account d 
	where d.id = p_id_accd;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_debit_sc
		, v_debit_sub
		, v_nm
		, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	select c.sc, c.sub_sc, c.nm, c.rem
	into v_credit_sc, v_credit_sub, v_nm, v_rem
	from account c 
	where c.id = p_id_accc;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_credit_sc
		, v_credit_sub
		, v_nm, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	if (p_id_m_xoz is not null or p_id_m_xoz != 0) then
		select nm
		into v_purpose
		from m_xoz m
		where m.id = p_id_m_xoz;

		call admin.slave_set_purpose_prior (
	    	  v_purpose
	    	, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub 
	    	, v_purpose_id
		);
	end if;

	set v_kredDebitor = admin.wf_kreditor_debitor(p_id_deb);

	select nu into v_note from jdog where id = p_id_jdog;

	call admin.slave_put_xoz_prior(
		  admin.get_server_name() 
		, p_id
		, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub
		, convert(varchar(20), p_dat, 115)
		, p_sum
		, p_sumv
		, p_id_curr
		, p_rem
		, v_purpose_id
		, v_kredDebitor
		, v_note
		, p_bind_zakaz
	);
end;
		



if exists (select '*' from sysprocedure where proc_name like 'wf_kreditor_debitor') then  
	drop function wf_kreditor_debitor;
end if;

create function wf_kreditor_debitor (
	p_id_deb integer
) returns integer
begin
	declare v_kredDebitor integer;
	declare varchar_id_deb varchar(20);
	declare v_values varchar(1024);
	declare v_deb_name varchar(203);

	
	select nm into v_deb_name from voc_names where id = p_id_deb;
	call admin.slave_select_prior(varchar_id_deb, 'GuideFirms', 'min(firmid)', 'Name='''+ v_deb_name +'''');
	set v_kredDebitor = convert(integer, varchar_id_deb);

	if v_kredDebitor is null then
		call admin.slave_select_prior(varchar_id_deb, 'yDebKreditor', 'min(id)', 'Name='''+ v_deb_name +'''');
		set v_kredDebitor = convert(integer, varchar_id_deb);
	end if;

	if v_kredDebitor is null then
		call admin.slave_select_prior(varchar_id_deb, 'yDebKreditor', 'min(id)', '1=1');
		
		set v_kredDebitor = convert(integer, isnull(varchar_id_deb, '')) - 1;
		set v_values = '''' + convert(varchar(20), v_kredDebitor) + ''''
			+ ', ''' + v_deb_name + ''''
			+ ', '''  + admin.get_server_name() + ''''
		;
		call admin.slave_insert_prior('yDebKreditor', 'id, name, note', v_values);
	end if;

	return v_kredDebitor;
end;

-------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_xoz_update' and tname = 'xoz') then 
	drop trigger xoz.wf_xoz_update;
end if;

create TRIGGER wf_xoz_update before update order 211 on
xoz
referencing old as old_name new as new_name
for each row
begin
	declare v_debit_sc varchar(26);
	declare v_debit_sub varchar(10);
	declare v_credit_sc varchar(26);
	declare v_credit_sub varchar(10);
	declare v_nm varchar(98);
	declare v_rem varchar(98);
	declare f_account_exists integer;
	declare v_currency_rate varchar(20);
	declare v_currency float;
	declare v_m_xoz varchar(100);
	declare v_kredDebitor integer;
--	declare varchar_id varchar(20);
	declare v_values varchar(1024);
	declare v_ventureid char(20);

	declare v_note varchar(10);
	declare v_zakaz varchar(150);
	declare s_id_shiz varchar(20);
	declare v_id_m_xoz integer;
	declare m_xoz_updated integer;

	set m_xoz_updated = 0;

	set v_ventureid = admin.select_remote('prior', 'guideventure', 'ventureid', 'sysname = ''''' + admin.get_server_name() + '''''');

	if update(dat) then
		call admin.update_remote('prior', 'ybook', 'xDate', '''''' + convert(varchar(20), new_name.dat, 115) + ''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
    		
	end if;


	if update(id_jdog) then
		select nu into v_note from jdog where id = new_name.id_jdog;

		if char_length(v_note) > 0 then
		
			call admin.update_remote('prior', 'ybook', 'note', '''''' + v_note + ''''''
				, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
			);
    		
			select c.sc
			into v_credit_sc
			from account c 
			where c.id = new_name.id_accc;

			call admin.slave_bind_zakaz_prior (
				v_zakaz
				, admin.get_server_name()
				, v_note
				, old_name.sum
				, v_credit_sc
				, old_name.id
			);
		end if;

	end if;

	if update(id_deb) then
		set v_kredDebitor = admin.wf_kreditor_debitor(new_name.id_deb);

		call admin.update_remote('prior', 'ybook', 'KredDebitor', convert(varchar(20), v_kredDebitor)
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
    end if;

	if update(id_accd) and isnull(new_name.id_accd, 0) != 0 then
		    
		select d.sc, d.sub_sc, d.nm, d.rem
		into v_debit_sc, v_debit_sub, v_nm, v_rem
		from account d 
		where d.id = new_name.id_accd;

	    
		call admin.slave_put_account_prior(
			f_account_exists
			, v_debit_sc
			, v_debit_sub
			, v_nm, v_rem
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'purposeId'
			, 'null'
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'Debit'
			, '''''' + v_debit_sc + ''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'subDebit'
			, '''''' + v_debit_sub + ''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
    
		set m_xoz_updated = 1;
	end if;

	if update(id_accc) and isnull(new_name.id_accc, 0) != 0  then

		select c.sc, c.sub_sc, c.nm, c.rem
		into v_credit_sc, v_credit_sub, v_nm, v_rem
		from account c 
		where c.id = new_name.id_accc;
	    
		call admin.slave_put_account_prior(
			f_account_exists
			, v_credit_sc
			, v_credit_sub
			, v_nm, v_rem
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'purposeId'
			, 'null'
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'Kredit'
			, ''''''+v_credit_sc+''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'subKredit'
			, ''''''+v_credit_sub+''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
		set m_xoz_updated = 1;

	end if;

	if update(sum) then
		call admin.slave_select_prior(
			v_currency_rate
			,'system'
			,'Kurs'
			,'1=1'
		);

		set v_currency = new_name.sum / convert(float, abs(v_currency_rate));

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'UEsumm'
			, v_currency
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

	end if;


	if (update(id_m_xoz) or isnull(new_name.id_m_xoz, 0) != 0) 
			or m_xoz_updated = 1 
	then

		select d.sc, d.sub_sc, d.nm, d.rem
		into v_debit_sc, v_debit_sub, v_nm, v_rem
		from account d 
		where d.id = new_name.id_accd;
	    
		select c.sc, c.sub_sc, c.nm, c.rem
		into v_credit_sc, v_credit_sub, v_nm, v_rem
		from account c 
		where c.id = new_name.id_accc;

		call admin.wf_purpose_sync (
			old_name.id
			, v_ventureid
			, new_name.id_m_xoz
			, v_debit_sc
			, v_debit_sub
			, v_credit_sc
			, v_credit_sub
		);
		    

	end if;


	if update(rem) then

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'descript'
			, ''''''+new_name.rem+''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

	end if;

	if update(id_sh_zatrat) then
		if isnull(new_name.id_sh_zatrat, 0) = 0 then 
			set s_id_shiz = 'null';
		else 
			set s_id_shiz = convert(varchar(20), new_name.id_sh_zatrat);
		end if;

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'id_shiz'
			, s_id_shiz
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
	end if;

end;


if exists (select 1 from systriggers where trigname = 'wf_xoz_delete' and tname = 'xoz') then 
	drop trigger xoz.wf_xoz_delete;
end if;

create TRIGGER wf_xoz_delete before delete order 211 on
xoz
referencing old as old_name
for each row
begin
	call admin.slave_delete_prior(
		 'ybook y'
		, 'v.sysname = '''
				+ admin.get_server_name() 
				+ ''' and v.ventureid = y.ventureid and y.id_xoz = '
				+ convert(varchar(20), old_name.id)
		, 'GuideVenture v'
	);
end;

---------------------------------------------------------------
-------------     jscet      ----------------------------------
---------------------------------------------------------------

if exists (select 1 from systriggers where trigname = 'wf_jscet_update' and tname = 'jscet') then 
	drop trigger jscet.wf_jscet_update;
end if;

create TRIGGER wf_jscet_update before update order 211 on
/*
	¬ыставл€ет плательщика в базе Prior
*/
jscet
referencing old as old_name new as new_name
for each row
begin
	declare no_echo integer;
	declare v_is_orders varchar(10);
	
	set no_echo = 0;

  	begin
		select @prior_jscet into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

--	message 'TRIGGER wf_jscet_update:: no_echo = ', no_echo to client;
	
	if update(id_d) then
		-- вы€снить это продажи или заказ
		set v_is_orders = admin.select_remote('prior'
			,'orders'
			,'count(*)'
			,'id_jscet = ' + convert(varchar(20), old_name.id)
		);

--		message 'v_is_orders = ', v_is_orders to client;

		if v_is_orders = '0' then
			-- нет такого счета в «аказах
			call admin.update_remote (
				'prior'
				, 'bayOrders'
				, 'id_bill'
				, convert(varchar(20), new_name.id_d)
				, 'id_jscet = ' + convert(varchar(20), old_name.id)
			);
		else
			call admin.update_remote (
				'prior'
				, 'orders'
				, 'id_bill'
				, convert(varchar(20), new_name.id_d)
				, 'id_jscet = ' + convert(varchar(20), old_name.id)
			);
		end if;
	end if;

end;


if exists (select '*' from sysprocedure where proc_name like 'wf_purpose_sync') then  
	drop procedure wf_purpose_sync;
end if;

create procedure wf_purpose_sync (
	in p_id_xoz integer
	, p_ventureid  varchar(20)
	, in p_id_m_xoz   integer
	, p_debit_sc   varchar(26)
	, p_debit_sub  varchar(10)
	, p_credit_sc  varchar(26)
	, p_credit_sub varchar(10)
)
begin
	declare v_m_xoz      varchar(99);
	declare v_purposeid integer;

	select m.nm into v_m_xoz 
	from m_xoz m
	where m.id = p_id_m_xoz;

	if isnull(v_m_xoz, '') != '' then
		call admin.slave_set_purpose_prior(
			 v_m_xoz
			, p_debit_sc, p_debit_sub, p_credit_sc, p_credit_sub
			, v_purposeid
		);
		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'purposeId'
			, v_purposeid
			, 'id_xoz=' + convert(varchar(20), p_id_xoz) + ' and ventureid = ' + p_ventureid
		);
	end if;
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_put_xoz') then  
	drop procedure wf_put_xoz;
end if;

create procedure wf_put_xoz (
	  new_name_id         integer
	, new_name_dat        timestamp
	, new_name_sum        float
	, new_name_sumv       float
	, new_name_id_curr    integer
	, new_name_id_accd    integer
	, new_name_id_accc    integer
	, new_name_id_deb     integer
	, new_name_id_jdog    integer
	, new_name_id_m_xoz   integer
	, new_name_rem        varchar(99)
) 
begin
	declare v_debit_sc varchar(26);
	declare v_debit_sub varchar(10);
	declare v_credit_sc varchar(26);
	declare v_credit_sub varchar(10);
	declare f_account_exists integer;
	declare v_nm varchar(98);
	declare v_rem varchar(98);
	declare v_purpose_id integer;
	declare v_detail_id integer;
	declare v_purpose varchar(99);
	declare v_kredDebitor integer;
	declare v_note varchar(50);

	select d.sc, d.sub_sc, d.nm, isnull(d.rem, '')
	into v_debit_sc, v_debit_sub, v_nm, v_rem
	from account d 
	where d.id = new_name_id_accd;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_debit_sc
		, v_debit_sub
		, v_nm
		, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	select c.sc, c.sub_sc, c.nm, c.rem
	into v_credit_sc, v_credit_sub, v_nm, v_rem
	from account c 
	where c.id = new_name_id_accc;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_credit_sc
		, v_credit_sub
		, v_nm, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	if (new_name_id_m_xoz is not null or new_name_id_m_xoz != 0) then
		select nm
		into v_purpose
		from m_xoz m
		where m.id = new_name_id_m_xoz;

		call admin.slave_set_purpose_prior (
	    	  v_purpose
	    	, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub 
	    	, v_purpose_id
		);
	end if;

	set v_kredDebitor = admin.wf_kreditor_debitor(new_name_id_deb);

	select nu into v_note from jdog where id = new_name_id_jdog;

	call admin.slave_put_xoz_prior(
		  admin.get_server_name() 
		, new_name_id
		, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub
		, convert(varchar(20), new_name_dat, 115)
		, new_name_sum
		, new_name_sumv
		, new_name_id_curr
		, new_name_rem
		, v_purpose_id
		, v_kredDebitor
		, v_note
	);
		
end;

/*
exit;


begin
    declare v_not_load varchar(20);
	
	for aCursor as b dynamic scroll cursor for
		select 
			  id       as r_id         
			, dat      as r_dat        
			, sum      as r_sum        
			, sumv     as r_sumv       
			, id_curr  as r_id_curr    
			, id_accd  as r_id_accd    
			, id_accc  as r_id_accc    
			, id_deb   as r_id_deb     
			, id_jdog  as r_id_jdog    
			, id_m_xoz as r_id_m_xoz   
			, rem      as r_rem        
		from xoz 
		where convert(varchar(8), now(), 112) = convert(varchar(8), dat, 112)
	do
		set v_not_load = select_remote('prior', 'ybook', 'count(*)', 'id_xoz = ' + convert(varchar(20), r_id));

		if v_not_load = '0' then
		message '*** load xoz operation for the id = ', r_id to client;
			call 
				 wf_put_xoz (
					  r_id         
					, r_dat        
					, r_sum        
					, r_sumv       
					, r_id_curr    
					, r_id_accd    
					, r_id_accc    
					, r_id_deb     
					, r_id_jdog    
					, r_id_m_xoz   
					, r_rem        
				);
		end if;

	end for;
end;

*/