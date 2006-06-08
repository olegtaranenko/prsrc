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
			
			message 'call slave_cre_block_var_' + r_server + '(''' + make_block_name(@@servername, r_table) + ''')' to log;
			execute immediate 'call slave_cre_block_var_' + r_server + '(''' + make_block_name(@@servername, r_table) + ''')';
			--call cre_block_var(make_block_name(r_server, r_table));
		end for;
	end for;

end;
	

--****************************************************************
--                               ID CORRECTION
-- � ������ 8.1.5 ����� ����� ���������� id, ������� ����� 
-- ������������ ��� ����������� ���������� ����� � �������
-- ��� ������������� ����� �������� ����� ������������ ��������,
-- ������� ����� ������� ��� ���� ������, ������ �� ������� ����� � inc_table
--****************************************************************

if exists (select '*' from sysprocedure where proc_name like 'build_id_track_trigger') then  
	drop procedure build_id_track_trigger;
end if;

create procedure build_id_track_trigger (
	p_table_name varchar(64)
)
begin
	declare sqls varchar(2000);

	set sqls = 
		'if exists (select 1 from systriggers where trigname = ''wf_correct_id_taid'' and tname = ''' + p_table_name + ''') then '
		+ '\n	drop trigger ' + p_table_name + '.wf_correct_id_taid;'
		+ '\n end if;'
		+ '\n'
		+ '\n create trigger wf_correct_id_taid after insert, delete order 100'
		+ '\n on ' + p_table_name + ' '
		+ '\n referencing old as old_name new as new_name'
		+ '\nbegin'
		+ '\n	declare idd integer;'
		+ '\n	select isnull(max(id), 0) + 1 into idd from ' + p_table_name + ';'
		+ '\n	update inc_table set next_id = idd where table_nm = ''' + p_table_name + ''';'
		+ '\nend;';
	execute immediate sqls;
end;



--****************************************************************
--              ������ ������������� ��������
--****************************************************************


if exists (select 1 from systriggers where trigname = 'wf_xoz_insert' and tname = 'xoz' and event='INSERT') then 
	drop trigger xoz.wf_xoz_insert;
end if;


create TRIGGER wf_xoz_insert before insert order 211 on
xoz
referencing new as new_name
for each row
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
	where d.id = new_name.id_accd;


//	message 'd.sc = '+v_debit_sc to client;
//	message 'd.sub_sc = '+v_debit_sub to client;
//	message 'nm = '+v_nm to client;
//	message 'rem = '+v_rem to client;

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
	where c.id = new_name.id_accc;

//	message 'c.sc = '+v_credit_sc to client;
//	message 'c.sub_sc = '+v_credit_sub to client;
//	message 'nm = '+v_nm to client;
//	message 'rem = '+v_rem to client;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_credit_sc
		, v_credit_sub
		, v_nm, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	if (new_name.id_m_xoz is not null or new_name.id_m_xoz != 0) then
		select nm
		into v_purpose
		from m_xoz m
		where m.id = new_name.id_m_xoz;

		call admin.slave_set_purpose_prior (
	    	  v_purpose
	    	, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub 
	    	, v_purpose_id
		);
	end if;

	set v_kredDebitor = admin.wf_kreditor_debitor(new_name.id_deb);

	select nu into v_note from jdog where id = new_name.id_jdog;

	call admin.slave_put_xoz_prior(
		  @@servername 
		, new_name.id
		, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub
		, convert(varchar(20), new_name.dat, 115)
		, new_name.sum
		, new_name.sumv
		, new_name.id_curr
		, new_name.rem
		, v_purpose_id
		, v_kredDebitor
		, v_note
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
	declare v_deb_name varchar(128);

	
	select nm into v_deb_name from voc_names where id = p_id_deb;
	call admin.slave_select_prior(varchar_id_deb, 'GuideFirms', 'firmid', 'Name='''+ v_deb_name +'''');
	set v_kredDebitor = convert(integer, varchar_id_deb);

	if v_kredDebitor is null then
		call admin.slave_select_prior(varchar_id_deb, 'yDebKreditor', 'id', 'Name='''+ v_deb_name +'''');
		set v_kredDebitor = convert(integer, varchar_id_deb);
	end if;

	if v_kredDebitor is null then
		call admin.slave_select_prior(varchar_id_deb, 'yDebKreditor', 'min(id)', '1=1');
		
		set v_kredDebitor = convert(integer, isnull(varchar_id_deb, '')) - 1;
		set v_values = '''' + convert(varchar(20), v_kredDebitor) + ''''
			+ ', ''' + v_deb_name + ''''
			+ ', '''  + @@servername + ''''
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
	declare v_purposeid integer;
	declare v_kredDebitor integer;
	declare varchar_id varchar(20);
	declare v_values varchar(1024);
	declare v_ventureid integer;
	declare v_note varchar(10);
	declare v_zakaz varchar(150);

	if update(dat) then
		call admin.slave_select_prior(varchar_id, 'guideventure', 'ventureid', 'sysname=''' + @@servername + '''');
		set v_ventureid = convert(integer, varchar_id);
	
		call admin.slave_update_prior('ybook', 'xDate', '''' + convert(varchar(20), new_name.dat, 115) + ''''
			, 'id_xoz=' + convert(varchar(20), old_name.id )
			  + ' and ventureid = ' + convert(varchar(20), v_ventureid)
		);
    		
	end if;


	if update(id_jdog) then
		select nu into v_note from jdog where id = new_name.id_jdog;

		if char_length(v_note) > 0 then
		
			select c.sc
			into v_credit_sc
			from account c 
			where c.id = new_name.id_accc;


			call admin.slave_select_prior(varchar_id, 'guideventure', 'ventureid', 'sysname=''' + @@servername + '''');
			set v_ventureid = convert(integer, varchar_id);
	    
			call admin.slave_update_prior('ybook', 'note', '''' + v_note + ''''
				, 'id_xoz=' + convert(varchar(20), old_name.id )
				  + ' and ventureid = ' + convert(varchar(20), v_ventureid)
			);
    		
			call admin.slave_bind_zakaz_prior (
				v_zakaz
				, @@servername
				, v_note
				, old_name.sum
				, v_credit_sc
				, old_name.id
			);
		end if;

	end if;

	if update(id_deb) then
		set v_kredDebitor = admin.wf_kreditor_debitor(new_name.id_deb);

		call admin.slave_select_prior(varchar_id, 'guideventure', 'ventureid', 'sysname=''' + @@servername + '''');
		set v_ventureid = convert(integer, varchar_id);
		call admin.slave_update_prior('ybook', 'KredDebitor', convert(varchar(20), v_kredDebitor)
			, 'id_xoz=' + convert(varchar(20), old_name.id )
			  + ' and ventureid = ' + convert(varchar(20), v_ventureid)
		);
    end if;

	if update(id_accd) and new_name.id_accd != 0 then
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
		if f_account_exists = 0 then
			--return;
		end if;

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'Debit'
			, ''''''+v_debit_sc+''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'subDebit'
			, '''''' + v_debit_sub + ''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);
	end if;

	if update(id_accc) and new_name.id_accc != 0 then
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
		if f_account_exists = 0 then
			--return;
		end if;

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'Kredit'
			, ''''''+v_credit_sc+''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'subKredit'
			, ''''''+v_credit_sub+''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);
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
			, 'ybook y'
			, 'UEsumm'
			, v_currency
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);

	end if;


	if update(id_m_xoz) and new_name.id_m_xoz!= 0 then
		select d.sc, d.sub_sc, d.nm, d.rem
		into v_debit_sc, v_debit_sub, v_nm, v_rem
		from account d 
		where d.id = old_name.id_accd;
	    
		select c.sc, c.sub_sc, c.nm, c.rem
		into v_credit_sc, v_credit_sub, v_nm, v_rem
		from account c 
		where c.id = old_name.id_accc;

		select nm into v_m_xoz from m_xoz where id = new_name.id_m_xoz;

		if v_debit_sc is not null and v_credit_sc is not null then
			call admin.slave_set_purpose_prior(
				 v_m_xoz
				, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub
				, v_purposeid
			);
			//message 'v_purposeid = ' + convert(varchar(20), v_purposeid) to client;	
			call admin.update_remote(
				'prior'
				, 'ybook y'
				, 'purposeId'
				, v_purposeid
				, 'v.sysname = '''''
						+ @@servername 
						+''''' and v.ventureid = y.ventureid and y.id_xoz = '
						+ convert(varchar(20), old_name.id)
				, 'GuideVenture v'
			);
		end if;

	end if;
	if update(rem) then
		select d.sc, d.sub_sc, d.nm, d.rem
		into v_debit_sc, v_debit_sub, v_nm, v_rem
		from account d 
		where d.id = old_name.id_accd;
	    
		select c.sc, c.sub_sc, c.nm, c.rem
		into v_credit_sc, v_credit_sub, v_nm, v_rem
		from account c 
		where c.id = old_name.id_accc;

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'descript'
			, ''''''+new_name.rem+''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
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
				+ @@servername 
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
	���������� ����������� � ���� Prior
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

	message 'TRIGGER wf_jscet_update:: no_echo = ', no_echo to client;
	
	if update(id_d) then
		-- �������� ��� ������� ��� �����
		set v_is_orders = admin.select_remote('prior'
			,'orders'
			,'count(*)'
			,'id_jscet = ' + convert(varchar(20), old_name.id)
		);

		message 'v_is_orders = ', v_is_orders to client;

		if v_is_orders = '0' then
			-- ��� ������ ����� � �������
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
/*
	if update(nu) then
		call admin.sync_next_nu(new_name.nu, old_name.nu);
	end if;
*/
end;


if exists (select 1 from systriggers where trigname = 'wf_jscet_insert' and tname = 'jscet' and event='INSERT') then 
	drop trigger jscet.wf_jscet_insert;
end if;

/*

create TRIGGER wf_jscet_insert before insert order 211 on
jscet
referencing new as new_name
for each row
begin
	call admin.sync_next_nu(new_name.nu);
end;
*/

if exists (select '*' from sysprocedure where proc_name like 'sync_next_nu') then  
	drop procedure sync_next_nu;
end if;

/*
create procedure sync_next_nu (
	  in p_nu_ch varchar(50)
	, in p_nu_old_ch varchar(50) default ''
) 
begin
	declare v_nu_prior_ch varchar(50);
	declare v_nu_prior integer;
	declare v_nu_old integer;
	declare v_nu_old_ceil integer;

	declare v_nu_update integer;

	declare v_last_inc integer;
	declare v_downgrade integer;

	if isnumeric(p_nu_ch) = 1 then

		set v_last_inc = 1;
		set v_nu_prior_ch = select_remote(
			'prior'
			,'guideVenture'
			,'intInvoice'
			, 'sysname = ''''' + @@servername + ''''''
		);
		-- �����, ������� ������� ��������� � ������� guideVenture
		set v_nu_prior = convert(integer, v_nu_prior_ch);
		message '0) v_nu_prior = ', v_nu_prior to client;

		-- ����� �����, ������� ������ ���� �������� �����
		-- ����� ��������� ������
		set v_nu_update = convert(integer, p_nu_ch);
		message '1) v_nu_update = ',v_nu_update  to client;

		-- ���������, ����� ��� ���� ����� ������?
		if isnumeric(p_nu_old_ch) = 1 then
			set v_nu_old = convert(integer, p_nu_old_ch);
			message '3) v_nu_old = ',v_nu_old  to client;

            -- ���� ������������ ��������, ������� ��������
            -- � ����, ����� ���� ��� p_nu_old_ch ����� �������
			select isnull(max(convert(integer, nu)), 0)
			into v_nu_old_ceil
			from jscet
			where isnumeric(nu) = 1
				and convert(varchar(4), dat, 112) = convert(varchar(4), now(), 112)
				and nu != p_nu_old_ch;
			message '4) v_nu_old_ceil = ',v_nu_old_ceil  to client;

		end if;

		if v_nu_old is null then
			-- ������� ��������� ������� ���������� 
		else
			set v_downgrade = 0;
			    
		end if;

		if v_nu_update >= v_nu_prior or v_downgrade = 1 then
			call slave_update_prior('guideVenture', 'intInvoice', v_nu_update + v_last_inc, 'sysname = ''' + @@servername + '''');
		end if;


	end if;

end;
*/



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


//	message 'd.sc = '+v_debit_sc to client;
//	message 'd.sub_sc = '+v_debit_sub to client;
//	message 'nm = '+v_nm to client;
//	message 'rem = '+v_rem to client;

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

//	message 'c.sc = '+v_credit_sc to client;
//	message 'c.sub_sc = '+v_credit_sub to client;
//	message 'nm = '+v_nm to client;
//	message 'rem = '+v_rem to client;


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
		  @@servername 
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