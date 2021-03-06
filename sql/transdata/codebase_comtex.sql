

-------------- ������� ����� �������� -------------------

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

--------------> ������� ����� �������� <-------------------





/**
 get_server_name() => @server_name 
 ��������� ������ ���������� ���� ��� �� bootstrap_blocking.
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
	

	-- ����� ���������������� ���������� � ������� � ������� �������
	-- �������� ���-�� ���� ����� 'GUIDE_803_129574.NM','JSCET__USER_129573'
	select nm, parent_col_name
	into v_table_name, v_column_name
	from browsers where id_guides = 1005 
	and nm like '%guid%' 
	and namer like '%���%';

	if v_table_name is null then 
		return;
	end if;
	-- ������� ��  'GUIDE_803','USER_129573'
	set v_table_name = 'GUIDE_' + substring(v_table_name, 7, charindex('_', substring(v_table_name, 7))-1);
	set v_column_name =  substring(v_column_name, charindex('__', v_column_name)+2);
	-- 
--	execute immediate 'select id into v_status_close_id from ' + v_table_name + ' where nm = ''��''';
	execute immediate 'select id into v_status_close_id from ' + v_table_name 
		+ ' where substring(lcase(nm), 1, 1) = char(228) and substring(lcase(nm), 2, 1) = char(224) and char_length(nm) = 2';
	--                                              '�'                                        '�'

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
--              ������ ������������� ��������
--****************************************************************



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