--Создание remote серверов

if not exists (select '*' from sys.sysservers where srvname = 'stime') then  
	create server stime class 'ASAODBC' USING 'DSN=stime;UID=admin;PWD=z';
end if;
if not exists (select '*' from sys.sysservers where srvname = 'accountN') then  
	create server accountN class 'ASAODBC' USING 'DSN=accountN;UID=admin;PWD=z';
end if;
if not exists (select '*' from sys.sysservers where srvname = 'markmaster') then  
	create server markmaster class 'ASAODBC' USING 'DSN=markmaster;UID=admin;PWD=z';
end if;


call build_host_procedure (
	  'currency_rate'
	,   ' out o_date char(20)'
	  + ',out o_rate real'
	  + ',in p_date char(20) default null'
	  + ',in p_id_cur integer default null'
);


call build_host_procedure (
	 'nextid'
	, '  IN table_name char(100)'
	  +',out id int'
);

call build_host_procedure (
	 'legacy_purpose'
	, '   IN purpose_name char(100)'
	  + ',in debit char(26)'
	  + ',in subdebit char(10)'
	  + ',in kredit char(26)'
	  + ',in subkredit char(10)'
);



call build_host_procedure (
	'renu_scet'
	,'in p_id_jscet integer'
);

call build_host_procedure (
	'move_uslug'
	,   'in p_id_jscet integer'
	  + ', in p_id_jscet_new integer'
	  + ', in p_quant real'
	  + ', in p_id_inv integer'
);



call build_host_procedure (
	 'nextnu'
	, '  in table_name char(100)'
	  + ', out p_nu char(32)'
	  + ', in p_nu_old char(32) default null'
	  + ', in p_dat_field char(32) default null'
	  + ', in p_dat char(32) default null'
);


-- для перевода накладных в режим простых или импортных
call build_remote_host (
	 'change_id_guide'
	, 'in p_id_jmat integer'
	+ ', in p_id_guide integer'
	+ ', in p_id_currency integer'
	+ ', in p_tp1 integer'
	+ ', in p_tp2 integer'
	+ ', in p_tp3 integer'
	+ ', in p_tp4 integer'
);

-- при изменени количества в приходной накладной Prior(stime)
call build_remote_host (
	 'change_mat_qty'
	,   'in p_id_mat integer'
	  + ', in p_new_quant float'
);

call build_rp_procedure (
	  'stime'
	, 'wf_calc_cost'
	,   'out out_ret float'
	  + ', p_id_inv integer'
);





if exists (select '*' from sysprocedure where proc_name like 'nextnu_remote') then
	drop function nextnu_remote;
end if;

create function nextnu_remote(
	  in p_server_name varchar(32)
	, in p_table_name varchar(32)
	, in p_nu_old varchar(32) default null
	, in p_dat_field char(32) default null
	, in p_dat char(32) default null
) returns varchar(32)
begin
	declare v_sql varchar(1000);
	set v_sql = 
	 'call slave_nextnu_' + p_server_name 
		+'(''' + p_table_name + ''''
		+ ', nextnu_remote' 
	;

	if p_nu_old is not null then
		set v_sql = v_sql
			+ ', ''' + p_nu_old + ''''
		;
	else 
		set v_sql = v_sql
			+ ', null'
		;
	end if;

	if p_dat_field is not null then
		set v_sql = v_sql
			+ ', ''' + p_dat_field + ''''
		;
	else 
		set v_sql = v_sql
			+ ', null'
		;
	end if;

	if p_dat is not null then
		set v_sql = v_sql
			+ ', ''' + p_dat + ''''
		;
	else 
		set v_sql = v_sql
			+ ', null'
		;
	end if;

	set v_sql = v_sql
		+ ')'
	;
	message v_sql to client;

	execute immediate v_sql;
end;


