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
		, '	out o_date char(20)'
		+ '	,out o_rate real'
		+ '	,in p_date char(20) default null'
		+ ' ,in p_id_cur integer default null'
);


call build_host_procedure (
		 'nextid'
		, '  IN table_name char(100)'
		  +',out id int'
);


call build_host_procedure (
		 'nextnu'
		, '  IN table_name char(100)'
		+ ', out p_nu char(32)'
		+ ', in p_dat_field char(32) default null'
);



if exists (select '*' from sysprocedure where proc_name like 'nextnu_remote') then
	drop function nextnu_remote;
end if;

create function nextnu_remote(
	  p_server_name varchar(32)
	, p_table_name varchar(32)
	, p_dat_field char(32) default null
) returns varchar(32)
begin
	declare v_sql varchar(1000);
	set v_sql = 
	 'call slave_nextnu_' + p_server_name 
		+'(''' + p_table_name + ''''
		+ ', nextnu_remote' 
	;
	if p_dat_field is not null then
		set v_sql = v_sql
			+ ', ''' + p_dat_field + ''''
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

